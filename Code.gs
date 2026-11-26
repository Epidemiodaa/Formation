/**
 * PLATEFORME DE FORMATION (type MOOC) — API Google Apps Script
 * - Inscription libre : pseudonyme + mot de passe.
 * - L'administrateur crée des cours ; chaque cours a des documents, un pré-test, un post-test et une discussion.
 * - Tout est dans le Google Sheet (un onglet par table) : modifiable à la main.
 * - L'utilisateur peut modifier son pseudo / mot de passe et ses propres messages.
 * Déploiement : Déployer > Application Web > Exécuter en tant que moi, accès : tout le monde.
 * Exécuter setup() une fois pour créer les onglets et autoriser Drive.
 */
const ADMIN_PASSWORD = 'Silo@2000'; // mot de passe de l'espace administrateur
const SHEET_ID = '1hZ7v3_5K2qhD5N4YRmt3C96OsxOIk0iuQdQUwYSrwQc';
const TABS = {
  // Users : l'administrateur voit et modifie pseudo et motdepasse directement ici (id = identifiant stable, ne pas modifier)
  Users: ['id', 'pseudo', 'motdepasse', 'jeton', 'cree'],
  // pretest : « obligatoire » (défaut, vide) ou « facultatif » (cours accessible sans pré-test) ; seuil : note de réussite des quiz en % (vide = 70, 0 = aucune exigence)
  // cree = date de création (pour signaler un nouveau cours) ; image = adresse de l'image de couverture
  // mode : « modules » (séquences : documents puis quiz, introduction animée) ; vide ou « simple » = cours d'une seule présentation, sans modules (comportement d'origine)
  Cours: ['id', 'titre', 'description', 'actif', 'pretest', 'seuil', 'cree', 'image', 'mode'],
  // facultatif = TRUE : ressource complémentaire, consultable par tous les inscrits mais sans obligation de lecture
  // sequence = id de la séquence ; vide = document hors séquence
  // type : video, word, pdf, ppt, excel, image, page (leçon HTML stockée dans Drive, url = « html:ID »), lien
  Documents: ['id', 'cours', 'titre', 'url', 'facultatif', 'sequence', 'type'],
  // options = JSON ["A","B"] ; bonnes = JSON [0] (index à partir de 0, plusieurs = choix multiple)
  // type : single | multi | chart | rpt | dates | match ; donnees = JSON (contexte, tableaux, dates, appariements)
  Questions: ['id', 'cours', 'phase', 'type', 'question', 'options', 'bonnes', 'donnees', 'explication', 'sequence'], // phase « Quiz » : sequence = séquence concernée
  Blocs: ['id', 'cours', 'titre', 'texte'],
  Inscriptions: ['uid', 'pseudo', 'cours', 'date'],
  // details = JSON de la correction (réponse de l'utilisateur, bonnes réponses) ; sequence : pour les quiz
  Resultats: ['date', 'uid', 'pseudo', 'cours', 'phase', 'score', 'total', 'sequence', 'details'],
  Messages: ['id', 'date', 'cours', 'uid', 'pseudo', 'texte'],
  // Lectures : une ligne par utilisateur et par document ; ouvert = date d'ouverture du lien, lu = date où l'utilisateur l'a marqué « lu en entier »
  Lectures: ['uid', 'cours', 'doc', 'ouvert', 'lu'],
  // Sequences : un cours = suite ordonnée de séquences (documents puis quiz). ordre = 1, 2, 3…
  // intro = texte de présentation dit par les personnages animés (une réplique par ligne, « Awa: » ou « Kofi: » pour choisir qui parle)
  Sequences: ['id', 'cours', 'titre', 'ordre', 'intro', 'image']
};
const PHASES = ['Pré-test', 'Post-test'];
const PHASES_ALL = PHASES.concat('Quiz'); // Quiz = quiz de fin de séquence

const API = { login, me, profil, updateProfile, enroll, unenroll, course, openDoc, markDoc, startTest, submitTest,
  getMessages, postMessage, editMessage, deleteMessage,
  adminLogin, adminOverview, adminSaveCourse, adminAddDoc, adminAddQuestion,
  adminCourseDetail, adminDelete, adminAddBloc, adminImporterMPE: adminImporterMPE_, adminDeleteCourse, adminUpdate, adminImportQuestions, adminSaveSequence, adminMoveSequence, adminDeleteSequence, adminSetCourseImage, adminSetModuleImage, adminUploadImage, adminAddPage, adminGetPage, adminSavePage, getPage, notifs };

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (!API.hasOwnProperty(b.action)) return out_({ ok: false, message: 'Action inconnue.' });
    return out_(API[b.action].apply(null, b.args || []));
  } catch (err) { return out_({ ok: false, message: err.message }); }
}
function doGet() { return out_({ ok: true, service: 'plateforme-formation' }); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function setup() { Object.keys(TABS).forEach(sh_); txt_(); DriveApp.getRootFolder(); }
/** Colonnes pseudo / mot de passe en texte brut : « 0123 » ne perd pas son zéro. */
function txt_() { sh_('Users').getRange('A:D').setNumberFormat('@'); }

/* ---------- Accès au Sheet ---------- */
function sh_(n) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(TABS[n]); s.setFrozenRows(1); if (n === 'Users') s.getRange('A:D').setNumberFormat('@'); }
  else { const h = hdr_(s); TABS[n].forEach(k => { if (h.indexOf(k) < 0) { s.getRange(1, h.length + 1).setValue(k); h.push(k); } }); }
  return s;
}
function hdr_(s) { return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; }
function rows_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.map((r, i) => { const o = { _r: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function add_(n, o) { if (n === 'Cours' && o.cree === undefined) o.cree = new Date(); const s = sh_(n); s.appendRow(hdr_(s).map(k => o[k] === undefined ? '' : o[k])); }
function set_(n, r, k, v) { const s = sh_(n); s.getRange(r, hdr_(s).indexOf(k) + 1).setValue(v); }
const S = v => String(v);
const id_ = p => p + Utilities.getUuid().slice(0, 8);

/* ---------- Comptes ---------- */
function findUser_(p) { return rows_('Users').filter(u => S(u.pseudo).toLowerCase() === S(p).trim().toLowerCase())[0]; }
function auth_(tok) {
  const u = tok && rows_('Users').filter(x => S(x.jeton) === S(tok))[0];
  if (!u) throw new Error('Session expirée : reconnectez-vous.');
  return u;
}
// Critères usuels : pseudonyme de 3 à 30 caractères (lettres, chiffres, . _ -), sans espace ; mot de passe de 6 caractères minimum
const PSEUDO_RE = /^[\p{L}\p{N}._-]{3,30}$/u;
function checkPseudo_(p) { return PSEUDO_RE.test(p) ? '' : 'Pseudonyme : 3 à 30 caractères (lettres, chiffres, point, tiret, souligné), sans espace.'; }
function checkPw_(pw) { return S(pw).length >= 6 ? '' : 'Mot de passe : 6 caractères minimum.'; }
/** Première utilisation : un pseudonyme inconnu crée le compte tout de suite (aucune validation de l'administrateur). */
function login(pseudo, pw) {
  pseudo = S(pseudo).trim(); pw = S(pw);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const u = findUser_(pseudo);
    if (u) {
      if (S(u.motdepasse) !== pw) return { ok: false, message: 'Mot de passe incorrect.' };
      if (!u.jeton) { u.jeton = 't' + Utilities.getUuid(); set_('Users', u._r, 'jeton', u.jeton); }
      return { ok: true, token: S(u.jeton), pseudo: S(u.pseudo) };
    }
    const err = checkPseudo_(pseudo) || checkPw_(pw);
    if (err) return { ok: false, message: err };
    const jeton = 't' + Utilities.getUuid();
    add_('Users', { id: id_('u'), pseudo, motdepasse: pw, jeton, cree: new Date() });
    return { ok: true, token: jeton, pseudo, created: true };
  } finally { lock.releaseLock(); }
}
/** Profil : infos du compte, cours suivis et résultats. */
function profil(tok) {
  const u = auth_(tok), res = rows_('Resultats').filter(r => S(r.uid) === S(u.id)), cs = rows_('Cours');
  const cours = rows_('Inscriptions').filter(i => S(i.uid) === S(u.id)).map(i => {
    const c = cs.filter(x => S(x.id) === S(i.cours))[0], r = ph => { const x = res.filter(y => S(y.cours) === S(i.cours) && y.phase === ph)[0]; return x ? x.score + '/' + x.total : null; };
    return { id: S(i.cours), titre: c ? c.titre : '(cours supprimé)', pre: r('Pré-test'), post: r('Post-test') };
  });
  return { ok: true, pseudo: S(u.pseudo), cree: u.cree, cours };
}
/** Modifier son pseudonyme et/ou son mot de passe (laisser vide pour ne pas changer). */
function updateProfile(tok, newPseudo, newPw) {
  const u = auth_(tok); newPseudo = S(newPseudo || '').trim();
  if (newPseudo && newPseudo !== S(u.pseudo)) {
    const err = checkPseudo_(newPseudo); if (err) return { ok: false, message: err };
    const ex = findUser_(newPseudo); if (ex && ex._r !== u._r) return { ok: false, message: 'Ce pseudonyme est déjà pris.' };
    set_('Users', u._r, 'pseudo', newPseudo);
  }
  if (newPw) { const err = checkPw_(newPw); if (err) return { ok: false, message: err }; set_('Users', u._r, 'motdepasse', S(newPw)); }
  return { ok: true, pseudo: newPseudo || S(u.pseudo) };
}

/* ---------- Cours et inscriptions ---------- */
function activeCourse_(cid) {
  const c = rows_('Cours').filter(x => S(x.id) === S(cid))[0];
  if (!c || c.actif === false || S(c.actif).toUpperCase() === 'FALSE') throw new Error('Cours introuvable.');
  return c;
}
function enrolled_(uid, cid) { return rows_('Inscriptions').some(i => S(i.uid) === S(uid) && S(i.cours) === S(cid)); }
function needEnrolled_(u, cid) { if (!enrolled_(u.id, cid)) throw new Error('Inscrivez-vous d’abord à ce cours.'); }

function me(tok) {
  const u = auth_(tok), mine = rows_('Inscriptions').filter(i => S(i.uid) === S(u.id)).map(i => S(i.cours));
  const cours = rows_('Cours').filter(c => c.actif !== false && S(c.actif).toUpperCase() !== 'FALSE')
    .map(c => ({ id: S(c.id), titre: c.titre, description: c.description, image: S(c.image), nouveau: !!c.cree && Date.now() - new Date(c.cree).getTime() < 7 * 864e5, inscrit: mine.indexOf(S(c.id)) >= 0 }));
  return { ok: true, pseudo: S(u.pseudo), cours };
}
function enroll(tok, cid) {
  const u = auth_(tok); activeCourse_(cid);
  if (!enrolled_(u.id, cid)) add_('Inscriptions', { uid: u.id, pseudo: u.pseudo, cours: S(cid), date: new Date() });
  return { ok: true };
}
function unenroll(tok, cid) {
  const u = auth_(tok);
  rows_('Inscriptions').filter(i => S(i.uid) === S(u.id) && S(i.cours) === S(cid)).reverse()
    .forEach(i => sh_('Inscriptions').deleteRow(i._r));
  return { ok: true };
}
/** Parcours d'un cours pour un utilisateur : pré-test → séquences dans l'ordre (documents obligatoires lus et marqués, puis quiz) → post-test.
 *  Documents sans séquence : les obligatoires forment une « introduction » avant la séquence 1 ; les facultatifs sont des ressources générales.
 *  Un cours sans pré-test n'impose pas ce verrou ; une séquence sans document ni quiz s'ouvre et se termine d'elle-même. */
const opt_ = d => d.facultatif === true || S(d.facultatif).toUpperCase() === 'TRUE';
function lectures_(uid, cid) {
  const m = {}; rows_('Lectures').filter(l => S(l.uid) === S(uid) && S(l.cours) === S(cid)).forEach(l => m[S(l.doc)] = l); return m;
}
function seqs_(cid) {
  return rows_('Sequences').filter(s => S(s.cours) === S(cid)).sort((a, b) => (Number(a.ordre) || 0) - (Number(b.ordre) || 0) || a._r - b._r);
}
/** Réglages d'un cours : pré-test facultatif ou non, note de réussite des quiz (%). */
function cfg_(cid) {
  const c = rows_('Cours').filter(x => S(x.id) === S(cid))[0] || {};
  let seuil = S(c.seuil).trim() === '' ? 70 : Number(c.seuil); if (!(seuil >= 0)) seuil = 70;
  return { pretestFac: S(c.pretest).trim().toLowerCase() === 'facultatif', seuil: Math.min(100, seuil), modules: S(c.mode).trim().toLowerCase() === 'modules' };
}
function plan_(uid, cid) {
  const cf = cfg_(cid);
  const preOk = cf.pretestFac || !questions_(cid, 'Pré-test').length || done_(uid, cid, 'Pré-test');
  const real = cf.modules ? seqs_(cid) : [], ids = {}; /* cours simple : les modules éventuels sont ignorés */ real.forEach(s => ids[S(s.id)] = 1);
  const docs = rows_('Documents').filter(d => S(d.cours) === S(cid)), lec = lectures_(uid, cid);
  const libres = docs.filter(d => !ids[S(d.sequence)]);
  const qs = rows_('Questions').filter(q => S(q.cours) === S(cid) && q.phase === 'Quiz');
  const res = rows_('Resultats').filter(r => S(r.uid) === S(uid) && S(r.cours) === S(cid) && r.phase === 'Quiz');
  const seqs = libres.some(d => !opt_(d)) ? [{ id: '', titre: 'Introduction', virtual: true }] : [];
  real.forEach(x => seqs.push({ id: S(x.id), titre: x.titre, intro: S(x.intro), image: S(x.image) }));
  let prev = true;
  const list = seqs.map(x => {
    const ds = x.virtual ? libres.filter(d => !opt_(d)) : docs.filter(d => S(d.sequence) === x.id);
    const req = ds.filter(d => !opt_(d)), lus = req.filter(d => lec[S(d.id)] && lec[S(d.id)].lu).length;
    const nq = x.virtual ? 0 : qs.filter(q => S(q.sequence) === x.id).length, att = res.filter(r => S(r.sequence) === x.id);
    // Quiz : plusieurs essais possibles ; le module est terminé quand le meilleur essai atteint la note de réussite
    const best = att.slice().sort((a, b) => b.score / b.total - a.score / a.total)[0] || null, last = att[att.length - 1] || null;
    const passed = !!best && (cf.seuil <= 0 || (best.total > 0 && 100 * best.score / best.total >= cf.seuil));
    const unlocked = preOk && prev, lu = lus === req.length, complete = unlocked && lu && (!nq || passed);
    prev = complete;
    return { id: x.id, titre: x.titre, image: x.image || '', intro: x.intro || '', virtual: !!x.virtual, docs: ds, req, lus, nq, att: att.length, best, last, passed, unlocked, lu, complete };
  });
  return { cf, preOk, list, lec, libres: libres.filter(opt_), postOk: preOk && prev };
}
function parseDetails_(r) { try { return r.details ? JSON.parse(r.details) : []; } catch (e) { return []; } }
function course(tok, cid) {
  const u = auth_(tok), c = activeCourse_(cid); needEnrolled_(u, cid);
  const qs = rows_('Questions').filter(q => S(q.cours) === S(cid));
  const res = {};
  rows_('Resultats').filter(r => S(r.uid) === S(u.id) && S(r.cours) === S(cid) && r.phase !== 'Quiz').forEach(r => res[r.phase] = { score: r.score, total: r.total, corrige: parseDetails_(r) });
  const p = plan_(u.id, cid);
  const dOut = d => { const l = p.lec[S(d.id)]; return { id: S(d.id), titre: d.titre, url: S(d.url).indexOf('html:') === 0 ? '' : d.url, type: S(d.type) || kind_(d.url), facultatif: opt_(d), ouvert: !!(l && l.ouvert), lu: !!(l && l.lu) }; };
  return { ok: true, id: S(c.id), titre: c.titre, description: c.description, image: S(c.image),
    // Contenu du cours envoyé seulement une fois le pré-test passé ; le contenu d'une séquence seulement quand elle est ouverte
    blocs: p.preOk ? rows_('Blocs').filter(b => S(b.cours) === S(cid)).map(b => ({ titre: b.titre, texte: b.texte })) : [],
    sequences: p.preOk ? p.list.map(s => ({ id: s.id, titre: s.titre, image: s.image, virtual: s.virtual, verrouille: !s.unlocked, complete: s.complete, lu: s.lu, lus: s.lus, total: s.req.length,
      documents: s.unlocked ? s.docs.map(dOut) : [],
      intro: s.unlocked ? s.intro : '',
      quiz: s.nq ? { n: s.nq, seuil: p.cf.seuil, essais: s.att, reussi: s.passed, meilleur: s.best ? { score: s.best.score, total: s.best.total } : null,
        dernier: s.last ? { score: s.last.score, total: s.last.total, corrige: parseDetails_(s.last) } : null } : null })) : [],
    ressources: p.preOk ? p.libres.map(dOut) : [],
    pretestFacultatif: p.cf.pretestFac, seuil: p.cf.seuil, mode: p.cf.modules ? 'modules' : 'simple',
    verrous: { docs: !p.preOk, post: !p.postOk, preManquant: !p.preOk, seqRestantes: p.list.filter(s => !s.complete).length, seqTotal: p.list.length },
    nbQuestions: { 'Pré-test': qs.filter(q => q.phase === 'Pré-test').length, 'Post-test': qs.filter(q => q.phase === 'Post-test').length },
    resultats: res };
}
/** Enregistre l'ouverture d'un document (condition pour pouvoir le marquer comme lu). */
function openDoc(tok, cid, did) { return lecture_(tok, cid, did, 'ouvert'); }
/** Marque un document comme lu en entier (seulement s'il a été ouvert). */
function markDoc(tok, cid, did) { return lecture_(tok, cid, did, 'lu'); }
function lecture_(tok, cid, did, etape) {
  const u = auth_(tok); activeCourse_(cid); needEnrolled_(u, cid);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const p = plan_(u.id, cid);
    if (!p.preOk) return { ok: false, message: 'Passez d’abord le pré-test.' };
    const doc = rows_('Documents').filter(d => S(d.id) === S(did) && S(d.cours) === S(cid))[0];
    if (!doc) return { ok: false, message: 'Document introuvable.' };
    if (opt_(doc)) return { ok: true, facultatif: true }; // ressource facultative : rien à enregistrer
    const sq = p.list.filter(s => s.req.some(d => S(d.id) === S(did)))[0];
    if (!sq || !sq.unlocked) return { ok: false, message: 'Terminez d’abord le module précédent.' };
    let l = lectures_(u.id, cid)[S(did)];
    if (!l) { add_('Lectures', { uid: u.id, cours: S(cid), doc: S(did), ouvert: etape === 'ouvert' ? new Date() : '', lu: '' }); l = lectures_(u.id, cid)[S(did)]; if (etape === 'ouvert') return { ok: true }; }
    if (etape === 'ouvert') { if (!l.ouvert) set_('Lectures', l._r, 'ouvert', new Date()); }
    else { if (!l.ouvert) return { ok: false, message: 'Ouvrez d’abord le document pour pouvoir le marquer comme lu.' }; if (!l.lu) set_('Lectures', l._r, 'lu', new Date()); }
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- Tests ---------- */
function shuffle_(a) { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; }
function questions_(cid, phase, sid) { return rows_('Questions').filter(q => S(q.cours) === S(cid) && q.phase === phase && (phase !== 'Quiz' || S(q.sequence) === S(sid))); }
function done_(uid, cid, phase, sid) { return rows_('Resultats').some(r => S(r.uid) === S(uid) && S(r.cours) === S(cid) && r.phase === phase && (phase !== 'Quiz' || S(r.sequence) === S(sid))); }

function typeOf_(q) { return q.type || (JSON.parse(q.bonnes).length > 1 ? 'multi' : 'single'); }
function dat_(q) { return q.donnees ? JSON.parse(q.donnees) : {}; }
function qOut_(q) {
  const t = typeOf_(q), d = dat_(q);
  const o = { id: S(q.id), type: t, q: q.question, ctx: d.ctx, ctxTable: d.ctxTable, chartData: d.chartData, rptDisease: d.rptDisease };
  if (t === 'dates') o.dates = d.dates.map(x => ({ k: x.k, label: x.label, o: x.o }));
  else if (t === 'match') { o.pre = d.pre; o.tubes = d.tubes; o.rows = d.rows.map(r => r.label); }
  else {
    const opts = JSON.parse(q.options).map((x, i) => ({ i, t: typeof x === 'string' ? x : x.t, chart: x.chart, rpt: x.rpt }));
    o.multi = t === 'multi'; o.o = d.fixe ? opts : shuffle_(opts);
  }
  return o;
}
function grade_(q, rep) {
  const t = typeOf_(q), d = dat_(q);
  if (t === 'dates') return !!rep && d.dates.every(x => S(rep[x.k]) === S(x.a));
  if (t === 'match') return Array.isArray(rep) && d.rows.every((r, i) => rep[i] && S(rep[i][0]) === S(r.p) && S(rep[i][1]) === S(r.t));
  const good = JSON.parse(q.bonnes).map(Number).sort(), got = (Array.isArray(rep) ? rep : []).map(Number).sort();
  return good.length === got.length && good.every((v, i) => v === got[i]);
}
function bonnes_(q) {
  const t = typeOf_(q), d = dat_(q);
  if (t === 'dates') return d.dates.map(x => x.label + ' : ' + x.o[x.a]);
  if (t === 'match') return d.rows.map(r => r.label + ' : ' + d.pre[r.p] + ' → ' + d.tubes[r.t].t);
  if (t === 'rpt') return [];
  return JSON.parse(q.bonnes).map(i => { const x = JSON.parse(q.options)[i]; return typeof x === 'string' ? x : x.t; });
}
/** Raison pour laquelle un test est fermé (chaîne vide = accessible). */
function testBlock_(uid, cid, phase, sid) {
  if (phase === 'Pré-test') return '';
  const p = plan_(uid, cid);
  if (!p.preOk) return 'Passez d’abord le pré-test.';
  if (phase === 'Post-test') return p.postOk ? '' : 'Terminez tous les modules du cours (documents lus et quiz réussi) avant le post-test.';
  const s = p.list.filter(x => x.id === S(sid) && !x.virtual)[0];
  if (!s) return 'Module introuvable.';
  if (!s.unlocked) return 'Terminez d’abord le module précédent.';
  if (s.passed) return 'Vous avez déjà réussi ce quiz.';
  if (!s.lu) return 'Lisez et marquez tous les documents de ce module (' + s.lus + '/' + s.req.length + ') avant le quiz.';
  return '';
}
function optTxt_(x, i) { return typeof x === 'string' ? x : (x && x.t) || ('Option ' + (i + 1)); }
/** Réponse de l'utilisateur sous forme de textes lisibles (pour la correction). */
function repTexte_(q, rep) {
  const t = typeOf_(q), d = dat_(q), vide = v => v === undefined || v === null || v === '';
  if (t === 'dates') return d.dates.map(x => x.label + ' : ' + (rep && !vide(rep[x.k]) ? x.o[Number(rep[x.k])] : '—'));
  if (t === 'match') return d.rows.map((r, i) => { const a = (Array.isArray(rep) && rep[i]) || []; return r.label + ' : ' + (vide(a[0]) ? '—' : d.pre[a[0]]) + ' → ' + (vide(a[1]) ? '—' : d.tubes[a[1]].t); });
  const o = JSON.parse(q.options || '[]');
  return (Array.isArray(rep) ? rep : []).map(Number).sort((x, y) => x - y).map(i => optTxt_(o[i], i));
}
function startTest(tok, cid, phase, sid) {
  const u = auth_(tok); activeCourse_(cid); needEnrolled_(u, cid);
  if (PHASES_ALL.indexOf(phase) < 0) return { ok: false, message: 'Phase inconnue.' };
  if (phase !== 'Quiz' && done_(u.id, cid, phase, sid)) return { ok: false, message: 'Vous avez déjà passé ce test.' };
  const e = testBlock_(u.id, cid, phase, sid); if (e) return { ok: false, message: e };
  const qs = questions_(cid, phase, sid);
  if (!qs.length) return { ok: false, message: 'Aucune question n’a encore été ajoutée.' };
  return { ok: true, pseudo: S(u.pseudo), limitSec: (phase === 'Quiz' ? 15 : 25) * 60, questions: shuffle_(qs.map(qOut_)) };
}
/** reponses : { idQuestion: réponse } — les questions non renseignées comptent comme fausses. */
function submitTest(tok, cid, phase, reponses, sid) {
  const u = auth_(tok); needEnrolled_(u, cid);
  if (PHASES_ALL.indexOf(phase) < 0) return { ok: false, message: 'Phase inconnue.' };
  if (phase !== 'Quiz' && done_(u.id, cid, phase, sid)) return { ok: false, message: 'Vous avez déjà passé ce test.' };
  const e = testBlock_(u.id, cid, phase, sid); if (e) return { ok: false, message: e };
  const qs = questions_(cid, phase, sid); let score = 0;
  const corr = qs.map(q => {
    const rep = (reponses || {})[S(q.id)], ok = grade_(q, rep); if (ok) score++;
    let r = []; try { r = repTexte_(q, rep); } catch (x) { }
    return { q: q.question, ok, rep: r, bonnes: bonnes_(q), explication: q.explication };
  });
  let det = JSON.stringify(corr);
  if (det.length > 45000) det = JSON.stringify(corr.map(c => ({ q: c.q, ok: c.ok, rep: c.rep, bonnes: c.bonnes }))); // limite d'une cellule Sheets
  add_('Resultats', { date: new Date(), uid: u.id, pseudo: u.pseudo, cours: S(cid), phase, score, total: qs.length, sequence: phase === 'Quiz' ? S(sid) : '', details: det });
  const seuil = cfg_(cid).seuil, reussi = phase === 'Quiz' ? (seuil <= 0 || (qs.length > 0 && 100 * score / qs.length >= seuil)) : null;
  return { ok: true, pseudo: S(u.pseudo), score, total: qs.length, corrige: corr, reussi, seuil };
}

/* ---------- Discussion ---------- */
function getMessages(tok, cid) {
  const u = auth_(tok); needEnrolled_(u, cid);
  const names = {}; rows_('Users').forEach(x => names[S(x.id)] = S(x.pseudo));
  const m = rows_('Messages').filter(x => S(x.cours) === S(cid)).slice(-100)
    .map(x => ({ id: S(x.id), date: x.date, pseudo: names[S(x.uid)] || S(x.pseudo), texte: x.texte, mine: S(x.uid) === S(u.id) }));
  return { ok: true, messages: m };
}
function postMessage(tok, cid, texte) {
  const u = auth_(tok); needEnrolled_(u, cid);
  texte = S(texte || '').trim().slice(0, 1000);
  if (!texte) return { ok: false, message: 'Message vide.' };
  add_('Messages', { id: id_('m'), date: new Date(), cours: S(cid), uid: u.id, pseudo: u.pseudo, texte });
  return { ok: true };
}
function ownMessage_(u, mid) {
  const m = rows_('Messages').filter(x => S(x.id) === S(mid) && S(x.uid) === S(u.id))[0];
  if (!m) throw new Error('Message introuvable.'); return m;
}
function editMessage(tok, mid, texte) { const m = ownMessage_(auth_(tok), mid); set_('Messages', m._r, 'texte', S(texte).trim().slice(0, 1000)); return { ok: true }; }
function deleteMessage(tok, mid) { const m = ownMessage_(auth_(tok), mid); sh_('Messages').deleteRow(m._r); return { ok: true }; }

/* ---------- Administration ---------- */
function admin_(pw) { if (ADMIN_PASSWORD === 'CHANGEZ-MOI') throw new Error('Définissez ADMIN_PASSWORD dans Code.gs puis redéployez.'); if (S(pw) !== ADMIN_PASSWORD) throw new Error('Mot de passe administrateur incorrect.'); }
function adminLogin(pw) { admin_(pw); return { ok: true }; }
function adminOverview(pw) {
  admin_(pw);
  const ins = rows_('Inscriptions'), res = rows_('Resultats'), qs = rows_('Questions'), docs = rows_('Documents');
  const moy = (cid, ph) => { const r = res.filter(x => S(x.cours) === cid && x.phase === ph && x.total > 0); return r.length ? Math.round(r.reduce((s, x) => s + 100 * x.score / x.total, 0) / r.length) : null; };
  return { ok: true, cours: rows_('Cours').map(c => { const id = S(c.id); return {
    id, titre: c.titre, actif: !(c.actif === false || S(c.actif).toUpperCase() === 'FALSE'),
    inscrits: ins.filter(i => S(i.cours) === id).length,
    documents: docs.filter(d => S(d.cours) === id).length,
    pre: qs.filter(q => S(q.cours) === id && q.phase === 'Pré-test').length,
    post: qs.filter(q => S(q.cours) === id && q.phase === 'Post-test').length,
    moyPre: moy(id, 'Pré-test'), moyPost: moy(id, 'Post-test') }; }) };
}
function adminSaveCourse(pw, c) {
  admin_(pw);
  const titre = S(c && c.titre).trim().slice(0, 120); if (!titre) return { ok: false, message: 'Titre requis.' };
  const desc = S(c.description || '').slice(0, 500);
  const ex = c.id && rows_('Cours').filter(x => S(x.id) === S(c.id))[0];
  const mode = c.mode === 'modules' ? 'modules' : 'simple', pre = c.pretest === 'facultatif' ? 'facultatif' : 'obligatoire', sc = S(c.seuil === undefined ? '' : c.seuil).trim(), seuil = sc === '' ? '' : Math.max(0, Math.min(100, Number(sc) || 0));
  if (ex) { set_('Cours', ex._r, 'titre', titre); set_('Cours', ex._r, 'description', desc); set_('Cours', ex._r, 'actif', c.actif !== false);
    if (c.mode !== undefined) set_('Cours', ex._r, 'mode', mode); if (c.pretest !== undefined) set_('Cours', ex._r, 'pretest', pre); if (c.seuil !== undefined) set_('Cours', ex._r, 'seuil', seuil); return { ok: true, id: S(c.id) }; }
  const id = id_('c'); add_('Cours', { id, titre, description: desc, actif: true, pretest: pre, seuil, mode });
  if (mode === 'modules') add_('Sequences', { id: id_('s'), cours: id, titre: 'Introduction et généralités', ordre: 1 }); // premier module : introduction (quiz facultatif)
  return { ok: true, id };
}
/** doc : { titre, url } ou { titre, fichier: { nom, mime, b64 } } (déposé dans Drive, lisible par lien). */
/** Identifiant de séquence valide du cours, sinon chaîne vide (hors séquence). */
function seqOk_(cid, sid) { return seqs_(cid).some(x => S(x.id) === S(sid)) ? S(sid) : ''; }
function adminAddDoc(pw, cid, doc) {
  admin_(pw); activeCourse_(cid);
  let url = S(doc.url || '').trim();
  if (doc.fichier) {
    const f = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(doc.fichier.b64), doc.fichier.mime, doc.fichier.nom));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); url = f.getUrl();
  }
  if (!url) return { ok: false, message: 'Indiquez un lien ou un fichier.' };
  const type = kind_(url, doc.fichier && doc.fichier.nom, doc.fichier && doc.fichier.mime);
  add_('Documents', { type, id: id_('d'), cours: S(cid), titre: S(doc.titre || (doc.fichier && doc.fichier.nom) || 'Document').slice(0, 120), url, facultatif: doc.facultatif === true, sequence: seqOk_(cid, doc.sequence) });
  return { ok: true };
}
function adminAddQuestion(pw, cid, q) {
  admin_(pw); activeCourse_(cid);
  const opts = (q.options || []).map(x => S(x).trim()).filter(Boolean);
  const bonnes = (q.bonnes || []).map(Number).filter(i => i >= 0 && i < opts.length);
  if (PHASES_ALL.indexOf(q.phase) < 0 || !S(q.question).trim() || opts.length < 2 || !bonnes.length) return { ok: false, message: 'Question, au moins 2 choix et une bonne réponse requis.' };
  const sq = q.phase === 'Quiz' ? seqOk_(cid, q.sequence) : ''; if (q.phase === 'Quiz' && !sq) return { ok: false, message: 'Choisissez le module du quiz.' };
  add_('Questions', { id: id_('q'), cours: S(cid), phase: q.phase, type: bonnes.length > 1 ? 'multi' : 'single', question: S(q.question).trim(), options: JSON.stringify(opts), bonnes: JSON.stringify(bonnes), donnees: '{}', explication: S(q.explication || ''), sequence: sq });
  return { ok: true };
}
/** Page d'administration d'un cours : tout son contenu et les résultats (pseudonyme = code de travail). */
function adminCourseDetail(pw, cid) {
  admin_(pw);
  const c = rows_('Cours').filter(x => S(x.id) === S(cid))[0]; if (!c) return { ok: false, message: 'Cours introuvable.' };
  const mine = r => S(r.cours) === S(cid);
  return { ok: true,
    cours: { id: S(c.id), titre: c.titre, description: c.description, image: S(c.image), actif: !(c.actif === false || S(c.actif).toUpperCase() === 'FALSE'), pretest: cfg_(cid).pretestFac ? 'facultatif' : 'obligatoire', seuil: cfg_(cid).seuil, mode: cfg_(cid).modules ? 'modules' : 'simple' },
    sequences: seqs_(cid).map(x => ({ id: S(x.id), titre: x.titre, intro: S(x.intro), image: S(x.image) })),
    documents: rows_('Documents').filter(mine).map(d => ({ id: S(d.id), titre: d.titre, url: d.url, type: S(d.type) || kind_(d.url), facultatif: opt_(d), sequence: seqOk_(cid, d.sequence) })),
    blocs: rows_('Blocs').filter(mine).map(b => ({ id: S(b.id), titre: b.titre, texte: b.texte })),
    questions: rows_('Questions').filter(mine).map(q => ({ id: S(q.id), phase: q.phase, sequence: S(q.sequence), type: typeOf_(q), question: q.question, explication: q.explication })),
    resultats: rows_('Resultats').filter(mine).map(r => ({ pseudo: r.pseudo, phase: r.phase, sequence: S(r.sequence), score: r.score, total: r.total, date: r.date, corrige: parseDetails_(r) })) };
}
function adminDelete(pw, table, id) {
  admin_(pw);
  if (['Documents', 'Questions', 'Blocs'].indexOf(table) < 0) return { ok: false, message: 'Table non autorisée.' };
  const r = rows_(table).filter(x => S(x.id) === S(id))[0];
  if (r && table === 'Documents' && S(r.url).indexOf('html:') === 0) { try { DriveApp.getFileById(S(r.url).slice(5)).setTrashed(true); } catch (e) { } }
  if (r) sh_(table).deleteRow(r._r);
  return { ok: true };
}
/** Élément de texte affiché sur la page du cours (consigne, résumé, objectifs…). */
function adminAddBloc(pw, cid, titre, texte) {
  admin_(pw); activeCourse_(cid);
  if (!S(texte).trim()) return { ok: false, message: 'Texte requis.' };
  add_('Blocs', { id: id_('b'), cours: S(cid), titre: S(titre || '').slice(0, 120), texte: S(texte).slice(0, 5000) });
  return { ok: true };
}
/** Import du cours MPE : nécessite le fichier Seed.gs dans le projet (sans lui, le reste de la plateforme fonctionne quand même). */
function adminImporterMPE_(pw) {
  admin_(pw);
  if (typeof importerCoursMPE !== 'function') return { ok: false, message: 'Ajoutez le fichier Seed.gs au projet Apps Script, puis redéployez.' };
  return importerCoursMPE();
}
/** Supprime un cours et tout ce qui s'y rattache (documents, textes, questions, inscriptions, résultats, messages). */
function adminDeleteCourse(pw, cid) {
  admin_(pw);
  const n = {};
  ['Documents', 'Blocs', 'Questions', 'Inscriptions', 'Resultats', 'Messages', 'Lectures', 'Sequences'].forEach(t => {
    const r = rows_(t).filter(x => S(x.cours) === S(cid)).reverse(); r.forEach(x => sh_(t).deleteRow(x._r)); n[t] = r.length;
  });
  const c = rows_('Cours').filter(x => S(x.id) === S(cid))[0];
  if (!c) return { ok: false, message: 'Cours introuvable.' };
  sh_('Cours').deleteRow(c._r);
  return { ok: true, supprime: n };
}
/** Modifie un élément d'un cours : document (titre, url), texte (titre, texte), question (question, explication, phase). */
function adminUpdate(pw, table, id, champs) {
  admin_(pw);
  const ok = { Documents: ['titre', 'url', 'facultatif', 'sequence'], Blocs: ['titre', 'texte'], Questions: ['question', 'explication', 'phase'] }[table];
  if (!ok) return { ok: false, message: 'Table non autorisée.' };
  const r = rows_(table).filter(x => S(x.id) === S(id))[0]; if (!r) return { ok: false, message: 'Élément introuvable.' };
  for (const k of Object.keys(champs || {})) {
    if (ok.indexOf(k) < 0) continue;
    if (k === 'facultatif') { set_(table, r._r, k, champs[k] === true || S(champs[k]).toLowerCase() === 'true'); continue; }
    if (k === 'sequence') { set_(table, r._r, k, seqOk_(r.cours, champs[k])); continue; }
    const v = S(champs[k]).trim();
    if (k === 'phase' && PHASES.indexOf(v) < 0) return { ok: false, message: 'Phase inconnue.' };
    if ((k === 'titre' && table !== 'Blocs') || k === 'question' || k === 'url' || k === 'texte') { if (!v) return { ok: false, message: 'Champ « ' + k + ' » vide.' }; }
    set_(table, r._r, k, v);
  }
  return { ok: true };
}
/** Import de questions depuis un fichier CSV lu par la page. lignes : [{phase,type,question,options,bonnes,explication,donnees}] ; bonnes = numéros à partir de 1. */
function adminImportQuestions(pw, cid, lignes, remplacer) {
  admin_(pw); activeCourse_(cid);
  const TYPES = ['single', 'multi', 'chart', 'rpt', 'dates', 'match'], bons = [], err = [];
  (lignes || []).forEach((l, i) => {
    try {
      const phase = S(l.phase).trim(), type = S(l.type || '').trim() || 'single', question = S(l.question).trim();
      if (PHASES_ALL.indexOf(phase) < 0) throw new Error('phase inconnue (« Pré-test », « Post-test » ou « Quiz »)');
      let sq = ''; if (phase === 'Quiz') { sq = seqRef_(cid, l.sequence); if (!sq) throw new Error('quiz : colonne « sequence » (numéro ou titre de la séquence) manquante ou inconnue'); }
      if (TYPES.indexOf(type) < 0) throw new Error('type inconnu');
      if (!question) throw new Error('question vide');
      const o = S(l.options || '').trim();
      const options = !o ? [] : (o[0] === '[' ? JSON.parse(o) : o.split('|').map(x => x.trim()).filter(Boolean));
      const bonnes = S(l.bonnes || '').split(/[,;\s]+/).filter(Boolean).map(n => Number(n) - 1);
      const d = S(l.donnees || '').trim() ? JSON.parse(l.donnees) : {};
      if (['dates'].indexOf(type) >= 0 && !(d.dates && d.dates.length)) throw new Error('« donnees » doit contenir les dates');
      if (type === 'match' && !(d.pre && d.tubes && d.rows)) throw new Error('« donnees » doit contenir pre, tubes et rows');
      if (type !== 'dates' && type !== 'match') {
        if (options.length < 2) throw new Error('au moins 2 choix');
        if (!bonnes.length || bonnes.some(b => !(b >= 0 && b < options.length))) throw new Error('numéros de bonnes réponses invalides');
        if (type !== 'multi' && bonnes.length !== 1) throw new Error('une seule bonne réponse attendue');
      }
      bons.push({ phase, sequence: sq, type, question, options, bonnes: (type === 'dates' || type === 'match') ? [] : bonnes, donnees: d, explication: S(l.explication || '') });
    } catch (e) { err.push('Ligne ' + (i + 2) + ' : ' + e.message); }
  });
  if (remplacer && bons.length) {
    const ph = {}; bons.forEach(b => ph[b.phase === 'Quiz' ? 'Quiz:' + b.sequence : b.phase] = 1);
    rows_('Questions').filter(q => S(q.cours) === S(cid) && ph[q.phase === 'Quiz' ? 'Quiz:' + S(q.sequence) : q.phase]).reverse().forEach(q => sh_('Questions').deleteRow(q._r));
  }
  bons.forEach(b => add_('Questions', { id: id_('q'), cours: S(cid), phase: b.phase, type: b.type, question: b.question,
    options: JSON.stringify(b.options), bonnes: JSON.stringify(b.bonnes), donnees: JSON.stringify(b.donnees), explication: b.explication, sequence: b.sequence }));
  return { ok: true, ajoutees: bons.length, erreurs: err };
}

/** Séquence désignée par un numéro (1 = première), un titre ou un identifiant ; chaîne vide si inconnue. */
function seqRef_(cid, ref) {
  ref = S(ref || '').trim(); if (!ref) return '';
  const l = seqs_(cid), n = Number(ref), x = (n >= 1 && l[n - 1]) || l.filter(y => S(y.id) === ref || S(y.titre).trim().toLowerCase() === ref.toLowerCase())[0];
  return x ? S(x.id) : '';
}
/* ---------- Séquences (administration) ---------- */
/** Crée une séquence (c = { titre }) ou la renomme (c = { id, titre }). Ajoutée à la fin du cours. */
function adminSaveSequence(pw, cid, c) {
  admin_(pw); activeCourse_(cid);
  const titre = S(c && c.titre).trim().slice(0, 120); if (!titre) return { ok: false, message: 'Titre requis.' };
  const ex = c.id && seqs_(cid).filter(x => S(x.id) === S(c.id))[0];
  const intro = S(c.intro || '').slice(0, 2000);
  if (ex) { set_('Sequences', ex._r, 'titre', titre); if (c.intro !== undefined) set_('Sequences', ex._r, 'intro', intro); return { ok: true, id: S(ex.id) }; }
  const id = id_('s'); add_('Sequences', { id, cours: S(cid), titre, ordre: seqs_(cid).length + 1, intro }); return { ok: true, id };
}
/** Déplace une séquence vers le haut (delta = -1) ou vers le bas (+1). */
function adminMoveSequence(pw, cid, sid, delta) {
  admin_(pw);
  const l = seqs_(cid), i = l.findIndex(x => S(x.id) === S(sid)), j = i + Number(delta);
  if (i < 0 || j < 0 || j >= l.length) return { ok: true };
  const t = l[i]; l[i] = l[j]; l[j] = t;
  l.forEach((x, k) => set_('Sequences', x._r, 'ordre', k + 1));
  return { ok: true };
}
/** Supprime une séquence avec ses documents, ses questions de quiz et les résultats de ce quiz. */
function adminDeleteSequence(pw, cid, sid) {
  admin_(pw);
  const x = seqs_(cid).filter(y => S(y.id) === S(sid))[0]; if (!x) return { ok: false, message: 'Module introuvable.' };
  [['Documents', r => true], ['Questions', r => r.phase === 'Quiz'], ['Resultats', r => r.phase === 'Quiz']].forEach(t =>
    rows_(t[0]).filter(r => S(r.cours) === S(cid) && S(r.sequence) === S(sid) && t[1](r)).reverse().forEach(r => sh_(t[0]).deleteRow(r._r)));
  sh_('Sequences').deleteRow(x._r);
  return { ok: true };
}

/* ---------- Image de couverture d'un cours ---------- */
/** img : { fichier: { nom, mime, b64 } } (déposé dans Drive, lisible par lien) ou { url } (adresse https) ; { url: '' } retire l'image. */
function adminSetCourseImage(pw, cid, img) {
  admin_(pw);
  const c = rows_('Cours').filter(x => S(x.id) === S(cid))[0]; if (!c) return { ok: false, message: 'Cours introuvable.' };
  let url = '';
  if (img && img.fichier) {
    if (!/^image\//.test(S(img.fichier.mime))) return { ok: false, message: 'Le fichier doit être une image.' };
    const f = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(img.fichier.b64), img.fichier.mime, img.fichier.nom));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    url = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1200';
  } else if (img && S(img.url).trim()) {
    url = S(img.url).trim(); if (!/^https:\/\//i.test(url)) return { ok: false, message: 'Le lien de l’image doit commencer par https://.' };
  }
  set_('Cours', c._r, 'image', url);
  return { ok: true, image: url };
}

/* ---------- Notifications ---------- */
/** Interrogé régulièrement par la page : nouveaux messages des autres dans mes cours, et nouveaux cours, depuis `since` (ms). */
function notifs(tok, since) {
  const u = auth_(tok), now = Date.now(); since = Number(since) || now;
  const mine = {}; rows_('Inscriptions').filter(i => S(i.uid) === S(u.id)).forEach(i => mine[S(i.cours)] = 1);
  const cs = rows_('Cours').filter(c => c.actif !== false && S(c.actif).toUpperCase() !== 'FALSE'), titre = {}; cs.forEach(c => titre[S(c.id)] = c.titre);
  const names = {}; rows_('Users').forEach(x => names[S(x.id)] = S(x.pseudo));
  const messages = rows_('Messages').filter(m => mine[S(m.cours)] && titre[S(m.cours)] !== undefined && S(m.uid) !== S(u.id) && new Date(m.date).getTime() > since).slice(-20)
    .map(m => ({ id: S(m.id), cours: S(m.cours), titreCours: titre[S(m.cours)], pseudo: names[S(m.uid)] || S(m.pseudo), texte: S(m.texte).slice(0, 140) }));
  const cours = cs.filter(c => c.cree && new Date(c.cree).getTime() > since && !mine[S(c.id)]).map(c => ({ id: S(c.id), titre: c.titre, description: S(c.description).slice(0, 140), image: S(c.image) }));
  return { ok: true, now, messages, cours };
}

/* ---------- Types de documents, leçons HTML, images de modules ---------- */
/** Type d'un document d'après son adresse, son nom ou son type MIME. */
function kind_(url, nom, mime) {
  const t = (S(url) + ' ' + S(nom) + ' ' + S(mime)).toLowerCase();
  if (/^html:/.test(S(url))) return 'page';
  if (/youtube\.com|youtu\.be|vimeo\.com|video\/|\.(mp4|webm|mov|m4v|ogg)(\?|\s|$)/.test(t)) return 'video';
  if (/\.pdf(\?|\s|$)|pdf/.test(t)) return 'pdf';
  if (/\.docx?(\?|\s|$)|msword|wordprocessingml|docs\.google\.com\/document/.test(t)) return 'word';
  if (/\.pptx?(\?|\s|$)|presentation|docs\.google\.com\/presentation/.test(t)) return 'ppt';
  if (/\.xlsx?(\?|\s|$)|spreadsheet|docs\.google\.com\/spreadsheets/.test(t)) return 'excel';
  if (/image\/|\.(png|jpe?g|gif|webp)(\?|\s|$)/.test(t)) return 'image';
  return 'lien';
}
/** Dépose une image dans Drive (lisible par lien) et renvoie son adresse d'affichage. */
function driveImage_(f) {
  if (!f || !/^image\//.test(S(f.mime))) return { ok: false, message: 'Le fichier doit être une image.' };
  const file = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(f.b64), f.mime, f.nom));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200' };
}
/** Image de présentation d'un module (affichée avec son titre ; un clic dessus ouvre le module). img = { fichier } ou { url } ; { url: '' } la retire. */
function adminSetModuleImage(pw, cid, sid, img) {
  admin_(pw);
  const x = seqs_(cid).filter(y => S(y.id) === S(sid))[0]; if (!x) return { ok: false, message: 'Module introuvable.' };
  let url = '';
  if (img && img.fichier) { const r = driveImage_(img.fichier); if (!r.ok) return r; url = r.url; }
  else if (img && S(img.url).trim()) { url = S(img.url).trim(); if (!/^https:\/\//i.test(url)) return { ok: false, message: 'Le lien de l’image doit commencer par https://.' }; }
  set_('Sequences', x._r, 'image', url); return { ok: true, image: url };
}
/** Dépose une image pour l'insérer dans une leçon HTML : renvoie son adresse. */
function adminUploadImage(pw, cid, img) { admin_(pw); activeCourse_(cid); return driveImage_(img && img.fichier); }
/** Leçon HTML : page complète (HTML/CSS/JS, personnages, photos…) enregistrée dans Drive et affichée dans un cadre isolé. */
function adminAddPage(pw, cid, p) {
  admin_(pw); activeCourse_(cid);
  const html = S(p && p.html); if (!html.trim()) return { ok: false, message: 'Le code HTML est vide.' };
  if (html.length > 3000000) return { ok: false, message: 'Page trop volumineuse (3 Mo maximum) : utilisez des images hébergées.' };
  const titre = S(p.titre || 'Leçon').trim().slice(0, 120) || 'Leçon';
  const f = DriveApp.createFile(titre + '.html', html, MimeType.HTML);
  add_('Documents', { id: id_('d'), cours: S(cid), titre, url: 'html:' + f.getId(), facultatif: p.facultatif === true, sequence: seqOk_(cid, p.sequence), type: 'page' });
  return { ok: true };
}
function pageDoc_(did) { const d = rows_('Documents').filter(x => S(x.id) === S(did))[0]; return d && S(d.url).indexOf('html:') === 0 ? d : null; }
function adminGetPage(pw, did) {
  admin_(pw); const d = pageDoc_(did); if (!d) return { ok: false, message: 'Leçon introuvable.' };
  return { ok: true, titre: d.titre, html: DriveApp.getFileById(S(d.url).slice(5)).getBlob().getDataAsString('UTF-8') };
}
function adminSavePage(pw, did, p) {
  admin_(pw); const d = pageDoc_(did); if (!d) return { ok: false, message: 'Leçon introuvable.' };
  const html = S(p && p.html); if (!html.trim()) return { ok: false, message: 'Le code HTML est vide.' };
  DriveApp.getFileById(S(d.url).slice(5)).setContent(html);
  if (S(p.titre).trim()) set_('Documents', d._r, 'titre', S(p.titre).trim().slice(0, 120));
  return { ok: true };
}
/** Renvoie le code d'une leçon HTML à un apprenant inscrit, seulement si le module est ouvert pour lui. */
function getPage(tok, cid, did) {
  const u = auth_(tok); activeCourse_(cid); needEnrolled_(u, cid);
  const d = pageDoc_(did); if (!d || S(d.cours) !== S(cid)) return { ok: false, message: 'Leçon introuvable.' };
  const p = plan_(u.id, cid); if (!p.preOk) return { ok: false, message: 'Passez d’abord le pré-test.' };
  const sq = p.list.filter(x => x.docs.some(y => S(y.id) === S(did)))[0]; if (sq && !sq.unlocked) return { ok: false, message: 'Terminez d’abord le module précédent.' };
  return { ok: true, html: DriveApp.getFileById(S(d.url).slice(5)).getBlob().getDataAsString('UTF-8') };
}
