/**
 * PRÉ-TEST / POST-TEST — Surveillance épidémiologique, ZS Djidja-Abomey-Agbangnizoun
 * Contenu : définitions de cas (pré-test du Jour 1), fiche IDS et rapport hebdomadaire, tracé de la courbe.
 *
 * Fonctionnement :
 *  - L'ADMINISTRATEUR (mot de passe) choisit la phase active (Pré-test ou Post-test), un titre
 *    d'activité visible par tous et une session/classe (sert à comparer pré et post), puis ouvre le formulaire.
 *  - Les participants saisissent seulement leur identifiant unique et passent la phase activée.
 *  - Les 20 questions (et leurs choix) sont mélangées pour chaque personne et chaque activité.
 *  - Types de questions : choix unique, choix multiple (tout ou rien), trois dates à choisir, figure ou dessin à choisir, maladies à relier aux prélèvements et aux tubes.
 *  - Les bonnes réponses restent sur le serveur ; le corrigé n'est envoyé qu'après la soumission.
 *  - Un identifiant ne peut soumettre qu'une fois par session et par phase.
 *  - Durée limitée à 15 minutes : à l'échéance les réponses partent automatiquement ; toute question non renseignée compte comme fausse.
 *  - Le post-test n'est accessible qu'avec un identifiant déjà enregistré au pré-test de la MÊME session.
 *  - Résultats écrits dans le Google Sheet (onglets Resultats et Questions) ; statistiques dans l'espace admin.
 *  - Ce script est seulement l'API (pas de page HTML) : la page index.html est hébergée sur GitHub Pages ou un autre serveur.
 */

// >>> À MODIFIER avant le déploiement : mot de passe de l'espace administrateur <<<
const ADMIN_PASSWORD = 'Silo@2000';

const SHEET_ID = '1hZ7v3_5K2qhD5N4YRmt3C96OsxOIk0iuQdQUwYSrwQc';
const PHASES = ['Pré-test', 'Post-test'];
const TIME_LIMIT_SEC = 15 * 60; // durée du test : à l'échéance, la page envoie les réponses ; ce qui n'est pas renseigné compte comme faux
const COL = { TS: 0, ID: 1, SESSION: 2, PHASE: 3, TITLE: 4, SCORE: 5, PCT: 6, DUR: 7, ORDER: 8 };

/* ------------------------------------------------------------------ */
/* BANQUE DE QUESTIONS                                                 */
/*  type : 'single' (défaut) | 'multi' (tout ou rien) | 'dates'        */
/*  a    : index de la bonne réponse (single) ou liste d'index (multi) */
/*  Un choix peut être un texte ou { rpt: [[suspects],[décès],[confirmés]] } */
/* ------------------------------------------------------------------ */
const QUESTIONS = [
  // ---------- Définitions de cas (pré-test du Jour 1) ----------
  {
    q: 'MPE signifie :',
    o: ['Maladies à Potentiel Épidémique', 'Maladies à Prise en charge Éducative', 'Maladies Parasitaires Endémiques'],
    a: 0,
    e: 'MPE signifie Maladies à Potentiel Épidémique.'
  },
  {
    q: 'Un cas suspect de rougeole présente :',
    o: [
      'Une toux de plus de 3 semaines',
      'Une fièvre avec éruption cutanée maculopapulaire',
      'Une diarrhée sanglante isolée'
    ],
    a: 1,
    e: 'Cas suspect de rougeole : fièvre (> 38 °C ou corps chaud) avec éruption cutanée maculopapulaire, associée ou non à une toux, un coryza ou une conjonctivite.'
  },
  {
    q: 'La rougeole est confirmée au laboratoire par la recherche des :',
    o: ['Parasites dans le sang', 'Globules blancs', 'Anticorps IgM'],
    a: 2,
    e: 'La rougeole est confirmée au laboratoire par la présence d’anticorps IgM (IgM positives).'
  },
  {
    q: 'Un cas suspect de shigellose est une diarrhée avec :',
    o: [
      'Du sang rouge visible dans les selles',
      'Des glaires blanches seulement',
      'Des vomissements seulement'
    ],
    a: 0,
    e: 'Shigellose : diarrhée accompagnée d’une présence physique de sang rouge dans les selles.'
  },
  {
    q: 'Quelle est la définition d’un cas de PFA (cas suspect de poliomyélite) ?',
    o: [
      'Tout enfant de moins de 15 ans avec une paralysie flasque de début brutal, sans notion de traumatisme',
      'Tout enfant de moins de 5 ans avec une fièvre et une éruption cutanée',
      'Tout adulte avec une paralysie spastique installée depuis plus d’un mois',
      'Tout enfant de moins de 15 ans avec une paralysie survenue après une chute ou une fracture'
    ],
    a: 0,
    e: 'Cas de PFA (cas suspect de poliomyélite) : tout enfant de moins de 15 ans avec un trouble de la marche, une boiterie ou une paralysie flasque (même transitoire), de début brutal et sans notion de traumatisme.'
  },
  {
    q: 'Un cas suspect de fièvre jaune associe une fièvre brutale et :',
    o: ['Une éruption', 'Une jaunisse', 'Une toux'],
    a: 1,
    e: 'Fièvre jaune : brusque accès de fièvre (> 38 °C) accompagné d’une jaunisse dans les 2 semaines qui suivent l’apparition des premiers symptômes.'
  },
  {
    q: 'Le choléra se manifeste par une diarrhée :',
    o: ['Chronique depuis des mois', 'Glaireuse avec ténesme', 'Aqueuse aiguë'],
    a: 2,
    e: 'Choléra : diarrhée aqueuse aiguë, avec ou sans déshydratation.'
  },
  {
    type: 'match',
    q: 'Pour chaque maladie, indiquez le prélèvement à faire, puis le contenant (tube ou pot) à utiliser.',
    pre: ['Sang', 'Selles', 'Urines', 'Prélèvement de gorge'],
    tubes: [
      { id: 'dry', t: 'Tube sec (bouchon rouge)' },
      { id: 'edta', t: 'Tube à bouchon violet' },
      { id: 'green', t: 'Tube à bouchon vert' },
      { id: 'pot', t: 'Pot à selles' }
    ],
    rows: [
      { label: 'Rougeole', p: 0, t: 0 },
      { label: 'Fièvre jaune', p: 0, t: 0 },
      { label: 'PFA', p: 1, t: 3 },
      { label: 'Shigellose', p: 1, t: 3 }
    ],
    e: 'Rougeole et fièvre jaune : prélèvement de sang, dans un tube sec. PFA : deux prélèvements de selles, à 24 heures d’intervalle, dans un pot à selles. Shigellose : prélèvement de selles, dans un pot à selles. Tous les prélèvements sont acheminés au laboratoire entre +2 et +8 °C, avec la fiche IDS remplie.'
  },
  {
    q: 'Devant un cas suspect de MPE, il faut :',
    o: ['Attendre la confirmation du laboratoire', 'Notifier immédiatement', 'Attendre d’avoir 3 cas'],
    a: 1,
    e: 'Cas suspect = on notifie tout de suite. On n’attend jamais la confirmation du laboratoire ni un nombre minimum de cas.'
  },

  // ---------- Fiche IDS, circuit et rapport hebdomadaire ----------
  {
    q: 'Quel est le circuit de la fiche IDS et de l’information ?',
    o: [
      'Formation sanitaire, puis zone sanitaire, puis laboratoire national ; les résultats reviennent à la zone sanitaire puis au clinicien',
      'Laboratoire national, puis formation sanitaire, puis zone sanitaire',
      'Formation sanitaire, puis laboratoire national directement ; la zone sanitaire n’est informée qu’à la fin',
      'Zone sanitaire, puis formation sanitaire, puis laboratoire national'
    ],
    a: 0,
    e: 'La formation sanitaire détecte le cas, remplit la fiche et prélève ; la zone sanitaire reçoit la notification, investigue et envoie le prélèvement ; le laboratoire national analyse et inscrit les résultats ; ils sont renvoyés à la zone sanitaire puis au clinicien. En cas grave ou épidémique, on appelle la zone sanitaire tout de suite : la fiche suit mais ne remplace pas l’alerte.'
  },
  {
    q: 'À quoi sert le rapport hebdomadaire ?',
    o: [
      'À remplacer la fiche IDS de chaque cas',
      'À informer chaque semaine la zone sanitaire de la situation des maladies et événements sous surveillance dans la formation sanitaire',
      'À confirmer les cas au laboratoire',
      'À enregistrer les consultations de routine'
    ],
    a: 1,
    e: 'Le rapport hebdomadaire est envoyé chaque lundi avant 10 h : il donne à la zone sanitaire la situation de la semaine (cas suspects, décès, confirmés, flambées). Il ne remplace pas la fiche IDS, qui reste individuelle.'
  },
  {
    type: 'multi',
    q: 'Quels éléments figurent dans le rapport hebdomadaire ?',
    o: [
      'L’en-tête : département, zone sanitaire, commune, formation sanitaire, semaine',
      'La situation de la SIMR : le tableau des maladies et événements sous surveillance',
      'Les résultats du laboratoire pour chaque cas',
      'La situation des décès',
      'Le statut vaccinal de chaque patient',
      'La situation des flambées',
      'La situation du paludisme',
      'Le numéro EPID de chaque cas',
      'Le rapporteur : nom et signature'
    ],
    a: [0, 1, 3, 5, 6, 8],
    e: 'Le rapport hebdomadaire comporte : l’en-tête, la situation de la SIMR (tableau des maladies), la situation des décès, la situation des flambées, la situation du paludisme et le rapporteur (nom et signature). Les résultats du laboratoire, le statut vaccinal et le numéro EPID relèvent de la fiche IDS.'
  },
  {
    q: 'Combien de cas suspects de rougeole de moins de 5 ans compte-t-on dans ce registre ?',
    ctx: 'Registre des cas suspects vus dans la formation sanitaire cette semaine :',
    ctxTable: {
      head: ['N°', 'Maladie', 'Âge', 'Sexe', 'Issue'],
      rows: [
        [1, 'Rougeole', '3 ans', 'F', 'Vivant'],
        [2, 'Choléra', '30 ans', 'M', 'Vivant'],
        [3, 'Rougeole', '8 mois', 'M', 'Vivant'],
        [4, 'Rougeole', '6 ans', 'M', 'Vivant'],
        [5, 'PFA', '5 ans', 'M', 'Vivant'],
        [6, 'Rougeole', '2 ans', 'M', 'Décédé'],
        [7, 'Shigellose', '1 an', 'F', 'Vivant'],
        [8, 'Rougeole', '12 ans', 'F', 'Vivant']
      ]
    },
    o: ['2', '3', '4', '5'],
    a: 1,
    e: 'Rougeole de moins de 5 ans : 3 ans (F), 8 mois (M) et 2 ans (M), soit 3 cas. Les rougeoles de 6 ans et de 12 ans sont dans la colonne 5 ans et plus ; le choléra, la PFA et la shigellose se comptent sur leur propre ligne. Le total de 5 est celui de toute la rougeole, tous âges confondus.'
  },
  {
    q: 'Quelle figure du rapport hebdomadaire est juste pour la ligne Rougeole de cette semaine ?',
    ctx: 'Voici les cas suspects de rougeole vus cette semaine. Aucun cas n’est confirmé.',
    ctxTable: {
      head: ['N°', 'Âge', 'Sexe', 'Issue'],
      rows: [
        [1, '3 ans', 'F', 'Vivant'],
        [2, '8 mois', 'M', 'Vivant'],
        [3, '6 ans', 'M', 'Décédé'],
        [4, '2 ans', 'M', 'Décédé'],
        [5, '12 ans', 'F', 'Vivant'],
        [6, '4 ans', 'F', 'Vivant']
      ]
    },
    rptDisease: 'Rougeole',
    o: [
      { rpt: [[2, 2, 1, 1, 6], [1, 0, 1, 0, 2], [0, 0, 0, 0, 0]] },
      { rpt: [[2, 2, 1, 1, 6], [1, 0, 1, 0, 2], [2, 2, 1, 1, 6]] },
      { rpt: [[3, 2, 2, 1, 8], [1, 0, 1, 0, 2], [0, 0, 0, 0, 0]] },
      { rpt: [[2, 2, 1, 1, 6], [0, 1, 0, 1, 2], [0, 0, 0, 0, 0]] }
    ],
    a: 0,
    e: '6 cas suspects : moins de 5 ans, 2 garçons (8 mois, 2 ans) et 2 filles (3 ans, 4 ans) ; 5 ans et plus, 1 garçon (6 ans) et 1 fille (12 ans) ; total 6. Décès : 1 garçon de moins de 5 ans (2 ans) et 1 garçon de 5 ans et plus (6 ans) ; total 2. Confirmé : 0 partout. Les décès font déjà partie des cas suspects : on ne les ajoute pas.'
  },
  {
    type: 'dates',
    q: 'Choisissez les bonnes dates pour remplir le bloc « Histoire de la maladie » de la fiche IDS.',
    ctx: 'Un enfant de 3 ans, né le 12/05/2023, présente de la fièvre et une éruption cutanée à partir du 02/11/2026. Ses parents le conduisent au centre de santé le 04/11/2026, où l’infirmier suspecte une rougeole. Le 05/11/2026, la formation sanitaire envoie la fiche de notification à la zone sanitaire. Le prélèvement est envoyé au laboratoire le 06/11/2026.',
    dates: [
      { k: 'debut', label: 'Date de début des symptômes', o: ['12/05/2023', '02/11/2026', '04/11/2026', '05/11/2026', '06/11/2026'], a: 1 },
      { k: 'consult', label: 'Date de consultation', o: ['12/05/2023', '02/11/2026', '04/11/2026', '05/11/2026', '06/11/2026'], a: 2 },
      { k: 'notif', label: 'Date de notification', o: ['12/05/2023', '02/11/2026', '04/11/2026', '05/11/2026', '06/11/2026'], a: 3 }
    ],
    e: 'Début des symptômes : 02/11/2026. Consultation : 04/11/2026. Notification : 05/11/2026, c’est la date à laquelle la formation sanitaire envoie la fiche de notification à la zone sanitaire. Le 12/05/2023 est la date de naissance et le 06/11/2026 la date d’envoi du prélèvement au laboratoire. Les dates doivent se suivre : début, consultation, notification.'
  },
  {
    type: 'multi',
    q: 'Quels blocs de la fiche IDS sont remplis par la formation sanitaire ?',
    o: [
      'Le bloc 1 : en-tête',
      'Le bloc 2 : identification du cas',
      'Le bloc 3 : histoire de la maladie',
      'Le bloc 4 : suivi du cas',
      'Le bloc 5 : prélèvement',
      'Le bloc 6 : laboratoire'
    ],
    a: [0, 1, 2, 3, 4],
    e: 'Les blocs 1 à 5 sont remplis par la formation sanitaire. Le bloc 6 (laboratoire) est rempli par le laboratoire et reste vide à la formation sanitaire.'
  },

  // ---------- Tracé de la courbe ----------
  {
    q: 'Quelle est l’utilité essentielle du tracé de la courbe épidémique ?',
    o: [
      'Suivre l’évolution des cas et se repérer dans le temps',
      'Remplacer la fiche IDS',
      'Confirmer les cas au laboratoire',
      'Calculer la couverture vaccinale'
    ],
    a: 0,
    e: 'La courbe épidémique permet de suivre l’évolution des cas semaine après semaine et de se repérer dans le temps : on voit si le nombre de cas augmente ou baisse.'
  },
  {
    q: 'Quel dessin correspond au tracé de la courbe épidémique ?',
    chartData: [0, 1, 0, 2, 3, 5, 8, 11, 7, 4, 2, 1],
    o: [
      { chart: 'pie', t: 'Camembert' },
      { chart: 'line', t: 'Diagramme en ligne' },
      { chart: 'curve', t: 'Courbe' },
      { chart: 'histo', t: 'Histogramme (colonnes collées)' },
      { chart: 'bars', t: 'Diagramme en barres' }
    ],
    a: 3,
    e: 'La courbe épidémique se trace en histogramme : une colonne par semaine, les colonnes collées les unes aux autres, chacune remplie jusqu’à la hauteur du nombre de cas (1 carreau = 1 cas). Ce n’est ni un camembert, ni une ligne ou une courbe lissée, ni un diagramme en barres dont les barres sont séparées.'
  },
  {
    q: 'En semaine 8, le rapport hebdomadaire indique 11 cas suspects. Que fait-on sur la courbe épidémique ?',
    o: [
      'On remplit la colonne de la semaine 8 jusqu’à 11 carreaux',
      'On remplit la colonne 11 jusqu’à 8 carreaux',
      'On remplit la colonne de la semaine 8 jusqu’à 8 carreaux',
      'On ne trace rien tant que les cas ne sont pas confirmés'
    ],
    a: 0,
    e: 'On repère la colonne de la semaine (8) et on la remplit jusqu’à la hauteur du nombre de cas : 1 carreau = 1 cas, donc 11 carreaux. On trace à partir du total des cas suspects.'
  },
  {
    type: 'multi',
    fixed: true,
    q: 'Pour quels événements prioritaires faut-il tracer la courbe épidémique ?',
    o: [
      'Anthrax',
      'Choléra',
      'Décès de 5 ans et plus',
      'Décès infantile',
      'Décès infanto-juvénile',
      'Décès maternel',
      'Décès néonatals',
      'Dengue',
      'Dracunculose',
      'Fièvre de Lassa',
      'Fièvre jaune',
      'Fièvre typhoïde',
      'Grippe HP',
      'Hépatites',
      'Intoxication',
      'Lèpre',
      'Méningite',
      'Morsure de chien',
      'Mpox',
      'Near miss',
      'Noma',
      'Paludisme grave',
      'Paludisme simple',
      'Peste',
      'PFA',
      'Pian',
      'Rage',
      'Rougeole',
      'Shigellose',
      'Syphilis endémique',
      'Tétanos néonatal',
      'Ulcère de Buruli',
      'Zika'
    ],
    a: [1, 2, 5, 6, 7, 9, 10, 16, 17, 18, 24, 26, 27, 28, 30],
    e: 'Les événements prioritaires pour lesquels on trace la courbe épidémique sont : choléra, méningite, fièvre jaune, tétanos néonatal, PFA, rougeole, shigellose, rage, fièvre de Lassa, dengue, mpox, morsure de chien, décès maternel, décès néonatals et décès de 5 ans et plus.'
  }
];

/* ------------------------------------------------------------------ */
/* API (appelée par la page hébergée sur GitHub)                       */
/* ------------------------------------------------------------------ */
const ACTIONS = {
  getPublicConfig: getPublicConfig,
  startTest: startTest,
  submitTest: submitTest,
  adminLogin: adminLogin,
  adminSaveConfig: adminSaveConfig,
  adminStats: adminStats
};

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = ACTIONS.hasOwnProperty(body.action) ? ACTIONS[body.action] : null;
    if (!fn) return json_({ ok: false, message: 'Action inconnue.' });
    const args = Array.isArray(body.args) ? body.args : [];
    return json_(fn.apply(null, args));
  } catch (err) {
    return json_({ ok: false, message: 'Erreur du serveur : ' + err.message });
  }
}

/** Test rapide : ouvrir l'URL /exec dans un navigateur doit afficher {"ok":true,...}. */
function doGet() {
  return json_({ ok: true, service: 'pretest-posttest', message: 'API en ligne. La page du test est hébergée séparément.' });
}

/** À exécuter une fois depuis l'éditeur (autorise l'accès au Sheet et crée les onglets). */
function setup() {
  ensureSheets_();
}

/* ------------------------------------------------------------------ */
/* CONFIGURATION DE L'ACTIVITÉ (pilotée par l'administrateur)          */
/* ------------------------------------------------------------------ */
function getConfig_() {
  const raw = PropertiesService.getScriptProperties().getProperty('CONFIG');
  if (!raw) return { open: false, title: '', session: '', phase: 'Pré-test', stamp: '' };
  return JSON.parse(raw);
}

/** Ce que voient les participants : seulement l'état ouvert/fermé et le titre. */
function getPublicConfig() {
  const c = getConfig_();
  return { open: !!c.open && !!c.title, title: c.title || '' };
}

function checkAdmin_(pw) {
  if (ADMIN_PASSWORD === 'CHANGEZ-MOI') {
    throw new Error('Mot de passe administrateur non défini : modifiez ADMIN_PASSWORD dans Code.gs puis redéployez.');
  }
  if (String(pw) !== ADMIN_PASSWORD) throw new Error('Mot de passe incorrect.');
}

function adminLogin(pw) {
  try {
    checkAdmin_(pw);
    return { ok: true, config: getConfig_() };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

function adminSaveConfig(pw, cfg) {
  try {
    checkAdmin_(pw);
    const title = String((cfg && cfg.title) || '').trim().slice(0, 100);
    const session = String((cfg && cfg.session) || '').trim().slice(0, 60);
    const phase = cfg && cfg.phase;
    if (!title) return { ok: false, message: 'Saisissez un titre d’activité.' };
    if (!session) return { ok: false, message: 'Saisissez la session / classe (ex. « Vague 1 – chefs de poste »).' };
    if (PHASES.indexOf(phase) < 0) return { ok: false, message: 'Choisissez Pré-test ou Post-test.' };

    const old = getConfig_();
    const changed = old.title !== title || old.session !== session || old.phase !== phase;
    const config = {
      open: !!(cfg && cfg.open),
      title: title,
      session: session,
      phase: phase,
      stamp: (changed || !old.stamp) ? String(Date.now()) : old.stamp
    };
    PropertiesService.getScriptProperties().setProperty('CONFIG', JSON.stringify(config));
    return { ok: true, config: config };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/* ------------------------------------------------------------------ */
/* STATISTIQUES (admin)                                                */
/* ------------------------------------------------------------------ */
function stat_(arr) {
  if (!arr.length) return null;
  let sum = 0, min = arr[0], max = arr[0];
  arr.forEach(function (v) { sum += v; if (v < min) min = v; if (v > max) max = v; });
  return { n: arr.length, min: min, max: max, mean: Math.round(sum / arr.length * 100) / 100 };
}

function adminStats(pw) {
  try {
    checkAdmin_(pw);
    const ss = ensureSheets_();
    const sh = ss.getSheetByName('Resultats');
    const last = sh.getLastRow();
    const groups = {};
    const keys = [];
    if (last >= 2) {
      const vals = sh.getRange(2, 1, last - 1, COL.DUR + 1).getValues();
      vals.forEach(function (r) {
        const s = String(r[COL.SESSION]);
        const ph = r[COL.PHASE];
        const id = String(r[COL.ID]);
        const score = Number(r[COL.SCORE]);
        if (!groups[s]) { groups[s] = { session: s, pre: [], post: [], byId: {} }; keys.push(s); }
        const g = groups[s];
        if (!g.byId[id]) g.byId[id] = { id: id };
        if (ph === 'Pré-test') { g.pre.push(score); g.byId[id].pre = score; }
        else if (ph === 'Post-test') { g.post.push(score); g.byId[id].post = score; }
      });
    }
    const sessions = keys.map(function (k) {
      const g = groups[k];
      const rows = Object.keys(g.byId).sort().map(function (id) {
        const x = g.byId[id];
        const hasBoth = typeof x.pre === 'number' && typeof x.post === 'number';
        return { id: id, pre: typeof x.pre === 'number' ? x.pre : null, post: typeof x.post === 'number' ? x.post : null, gain: hasBoth ? x.post - x.pre : null };
      });
      const gains = rows.filter(function (r) { return r.gain !== null; }).map(function (r) { return r.gain; });
      let gainMean = null;
      if (gains.length) gainMean = Math.round(gains.reduce(function (a, b) { return a + b; }, 0) / gains.length * 100) / 100;
      return { session: k, pre: stat_(g.pre), post: stat_(g.post), pairedN: gains.length, gainMean: gainMean, rows: rows };
    });
    return { ok: true, total: QUESTIONS.length, sessions: sessions, config: getConfig_() };
  } catch (err) {
    return { ok: false, message: err.message };
  }
}

/* ------------------------------------------------------------------ */
/* API APPELÉE PAR LES PARTICIPANTS                                    */
/* ------------------------------------------------------------------ */
function startTest(rawId) {
  try {
    const c = getConfig_();
    if (!c.open || !c.title) return { ok: false, message: 'Aucune activité n’est ouverte pour le moment.' };
    const id = normId_(rawId);
    if (!id) {
      return { ok: false, message: 'Identifiant invalide : 2 à 30 caractères (lettres, avec ou sans chiffres ; tiret ou _ acceptés).' };
    }
    const ss = ensureSheets_();
    if (alreadyDone_(ss, id, c.session, c.phase)) {
      return {
        ok: false,
        message: 'L’identifiant ' + id + ' a déjà passé cette activité. Chaque identifiant ne peut la remplir qu’une seule fois. ' +
                 'S’il s’agit d’une erreur, prévenez l’animateur.'
      };
    }
    if (c.phase === 'Post-test' && !alreadyDone_(ss, id, c.session, 'Pré-test')) {
      return { ok: false, message: NO_PRE_MSG_(id) };
    }
    const order = buildOrder_(id, c.session + '|' + c.phase);
    return { ok: true, id: id, title: c.title, stamp: c.stamp, limitSec: TIME_LIMIT_SEC, items: order.map(publicItem_) };
  } catch (err) {
    return { ok: false, message: 'Erreur du serveur : ' + err.message };
  }
}

function submitTest(p) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const c = getConfig_();
    if (!p || !c.stamp || String(p.stamp) !== c.stamp) {
      return { ok: false, message: 'L’administrateur a changé l’activité pendant votre test : vos réponses n’ont pas été enregistrées. Prévenez l’animateur.' };
    }
    const id = normId_(p.id);
    if (!id) return { ok: false, message: 'Identifiant invalide.' };

    const ss = ensureSheets_();
    if (alreadyDone_(ss, id, c.session, c.phase)) {
      return { ok: false, message: 'Cette activité a déjà été soumise pour l’identifiant ' + id + '.' };
    }
    if (c.phase === 'Post-test' && !alreadyDone_(ss, id, c.session, 'Pré-test')) {
      return { ok: false, message: NO_PRE_MSG_(id) };
    }

    const answers = p.answers || {};
    const forced = p.timeout === true; // envoi automatique à la fin des 15 minutes
    const order = buildOrder_(id, c.session + '|' + c.phase);
    if (!forced) {
      for (let i = 0; i < order.length; i++) {
        if (!isAnswered_(order[i], answers[order[i].k])) {
          return { ok: false, message: 'Il reste des questions sans réponse.' };
        }
      }
    }

    let score = 0;
    let unanswered = 0;
    const okByQuestion = {};
    const items = order.map(function (it, idx) {
      const v = answers[it.k];
      const answered = isAnswered_(it, v);
      const ok = answered && isCorrect_(it, v);
      if (!answered) unanswered++;
      if (ok) score++;
      okByQuestion[it.k] = ok ? 1 : 0;
      const q = it.def;
      return {
        n: idx + 1,
        type: it.type,
        q: q.q,
        ctx: q.ctx || null,
        ctxTable: q.ctxTable || null,
        rptDisease: q.rptDisease || null,
        chartData: q.chartData || null,
        chosen: answered ? chosenList_(it, v) : [{ t: 'Pas de réponse (compte comme fausse)' }],
        correct: correctList_(it),
        ok: ok,
        e: q.e
      };
    });

    const total = order.length;
    const pct = Math.round(score * 100 / total);
    const duration = forced ? TIME_LIMIT_SEC : Math.max(0, Math.round(Number(p.duration) || 0));
    const row = [new Date(), id, c.session, c.phase, c.title, score, pct, duration,
                 order.map(function (o) { return o.k; }).join('>')];
    for (let q = 1; q <= total; q++) row.push(okByQuestion['Q' + q]);
    row.push(forced ? 'Oui' : 'Non', unanswered);

    ss.getSheetByName('Resultats').appendRow(row);
    SpreadsheetApp.flush();

    return { ok: true, id: id, title: c.title, phase: c.phase, score: score, total: total, pct: pct, timedOut: forced, unanswered: unanswered, items: items };
  } catch (err) {
    return { ok: false, message: 'Enregistrement impossible : ' + err.message + ' — réessayez.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ------------------------------------------------------------------ */
/* OUTILS INTERNES                                                     */
/* ------------------------------------------------------------------ */
function NO_PRE_MSG_(id) {
  return 'L’identifiant ' + id + ' n’a pas été trouvé au pré-test de cette session. ' +
         'Pour passer le post-test, utilisez exactement le même identifiant qu’au pré-test. ' +
         'Si vous l’avez oublié, demandez-le à l’animateur.';
}

function normId_(raw) {
  const id = String(raw == null ? '' : raw).trim().toUpperCase().replace(/\s+/g, '');
  return /^[\p{L}0-9_-]{2,30}$/u.test(id) ? id : '';
}

/** Ordre propre à chaque couple (identifiant, activité) : stable si la page est rechargée, différent d'une personne à l'autre. */
function buildOrder_(id, activityKey) {
  const rnd = rng_(id + '|' + activityKey);
  const qs = QUESTIONS.map(function (q, i) { return { i: i, q: q }; });
  shuffle_(qs, rnd);
  return qs.map(function (x) {
    const q = x.q;
    const it = { k: 'Q' + (x.i + 1), type: q.type || 'single', def: q };
    if (it.type !== 'dates' && it.type !== 'match') {
      const opts = q.o.map(function (o, c) { return Object.assign({ c: c }, optObj_(o)); });
      if (!q.fixed) shuffle_(opts, rnd); // liste longue : ordre alphabétique conservé
      it.opts = opts;
    }
    return it;
  });
}

function optObj_(o) { return typeof o === 'string' ? { t: o } : o; }

/** Ce qui est envoyé au participant : jamais la bonne réponse ni l'explication. */
function publicItem_(it) {
  const q = it.def;
  const out = { k: it.k, type: it.type, q: q.q };
  if (q.ctx) out.ctx = q.ctx;
  if (q.ctxTable) out.ctxTable = q.ctxTable;
  if (q.rptDisease) out.rptDisease = q.rptDisease;
  if (q.chartData) out.chartData = q.chartData;
  if (it.type === 'match') {
    out.rows = q.rows.map(function (r) { return { label: r.label }; });
    out.pre = q.pre.map(function (t, c) { return { c: c, t: t }; });
    out.tubes = q.tubes.map(function (x, c) { return { c: c, id: x.id, t: x.t }; });
  } else if (it.type === 'dates') {
    out.dates = q.dates.map(function (f) {
      return { k: f.k, label: f.label, options: f.o.map(function (t, c) { return { c: c, t: t }; }) };
    });
  } else {
    out.options = it.opts;
  }
  return out;
}

function inRange_(v, n) { return typeof v === 'number' && isFinite(v) && Math.floor(v) === v && v >= 0 && v < n; }

function isAnswered_(it, v) {
  const q = it.def;
  if (it.type === 'match') {
    if (!v || typeof v !== 'object') return false;
    return q.rows.every(function (r, i) {
      const x = v[i];
      return !!x && x.p !== null && x.p !== '' && x.t !== null && x.t !== '' &&
             inRange_(Number(x.p), q.pre.length) && inRange_(Number(x.t), q.tubes.length);
    });
  }
  if (it.type === 'dates') {
    if (!v || typeof v !== 'object') return false;
    return q.dates.every(function (f) { return inRange_(Number(v[f.k]), f.o.length) && v[f.k] !== '' && v[f.k] !== null; });
  }
  if (it.type === 'multi') {
    return Array.isArray(v) && v.length > 0 && v.every(function (x) { return inRange_(Number(x), q.o.length); });
  }
  return v !== undefined && v !== null && v !== '' && inRange_(Number(v), q.o.length);
}

function uniqSorted_(v) {
  const seen = {};
  const out = [];
  (Array.isArray(v) ? v : []).forEach(function (x) { const n = Number(x); if (!seen[n]) { seen[n] = true; out.push(n); } });
  return out.sort(function (a, b) { return a - b; });
}

/** Tout ou rien pour les choix multiples et pour les trois dates. */
function isCorrect_(it, v) {
  const q = it.def;
  if (it.type === 'match') {
    return q.rows.every(function (r, i) { return Number(v[i].p) === r.p && Number(v[i].t) === r.t; });
  }
  if (it.type === 'multi') {
    const got = uniqSorted_(v);
    const exp = q.a.slice().sort(function (a, b) { return a - b; });
    return got.length === exp.length && got.every(function (x, i) { return x === exp[i]; });
  }
  if (it.type === 'dates') {
    return q.dates.every(function (f) { return Number(v[f.k]) === f.a; });
  }
  return Number(v) === q.a;
}

function chosenList_(it, v) {
  const q = it.def;
  if (it.type === 'match') {
    return q.rows.map(function (r, i) { return { t: r.label + ' : ' + q.pre[Number(v[i].p)] + ' · ' + q.tubes[Number(v[i].t)].t }; });
  }
  if (it.type === 'dates') return q.dates.map(function (f) { return { t: f.label + ' : ' + f.o[Number(v[f.k])] }; });
  if (it.type === 'multi') return uniqSorted_(v).map(function (c) { return optObj_(q.o[c]); });
  return [optObj_(q.o[Number(v)])];
}

function correctList_(it) {
  const q = it.def;
  if (it.type === 'match') {
    return q.rows.map(function (r) { return { t: r.label + ' : ' + q.pre[r.p] + ' · ' + q.tubes[r.t].t }; });
  }
  if (it.type === 'dates') return q.dates.map(function (f) { return { t: f.label + ' : ' + f.o[f.a] }; });
  if (it.type === 'multi') return q.a.map(function (c) { return optObj_(q.o[c]); });
  return [optObj_(q.o[q.a])];
}

/** Texte de la bonne réponse, pour l'onglet Questions du Sheet. */
function answerText_(q) {
  const it = { def: q, type: q.type || 'single' };
  return correctList_(it).map(function (o) {
    if (o.rpt) {
      return 'Suspect ' + o.rpt[0].join('-') + ' ; Décès ' + o.rpt[1].join('-') + ' ; Confirmé ' + o.rpt[2].join('-') + ' (<5 M, <5 F, ≥5 M, ≥5 F, total)';
    }
    return o.t;
  }).join(' | ');
}

function rng_(seed) {
  let h = 1779033703 ^ seed.length;
  for (let i = 0; i < seed.length; i++) {
    h = Math.imul(h ^ seed.charCodeAt(i), 3432918353);
    h = (h << 13) | (h >>> 19);
  }
  h = Math.imul(h ^ (h >>> 16), 2246822507);
  h = Math.imul(h ^ (h >>> 13), 3266489909);
  let a = (h ^ (h >>> 16)) >>> 0;
  return function () {
    a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffle_(arr, rnd) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(rnd() * (i + 1));
    const tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
  }
  return arr;
}

function alreadyDone_(ss, id, session, phase) {
  const sh = ss.getSheetByName('Resultats');
  const last = sh.getLastRow();
  if (last < 2) return false;
  const vals = sh.getRange(2, 2, last - 1, 3).getValues(); // B (ID), C (Session), D (Phase)
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toUpperCase() === id && String(vals[i][1]) === session && vals[i][2] === phase) return true;
  }
  return false;
}

function ensureSheets_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const total = QUESTIONS.length;

  // --- Onglet Resultats : une ligne par soumission ---
  let res = ss.getSheetByName('Resultats');
  if (!res) {
    res = ss.insertSheet('Resultats');
    const head = ['Horodatage', 'ID', 'Session', 'Phase', 'Titre de l’activité', 'Score /' + total, 'Pourcentage', 'Durée (s)', 'Ordre des questions'];
    for (let q = 1; q <= total; q++) head.push('Q' + q);
    head.push('Temps écoulé', 'Sans réponse');
    res.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground('#0f5c5c').setFontColor('#ffffff');
    res.setFrozenRows(1);
    res.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    res.getRange('B:E').setNumberFormat('@'); // ID, session, phase, titre : du texte (ex. 007)
    res.setColumnWidths(2, 4, 150);
    res.setColumnWidth(9, 220);
  }

  // --- Onglet Questions : référence et clé de correction ---
  let qs = ss.getSheetByName('Questions');
  if (!qs) {
    qs = ss.insertSheet('Questions');
    const rows = [['N°', 'Question', 'Bonne réponse', 'Explication']];
    QUESTIONS.forEach(function (q, i) { rows.push([i + 1, q.q, answerText_(q), q.e]); });
    qs.getRange(1, 1, rows.length, 4).setValues(rows);
    qs.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#0f5c5c').setFontColor('#ffffff');
    qs.setFrozenRows(1);
    qs.setColumnWidth(2, 420); qs.setColumnWidth(3, 320); qs.setColumnWidth(4, 420);
    qs.getRange(2, 2, rows.length - 1, 3).setWrap(true);
  }

  return ss;
}
