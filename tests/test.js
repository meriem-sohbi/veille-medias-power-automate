// Banc d'essai local des Office Scripts : faux classeur, données fictives, réponses IA simulées.
// Usage : « npm test » à la racine du dépôt (compile les scripts puis lance ce fichier).
const fs = require('fs'), vm = require('vm'), path = require('path');
const D = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'demo_data.json'), 'utf8'));
const IA = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'ia_simulee.json'), 'utf8'));
const OUT = path.join(__dirname, 'resultats'); fs.mkdirSync(OUT, { recursive: true });
const code = n => fs.readFileSync(path.join(__dirname, 'out', n, n + '.js'), 'utf8');

// Faux classeur : mêmes méthodes que l'API Office Scripts utilisées par les scripts.
function table(headers, rows) {
  const t = { headers, rows: rows.map(r => r.slice()) };
  return {
    _t: t,
    getRowCount: () => t.rows.length,
    getRangeBetweenHeaderAndTotal: () => ({ getValues: () => t.rows.map(r => r.slice()) }),
    deleteRowsAt: (i, n = 1) => { t.rows.splice(i, n); },
    addRows: (i = -1, vals = []) => { vals.forEach(v => { if (v.length !== headers.length) throw new Error('largeur de ligne incorrecte pour la table ' + headers.join(',')); });
      if (i < 0) t.rows.push(...vals.map(v => v.slice())); else t.rows.splice(i, 0, ...vals.map(v => v.slice())); },
  };
}
const ENTETES = {
  Brouillon: ['Garder', 'Categorie', 'Titre', 'Resume', 'A verifier', 'Controles', 'Statut', 'Sources et liens', 'Question reunion', 'Extrait d origine', 'Id', 'Resume IA'],
  Suivi: ['Date', 'Semaine du', 'Valide par', 'Proposes', 'Gardes', 'Resumes modifies', 'Duree relecture (min)', 'Commentaire'],
};
function classeur(audiences) {
  const vide = n => [Array(ENTETES[n].length).fill('')];
  const T = {
    Parametres: table(['Cle', 'Valeur'], [
      ['fenetreJours', '7'], ['maxArticlesIA', '30'],
      ['motsCles', 'audience;prime;grille;streaming;format;serie;plateforme;replay;quiz;regulation;consultation;programme'],
      ['motsExclus', 'people;mariage;divorce'],
      ['concurrents', 'Chaîne Concurrente 1;Chaîne Concurrente 2;Plateforme Stream'],
      ['categories', 'Programmation et concurrence;Audiences et mesures;Formats et tendances;Usages et plateformes;Régulation et juridique;Rumeurs à confirmer;Hors périmètre']]),
    Audiences: table(['Programme', 'Jour', 'Audience (M)', 'Semaine precedente (M)'], audiences),
    Brouillon: table(ENTETES.Brouillon, vide('Brouillon')),
    Contexte: table(['Cle', 'Valeur'], [['', '']]),
    Suivi: table(ENTETES.Suivi, vide('Suivi')),
  };
  return { T, getTable: n => T[n] };
}
const run = (n, ...args) => { const ctx = { args, console }; vm.runInNewContext(code(n) + '\nglobalThis.resultat = main(...args);', ctx); return ctx.resultat; };

const REF = '2026-11-23T05:30:00Z';   // lundi 6 h 30, heure de Paris
const src = Object.fromEntries(D.sources.map(s => [s.id, s]));
const articles = D.articles.map(a => ({ titre: a.titre, url: a.url, extrait: a.extrait,
  date: new Date(Date.parse(REF) - a.age * 86400000).toISOString(), source: src[a.sourceId].nom, fiabilite: src[a.sourceId].fiabilite }));
// Comme dans le flux : un tableau d'articles par source.
const parSource = Object.values(articles.reduce((m, a) => ((m[a.source] = m[a.source] || []).push(a), m), {}));
const echecs = [{ source: 'Écran Lab', erreur: 'indisponible après 3 tentatives' }];
const audiencesDemo = D.internal.programmes.map(p => [p.programme, p.jour, p.audience_m, p.audience_m_s1]);
const canon = u => { const s = String(u).trim().replace(/#.*$/, ''); const m = s.match(/^[a-z]+:\/\/([^\/?]+)([^?]*)\??(.*)$/i); if (!m) return s.toLowerCase();
  const q = m[3].split('&').filter(x => x && !/^(utm_[a-z]+|fbclid|gclid|xtor)=/i.test(x)).join('&'); return m[1].toLowerCase().replace(/^www\./, '') + m[2].replace(/\/+$/, '') + (q ? '?' + q : ''); };
// Ce que renverrait l'action « Exécuter un prompt » (réponses préparées à la main, au format demandé par le prompt).
const iaSimulee = prep => JSON.stringify(prep.candidats.filter(c => IA[canon(c.url)]).map(c => ({ id: c.id, ...IA[canon(c.url)] })));

const bilan = [];
function scenario(nom, { ia, audiences = audiencesDemo, relecture }) {
  const wb = classeur(audiences);
  const prepTxt = run('1_Preparer', wb, JSON.stringify(parSource), REF);
  const prep = JSON.parse(prepTxt);
  const envoye = prep.elementsIA;
  const fuite = /https?:\/\//.test(envoye) || /Duel de Cuisine|Enquêtes Ouvertes/.test(envoye);
  const rep = run('2_Controler', wb, prepTxt, ia(prep), JSON.stringify(echecs));
  fs.writeFileSync(path.join(OUT, nom + '_brouillon.html'), page(rep.html));
  const lignes = wb.T.Brouillon._t.rows;
  const b = { scenario: nom, stats: prep.stats, fuiteVersIA: fuite, proposes: rep.nbProposes, aVerifier: rep.nbAVerifier, iaOk: rep.iaOk,
    alertes: lignes.filter(l => l[5] !== 'OK' || l[6] !== 'normal').map(l => l[0] + ' | ' + l[6] + ' | ' + l[2].slice(0, 40) + ' -> ' + l[5]) };
  if (relecture) {
    relecture(lignes);
    try { run('3_Publier', wb, '', '', 0); b.sansValideur = 'ACCEPTÉ (anomalie)'; } catch (e) { b.sansValideur = 'refusé : ' + e.message; }
    const pub = run('3_Publier', wb, 'Camille Martin', 'RAS', 18);
    fs.writeFileSync(path.join(OUT, nom + '_valide.html'), page(pub.html));
    b.publication = { gardes: pub.nbGardes, modifies: pub.nbModifies, sansResume: pub.nbSansResume, suivi: wb.T.Suivi._t.rows };
  }
  bilan.push(b);
  return { wb, rep };
}
function page(html) { return '<!doctype html><meta charset="utf-8"><body style="font-family:sans-serif;max-width:820px;margin:24px auto;line-height:1.4">' + html + '</body>'; }

scenario('1_nominal', { ia: iaSimulee, relecture: lignes => {
  lignes.find(l => /quiz/i.test(l[2]))[0] = 'Non';                          // la relectrice retire un élément
  lignes.find(l => /policière/.test(l[2]))[3] = 'Chaîne Concurrente 1 lance en janvier une série policière de six épisodes, le mardi soir.'; // et corrige un résumé
} });
scenario('2_ia_illisible', { ia: () => 'Désolé, je ne peux pas répondre au format demandé.' });
scenario('3_ia_invente', { ia: prep => { const a = JSON.parse(iaSimulee(prep)); a[0].resume += ' Selon Jean Duval, 14 épisodes sont prévus.'; return JSON.stringify(a); } });
scenario('4_audiences_absentes', { ia: iaSimulee, audiences: [] });
scenario('5_ia_bloc_json', { ia: prep => '```json\n' + iaSimulee(prep) + '\n```' });
fs.writeFileSync(path.join(OUT, 'bilan.json'), JSON.stringify(bilan, null, 2));

// Entrée d'exemple pour tester le prompt dans AI Builder.
{ const p = JSON.parse(run('1_Preparer', classeur(audiencesDemo), JSON.stringify(parSource), REF));
  fs.writeFileSync(path.join(OUT, 'exemple_elements.json'), JSON.stringify(JSON.parse(p.elementsIA), null, 1)); }
console.log(JSON.stringify(bilan, null, 1));
