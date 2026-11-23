// Office Script 2, « 2_Controler ».
// Reçoit la préparation (script 1) et la réponse de l'IA (action « Exécuter un prompt »).
// Contrôle chaque résumé, calcule les audiences internes sans IA, écrit le brouillon
// dans la table « Brouillon » du classeur partagé (la personne qui relit y corrige et y retire),
// et renvoie un brouillon HTML et un texte court pour la demande d'approbation.

interface Reprise { source: string; fiabilite: string; url: string; titre: string; extrait: string; }
interface Candidat { id: string; titre: string; url: string; source: string; fiabilite: string; date: string; extrait: string; reprises: Reprise[]; divergence: string[]; scorePre: number; }
interface Stats { collectes: number; horsFenetre: number; exclus: number; doublonsFusionnes: number; sujets: number; envoyesIA: number; }
interface Preparation { candidats: Candidat[]; stats: Stats; fenetreDebut: string; fenetreFin: string; elementsIA: string; }
interface ReponseIA { id: string; categorie: string; resume: string; a_verifier: string; question_reunion: string; }
interface Echec { source: string; erreur: string; }
interface Ligne { garder: string; categorie: string; titre: string; resume: string; aVerifier: string; controles: string; statut: string; liens: string; question: string; extrait: string; id: string; resumeIA: string; }
interface Chiffre { programme: string; jour: string; audience: number; precedente: number; evol: number; }
interface Resultat { html: string; resumeApprobation: string; nbProposes: number; nbAVerifier: number; nbSourcesEnEchec: number; iaOk: boolean; }

const CATEGORIES_DEFAUT = 'Programmation et concurrence;Audiences et mesures;Formats et tendances;Usages et plateformes;Régulation et juridique;Rumeurs à confirmer;Hors périmètre';
const INJECTION = /(ignore[rz]?\s+(les\s+|tes\s+|toutes\s+)?(instructions|consignes)|ignore (all|previous)|system prompt|attribue la note|tu dois maintenant)/i;

function main(workbook: ExcelScript.Workbook, preparationJson: string, reponseIA: string, sourcesEnEchecJson: string): Resultat {
  const prep = JSON.parse(preparationJson) as Preparation;
  const categories = lireParam(workbook, 'categories', CATEGORIES_DEFAUT).split(';').map(x => x.trim()).filter(x => x.length > 0);
  const echecs = lireEchecs(sourcesEnEchecJson);

  // 1. Lecture prudente de la réponse IA. Si elle est illisible, on continue sans elle.
  let reponses: ReponseIA[] = [];
  let erreurIA = '';
  try {
    const propre = String(reponseIA || '').trim().replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
    if (!propre) throw new Error('réponse vide');
    const brut = JSON.parse(propre) as ReponseIA[] | { elements: ReponseIA[] };
    const liste: ReponseIA[] = Array.isArray(brut) ? brut : (brut && Array.isArray(brut.elements) ? brut.elements : []);
    if (!Array.isArray(brut) && liste.length === 0) throw new Error('JSON valide mais sans tableau');
    reponses = liste;
  } catch (e) {
    erreurIA = String(e).slice(0, 120);
    reponses = [];
  }

  // 2. Contrôles automatiques, élément par élément.
  const lignes: Ligne[] = prep.candidats.map(c => controler(c, reponses.find(r => r && r.id === c.id), categories));

  // 3. Alertes de collecte et bilan.
  const alertes: string[] = [];
  echecs.forEach(e => alertes.push('La source ' + e.source + ' n’a pas pu être collectée (' + e.erreur + ').'));
  const nbEcartes = lignes.filter(l => l.statut === 'écarté').length;
  if (nbEcartes) alertes.push(nbEcartes + (nbEcartes > 1 ? ' articles contenant une consigne cachée adressée à l’IA ont été écartés.' : ' article contenant une consigne cachée adressée à l’IA a été écarté.'));
  if (erreurIA) alertes.push('Réponse IA inexploitable (' + erreurIA + '). Les résumés sont à rédiger à la main à partir des extraits.');
  const s = prep.stats;
  const bilan = 'Collecte, ' + s.collectes + ' articles, ' + s.horsFenetre + ' hors période, ' + s.exclus + ' exclus par mots-clés, ' + s.doublonsFusionnes + ' doublons regroupés, ' + s.sujets + ' sujets analysés.';

  // 4. Audiences internes, calculées sans IA.
  const chiffres = lireAudiences(workbook);

  // 5. Écriture du brouillon dans le classeur (ordre du pré-score, éléments retirés en bas).
  lignes.sort((a, b) => (a.garder === 'Oui' ? 0 : 1) - (b.garder === 'Oui' ? 0 : 1));
  ecrireTable(workbook, 'Brouillon', lignes.map(l => [l.garder, l.categorie, l.titre, l.resume, l.aVerifier, l.controles, l.statut, l.liens, l.question, l.extrait, l.id, l.resumeIA]));
  ecrireTable(workbook, 'Contexte', [
    ['fenetreDebut', prep.fenetreDebut], ['fenetreFin', prep.fenetreFin],
    ['alertes', alertes.join(' || ')], ['bilan', bilan], ['chiffres', JSON.stringify(chiffres)]
  ]);

  const html = rendre(lignes, chiffres, alertes, bilan, prep.fenetreDebut, prep.fenetreFin, '');
  const nbAVerifier = lignes.filter(l => l.garder === 'Oui' && (l.statut !== 'normal' || l.controles !== 'OK')).length;
  const proposes = lignes.filter(l => l.garder === 'Oui');
  const resumeApprobation = 'Veille médias, semaine ' + periode(prep.fenetreDebut, prep.fenetreFin) + '. '
    + proposes.length + ' éléments proposés, dont ' + nbAVerifier + ' à vérifier' + (echecs.length ? ', ' + echecs.length + ' source(s) non collectée(s)' : '') + '. '
    + 'Ouvrez la table Brouillon du classeur, corrigez la colonne Resume si besoin, mettez Non dans Garder pour retirer un élément, puis approuvez.';
  return { html: html, resumeApprobation: resumeApprobation, nbProposes: proposes.length, nbAVerifier: nbAVerifier, nbSourcesEnEchec: echecs.length, iaOk: erreurIA === '' };
}

function controler(c: Candidat, r: ReponseIA | undefined, categories: string[]): Ligne {
  const motifs: string[] = [];
  const corpus = [c.titre, c.extrait, c.source].concat(c.reprises.map(x => x.titre + ' ' + x.extrait + ' ' + x.source)).join(' ');
  const suspect = INJECTION.test(corpus);
  const uneSourceFaible = c.fiabilite === 'C' && c.reprises.length === 0;
  let statut = 'normal';
  if (c.divergence.length) { statut = 'à arbitrer'; motifs.push('chiffres différents selon les sources (' + c.divergence.join(' ; ') + ')'); }
  if (uneSourceFaible) { statut = 'à confirmer'; motifs.push('une seule source peu fiable'); }
  if (suspect) { statut = 'écarté'; motifs.push('consigne cachée adressée à l’IA'); }

  let categorie = 'À classer'; let resume = ''; let aVerifier = ''; let question = '';
  if (!r) {
    motifs.push('aucune réponse IA, résumé à rédiger à partir de l’extrait');
  } else {
    categorie = categories.indexOf(String(r.categorie)) >= 0 ? String(r.categorie) : 'À classer';
    resume = String(r.resume || '').trim();
    const av = String(r.a_verifier || '').trim();
    aVerifier = (av && !/^aucun/i.test(av)) ? av : '';
    question = String(r.question_reunion || '').trim();
    if (/^sans objet/i.test(question)) question = '';
    if (resume.length > 400) motifs.push('résumé trop long');
    // Chaque chiffre du résumé doit exister dans la source.
    const cn = corpus.replace(/(\d),(\d)/g, '$1.$2');
    const nombres = resume.replace(/(\d),(\d)/g, '$1.$2').match(/\d+(?:\.\d+)?/g) || [];
    nombres.forEach(n => {
      const re = new RegExp('(^|[^0-9.])' + n.replace('.', '\\.') + '($|[^0-9])');
      if (!re.test(cn)) motifs.push('chiffre absent de la source, ' + n);
    });
    // Chaque nom propre (hors début de phrase) doit exister dans la source.
    const nc = ' ' + norm(corpus) + ' ';
    resume.replace(/([.!?])\s+/g, '$1\n').split('\n').forEach(phrase => {
      phrase.split(/\s+/).slice(1).forEach(w => {
        const mot = w.replace(/[^A-Za-zÀ-ÿ'’-]/g, '');
        if (/^[A-ZÀ-ÖØ-Þ][a-zà-öø-ÿ'’-]{3,}$/.test(mot) && nc.indexOf(' ' + norm(mot) + ' ') < 0) motifs.push('nom propre absent de la source, ' + mot);
      });
    });
  }
  if (uneSourceFaible && categorie !== 'Hors périmètre') categorie = 'Rumeurs à confirmer';
  const garder = (suspect || categorie === 'Hors périmètre') ? 'Non' : 'Oui';
  const liens = [c.source + ' (' + c.fiabilite + ') ' + c.url].concat(c.reprises.map(x => x.source + ' (' + x.fiabilite + ') ' + x.url)).join(' | ');
  return { garder: garder, categorie: categorie, titre: c.titre, resume: resume, aVerifier: aVerifier, controles: motifs.length ? motifs.join(' ; ') : 'OK',
    statut: statut, liens: liens, question: question, extrait: c.extrait, id: c.id, resumeIA: resume };
}

function lireParam(workbook: ExcelScript.Workbook, cle: string, defaut: string): string {
  const t = workbook.getTable('Parametres');
  if (!t) return defaut;
  const l = t.getRangeBetweenHeaderAndTotal().getValues().find(x => String(x[0]).trim() === cle);
  return l ? String(l[1]) : defaut;
}
function lireEchecs(json: string): Echec[] {
  try { const e = JSON.parse(json || '[]') as Echec[]; return Array.isArray(e) ? e : []; } catch (x) { return []; }
}
function lireAudiences(workbook: ExcelScript.Workbook): Chiffre[] {
  const t = workbook.getTable('Audiences');
  if (!t || t.getRowCount() === 0) return [];
  const res: Chiffre[] = [];
  t.getRangeBetweenHeaderAndTotal().getValues().forEach(l => {
    const a = Number(l[2]); const p = Number(l[3]);
    if (String(l[0]).trim() && a > 0 && p > 0) res.push({ programme: String(l[0]), jour: String(l[1]), audience: a, precedente: p, evol: (a - p) / p * 100 });
  });
  return res.sort((x, y) => y.evol - x.evol);
}
function ecrireTable(workbook: ExcelScript.Workbook, nom: string, valeurs: (string | number | boolean)[][]): void {
  const t = workbook.getTable(nom);
  if (!t) throw new Error('Table ' + nom + ' introuvable dans le classeur');
  const n = t.getRowCount();
  if (n > 0) t.deleteRowsAt(0, n);
  if (valeurs.length) t.addRows(-1, valeurs);
}

// ---------- Rendu HTML, identique dans le script 3 ----------
// Même structure que l'exemple du rapport : À retenir, catégories, audiences internes, À vérifier.
function rendre(lignes: Ligne[], chiffres: Chiffre[], alertes: string[], bilan: string, debut: string, fin: string, validePar: string): string {
  const H: string[] = [];
  const gardees = lignes.filter(l => l.garder === 'Oui' && l.resume);
  const fiables = gardees.filter(l => l.statut === 'normal');
  const douteuses = gardees.filter(l => l.statut !== 'normal');
  const aRediger = lignes.filter(l => l.garder === 'Oui' && !l.resume);
  H.push('<h2>Veille médias, semaine ' + periode(debut, fin) + '</h2>');
  H.push('<p>' + (validePar ? 'Validée par ' + esc(validePar) + '.' : '<b>BROUILLON À VALIDER, ne pas diffuser.</b>') + '</p>');
  if (fiables.length) H.push('<h3>À retenir</h3><ul>' + fiables.slice(0, 3).map(l => '<li>' + esc(l.titre) + '</li>').join('') + '</ul>');
  const cats: string[] = [];
  fiables.forEach(l => { if (cats.indexOf(l.categorie) < 0) cats.push(l.categorie); });
  cats.forEach(cat => {
    H.push('<h3>' + esc(cat) + '</h3>');
    fiables.filter(l => l.categorie === cat).forEach(l => {
      H.push('<p><b>' + esc(l.titre) + '</b><br>' + esc(l.resume)
        + (l.question ? '<br>Question pour la réunion : ' + esc(l.question) : '')
        + (l.aVerifier ? '<br>À vérifier : ' + esc(l.aVerifier) : '')
        + '<br><i>Sources : ' + sources(l.liens) + '</i>'
        + (l.controles !== 'OK' && !validePar ? '<br><span style="color:#B45309">Contrôle automatique : ' + esc(l.controles) + '</span>' : '') + '</p>');
    });
  });
  H.push('<h3>Audiences internes, calcul automatique sans IA</h3>');
  if (chiffres.length) {
    H.push('<table><tr><th>Programme</th><th>Jour</th><th>Audience (M)</th><th>Semaine précédente (M)</th><th>Évolution</th></tr>'
      + chiffres.map(c => '<tr><td>' + esc(c.programme) + '</td><td>' + esc(c.jour) + '</td><td>' + fr(c.audience, 2) + '</td><td>' + fr(c.precedente, 2) + '</td><td>' + (c.evol >= 0 ? '+' : '') + fr(c.evol, 1) + ' %</td></tr>').join('') + '</table>');
  } else {
    H.push('<p>Indisponibles cette semaine. Aucun ancien chiffre n’est affiché à la place.</p>');
  }
  const verif: string[] = douteuses.map(l => '<b>' + esc(l.titre) + '</b> (' + esc(l.statut) + '). ' + esc(l.resume)
    + (l.aVerifier ? ' ' + esc(l.aVerifier) : '') + ' <i>Sources : ' + sources(l.liens) + '</i>');
  aRediger.forEach(l => verif.push('<b>' + esc(l.titre) + '</b>, résumé à rédiger à la main. Extrait : ' + esc(l.extrait.slice(0, 250))));
  alertes.forEach(a => verif.push(esc(a)));
  if (verif.length) H.push('<h3>À vérifier</h3><ul>' + verif.map(v => '<li>' + v + '</li>').join('') + '</ul>');
  if (!validePar) H.push('<p><i>' + esc(bilan) + '</i></p>');
  H.push('<p><i>Résumés proposés par une IA à partir des seuls titres et extraits collectés, contrôlés automatiquement'
    + (validePar ? ', puis relus et validés par une personne' : '') + '. Aucune donnée interne n’est envoyée à l’IA.</i></p>');
  return H.join('\n');
}
function sources(liens: string): string {
  return liens.split(' | ').map(x => { const i = x.lastIndexOf(' '); return '<a href="' + esc(x.slice(i + 1)) + '">' + esc(x.slice(0, i)) + '</a>'; }).join(', ');
}
function esc(s: string): string { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function fr(n: number, k: number): string { return Number(n).toFixed(k).replace('.', ','); }
function periode(debut: string, fin: string): string {
  const d = new Date(debut); const f = new Date(fin);
  if (d.getUTCFullYear() === f.getUTCFullYear() && d.getUTCMonth() === f.getUTCMonth()) return 'du ' + d.getUTCDate() + ' au ' + dateFr(fin);
  return 'du ' + dateFr(debut) + ' au ' + dateFr(fin);
}
function dateFr(iso: string): string {
  const mois = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  const d = new Date(iso);
  return d.getUTCDate() + ' ' + mois[d.getUTCMonth()] + ' ' + d.getUTCFullYear();
}
function norm(s: string): string {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
}
