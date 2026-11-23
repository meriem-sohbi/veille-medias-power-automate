// Office Script 3, « 3_Publier ».
// Appelé après l'approbation. Lit la table « Brouillon » telle que la personne l'a corrigée,
// garde les lignes à « Oui », produit la version validée et ajoute une ligne au suivi du pilote.
// Sans nom de la personne qui valide, rien n'est produit.

interface Ligne { garder: string; categorie: string; titre: string; resume: string; aVerifier: string; controles: string; statut: string; liens: string; question: string; extrait: string; id: string; resumeIA: string; }
interface Chiffre { programme: string; jour: string; audience: number; precedente: number; evol: number; }
interface Publication { html: string; nbGardes: number; nbModifies: number; nbSansResume: number; }

function main(workbook: ExcelScript.Workbook, validePar: string, commentaire: string, dureeRelectureMin: number): Publication {
  if (!validePar || !String(validePar).trim()) throw new Error('Validation refusée, le nom de la personne qui valide est obligatoire');
  const t = workbook.getTable('Brouillon');
  if (!t) throw new Error('Table Brouillon introuvable');
  const lignes: Ligne[] = t.getRowCount() === 0 ? [] : t.getRangeBetweenHeaderAndTotal().getValues()
    .filter(v => String(v[2]).trim() !== '')
    .map(v => ({
      garder: String(v[0]).trim().toLowerCase() === 'oui' ? 'Oui' : 'Non', categorie: String(v[1]), titre: String(v[2]), resume: String(v[3]).trim(),
      aVerifier: String(v[4]).trim(), controles: String(v[5]), statut: String(v[6]) || 'normal', liens: String(v[7]), question: String(v[8]).trim(),
      extrait: String(v[9]), id: String(v[10]), resumeIA: String(v[11])
    }));
  const ctx: { [cle: string]: string } = {};
  const tc = workbook.getTable('Contexte');
  if (tc && tc.getRowCount() > 0) tc.getRangeBetweenHeaderAndTotal().getValues().forEach(v => { ctx[String(v[0])] = String(v[1]); });
  let chiffres: Chiffre[] = [];
  try { chiffres = JSON.parse(ctx['chiffres'] || '[]') as Chiffre[]; } catch (e) { chiffres = []; }
  // Dans la version diffusée, on garde les alertes utiles aux lecteurs (sources non collectées).
  const alertes = (ctx['alertes'] || '').split(' || ').filter(a => a.indexOf('La source ') === 0);
  const gardees = lignes.filter(l => l.garder === 'Oui');
  const sansResume = gardees.filter(l => !l.resume);
  // Un élément gardé sans résumé n'est pas publié.
  const publiees = lignes.map(l => (l.garder === 'Oui' && !l.resume) ? { garder: 'Non', categorie: l.categorie, titre: l.titre, resume: l.resume, aVerifier: l.aVerifier, controles: l.controles,
    statut: l.statut, liens: l.liens, question: l.question, extrait: l.extrait, id: l.id, resumeIA: l.resumeIA } : l);
  sansResume.forEach(l => alertes.push('« ' + l.titre + ' » a été gardé sans résumé et n’est pas publié.'));
  const nbModifies = gardees.filter(l => l.resume && l.resume !== l.resumeIA).length;
  const html = rendre(publiees, chiffres, alertes, '', ctx['fenetreDebut'] || new Date().toISOString(), ctx['fenetreFin'] || new Date().toISOString(), String(validePar).trim());
  const suivi = workbook.getTable('Suivi');
  if (suivi && suivi.getRowCount() > 0) {
    const v = suivi.getRangeBetweenHeaderAndTotal().getValues();
    for (let i = v.length - 1; i >= 0; i--) if (v[i].every(c => String(c).trim() === '')) suivi.deleteRowsAt(i, 1);
  }
  if (suivi) suivi.addRows(-1, [[new Date().toISOString().slice(0, 10), (ctx['fenetreDebut'] || '').slice(0, 10), String(validePar).trim(),
    lignes.length, gardees.length - sansResume.length, nbModifies, Math.round(Number(dureeRelectureMin) || 0), String(commentaire || '')]]);
  return { html: html, nbGardes: gardees.length - sansResume.length, nbModifies: nbModifies, nbSansResume: sansResume.length };
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
