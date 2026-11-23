// Office Script 1, « 1_Preparer ».
// Appelé par Power Automate (Excel Online Business, action « Exécuter un script »).
// Étape déterministe, sans IA. Fenêtre de dates, mots exclus, doublons, chiffres divergents, pré-score.
// Entrée  articlesJson, tableau JSON construit dans le flux à partir des flux RSS (un tableau par source accepté).
// Sortie  un texte JSON transmis tel quel à l'étape IA et au script 2.

interface Article { titre: string; url: string; extrait: string; date: string; source: string; fiabilite: string; }
interface Reprise { source: string; fiabilite: string; url: string; titre: string; extrait: string; }
interface Candidat {
  id: string; titre: string; url: string; source: string; fiabilite: string; date: string; extrait: string;
  reprises: Reprise[]; divergence: string[]; scorePre: number;
}
interface Stats { collectes: number; horsFenetre: number; exclus: number; doublonsFusionnes: number; sujets: number; envoyesIA: number; }
interface ElementIA { id: string; titre: string; extrait: string; source: string; fiabilite: string; date: string; variantes: { source: string; titre: string; extrait: string }[]; }
interface Preparation { candidats: Candidat[]; stats: Stats; fenetreDebut: string; fenetreFin: string; elementsIA: string; }
interface Parametres { fenetreJours: number; maxArticlesIA: number; motsCles: string[]; motsExclus: string[]; concurrents: string[]; }
interface ArticleTravail extends Article { canon: string; tok: Set<string>; }

function main(workbook: ExcelScript.Workbook, articlesJson: string, dateReference: string): string {
  const p = lireParametres(workbook);
  const ref = dateReference ? new Date(dateReference) : new Date();
  const minDate = new Date(ref.getTime() - p.fenetreJours * 86400000);
  // Le flux ajoute un tableau d'articles par source. On aplatit d'un niveau.
  const recu = JSON.parse(articlesJson || '[]') as (Article | Article[])[];
  const brut: Article[] = [];
  recu.forEach(x => { if (Array.isArray(x)) x.forEach(a => brut.push(a)); else brut.push(x); });
  const stats: Stats = { collectes: brut.length, horsFenetre: 0, exclus: 0, doublonsFusionnes: 0, sujets: 0, envoyesIA: 0 };

  let arts: Article[] = brut.filter(a => a && a.titre && a.url).map(a => ({
    titre: texte(a.titre), url: String(a.url), extrait: texte(a.extrait).slice(0, 600),
    date: a.date ? String(a.date) : '', source: String(a.source || ''), fiabilite: ['A', 'B', 'C'].indexOf(String(a.fiabilite)) >= 0 ? String(a.fiabilite) : 'C'
  }));
  arts = arts.filter(a => {
    if (a.date && new Date(a.date).getTime() < minDate.getTime()) { stats.horsFenetre++; return false; }
    return true;
  });
  const exclus = p.motsExclus.map(norm).filter(w => w.length > 0);
  arts = arts.filter(a => {
    const t = ' ' + norm(a.titre + ' ' + a.extrait) + ' ';
    if (exclus.some(w => t.indexOf(' ' + w + ' ') >= 0)) { stats.exclus++; return false; }
    return true;
  });
  const rang = (f: string): number => f === 'A' ? 0 : f === 'B' ? 1 : 2;
  arts.sort((x, y) => (rang(x.fiabilite) - rang(y.fiabilite)) || (temps(x.date) - temps(y.date)));

  const groupes: ArticleTravail[][] = [];
  for (const a of arts) {
    const t: ArticleTravail = { titre: a.titre, url: a.url, extrait: a.extrait, date: a.date, source: a.source, fiabilite: a.fiabilite, canon: canon(a.url), tok: jetons(a.titre) };
    const g = groupes.find(gr => gr.some(m => m.canon === t.canon || jaccard(m.tok, t.tok) >= 0.55));
    if (g) { g.push(t); stats.doublonsFusionnes++; } else { groupes.push([t]); }
  }
  stats.sujets = groupes.length;

  const kw = p.motsCles.map(norm).filter(w => w.length > 0);
  const conc = p.concurrents.map(norm).filter(w => w.length > 0);
  const candidats: Candidat[] = groupes.map(g => {
    const prem = g[0];
    const vus = new Set<string>([prem.canon]);
    const reprises: Reprise[] = [];
    for (const m of g.slice(1)) {
      if (vus.has(m.canon)) continue;
      vus.add(m.canon);
      reprises.push({ source: m.source, fiabilite: m.fiabilite, url: m.url, titre: m.titre, extrait: m.extrait });
    }
    const divergence = chiffresDivergents([prem.titre + ' ' + prem.extrait].concat(reprises.map(r => r.titre + ' ' + r.extrait)));
    const txt = ' ' + norm(prem.titre + ' ' + prem.extrait) + ' ';
    const nbKw = Math.min(3, kw.filter(w => txt.indexOf(w) >= 0).length);
    const age = prem.date ? (ref.getTime() - temps(prem.date)) / 86400000 : 7;
    const poids = prem.fiabilite === 'A' ? 3 : prem.fiabilite === 'B' ? 2 : 1;
    const score = poids + (age <= 2 ? 2 : age <= 5 ? 1 : 0) + nbKw * 0.5 + (conc.some(w => txt.indexOf(w) >= 0) ? 1 : 0)
      + Math.min(2, reprises.length * 0.5) + (divergence.length ? 1 : 0);
    return { id: empreinte(prem.canon), titre: prem.titre, url: prem.url, source: prem.source, fiabilite: prem.fiabilite, date: prem.date,
      extrait: prem.extrait, reprises: reprises, divergence: divergence, scorePre: Math.round(score * 10) / 10 };
  }).sort((a, b) => b.scorePre - a.scorePre).slice(0, p.maxArticlesIA);
  stats.envoyesIA = candidats.length;

  // Ce que l'IA verra. Ni adresse web, ni donnée interne.
  const elements: ElementIA[] = candidats.map(c => ({ id: c.id, titre: c.titre, extrait: c.extrait, source: c.source, fiabilite: c.fiabilite, date: c.date,
    variantes: c.reprises.map(r => ({ source: r.source, titre: r.titre, extrait: r.extrait })) }));
  const sortie: Preparation = { candidats: candidats, stats: stats, fenetreDebut: minDate.toISOString(), fenetreFin: new Date(ref.getTime() - 86400000).toISOString(), elementsIA: JSON.stringify({ elements: elements }) };
  return JSON.stringify(sortie);
}

function lireParametres(workbook: ExcelScript.Workbook): Parametres {
  const p: Parametres = { fenetreJours: 7, maxArticlesIA: 30, motsCles: [], motsExclus: [], concurrents: [] };
  const t = workbook.getTable('Parametres');
  if (!t) return p;
  const lignes = t.getRangeBetweenHeaderAndTotal().getValues();
  lignes.forEach(l => {
    const cle = String(l[0]).trim(); const val = String(l[1]).trim();
    const liste = val.split(';').map(x => x.trim()).filter(x => x.length > 0);
    if (cle === 'fenetreJours') p.fenetreJours = Number(val) || 7;
    if (cle === 'maxArticlesIA') p.maxArticlesIA = Number(val) || 30;
    if (cle === 'motsCles') p.motsCles = liste;
    if (cle === 'motsExclus') p.motsExclus = liste;
    if (cle === 'concurrents') p.concurrents = liste;
  });
  return p;
}

function texte(s: string): string {
  // Les résumés RSS contiennent souvent du HTML.
  return String(s || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&rsquo;/g, '’').replace(/\s+/g, ' ').trim();
}
function norm(s: string): string {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
}
function canon(u: string): string {
  const s = String(u || '').trim().replace(/#.*$/, '');
  const m = s.match(/^[a-z]+:\/\/([^\/?]+)([^?]*)\??(.*)$/i);
  if (!m) return s.toLowerCase();
  const q = m[3].split('&').filter(x => x.length > 0 && !/^(utm_[a-z]+|fbclid|gclid|xtor)=/i.test(x)).join('&');
  return m[1].toLowerCase().replace(/^www\./, '') + m[2].replace(/\/+$/, '') + (q ? '?' + q : '');
}
function jetons(s: string): Set<string> {
  const vides = ['les', 'des', 'une', 'aux', 'pour', 'sur', 'dans', 'par', 'avec', 'ces', 'son', 'ses', 'qui', 'que', 'est', 'sont', 'des', 'plus', 'selon', 'cette', 'leur', 'leurs', 'chez', 'entre', 'vers', 'sans', 'sous', 'tout', 'tous', 'nous', 'vous', 'elle', 'ils', 'elles'];
  return new Set<string>(norm(s).split(' ').filter(w => w.length >= 3 && vides.indexOf(w) < 0));
}
function jaccard(a: Set<string>, b: Set<string>): number {
  let i = 0; a.forEach(x => { if (b.has(x)) i++; });
  return i / ((a.size + b.size - i) || 1);
}
function empreinte(s: string): string {
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0;
  return ('00000000' + h.toString(16)).slice(-8);
}
function temps(d: string): number { const t = new Date(d).getTime(); return isNaN(t) ? 0 : t; }
function chiffresDivergents(textes: string[]): string[] {
  // Compare les chiffres suivis d'une même unité d'un texte à l'autre.
  const parUnite: { [u: string]: string[][] } = {};
  textes.forEach(t => {
    const vu: { [u: string]: string[] } = {};
    const re = /(\d+(?:[.,]\d+)?)\s?(millions?|%|points?|milliers?)/gi;
    let x = re.exec(t);
    while (x) {
      const u = x[2].toLowerCase().replace(/s$/, '');
      const v = x[1].replace(',', '.');
      if (!vu[u]) vu[u] = [];
      if (vu[u].indexOf(v) < 0) vu[u].push(v);
      x = re.exec(t);
    }
    Object.keys(vu).forEach(u => { if (!parUnite[u]) parUnite[u] = []; parUnite[u].push(vu[u].sort()); });
  });
  const res: string[] = [];
  Object.keys(parUnite).forEach(u => {
    const ens = parUnite[u];
    if (ens.length >= 2 && ens.some(e => e.join('|') !== ens[0].join('|'))) {
      const toutes: string[] = [];
      ens.forEach(e => e.forEach(v => { if (toutes.indexOf(v) < 0) toutes.push(v); }));
      res.push(toutes.map(v => v.replace('.', ',')).join(' / ') + ' ' + u + (u === 'million' ? 's' : ''));
    }
  });
  return res;
}
