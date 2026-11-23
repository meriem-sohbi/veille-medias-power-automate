// Compile les trois Office Scripts en TypeScript strict, puis lance les cinq scénarios.
// Usage, à la racine du dépôt :  npm test   (ou  node tests/lancer_les_tests.js)
// Prérequis : Node.js 18 ou plus récent. TypeScript est téléchargé automatiquement par npx si besoin.
const { execSync } = require('child_process');
const path = require('path');
const ici = __dirname;
const scripts = ['1_Preparer', '2_Controler', '3_Publier'];
for (const s of scripts) {
  const cmd = `npx -y -p typescript@5 tsc --strict --noImplicitAny --noUnusedLocals --target ES2019 --lib ES2019 ` +
    `--outDir "${path.join(ici, 'out', s)}" "${path.join(ici, 'excelscript.d.ts')}" "${path.join(ici, '..', 'office-scripts', s + '.ts')}"`;
  process.stdout.write(`Compilation de ${s} en TypeScript strict... `);
  execSync(cmd, { stdio: ['ignore', 'inherit', 'inherit'] });
  console.log('OK');
}
console.log('\nExécution des cinq scénarios :');
const bilan = JSON.parse(execSync(`node "${path.join(ici, 'test.js')}"`, { encoding: 'utf8', maxBuffer: 1 << 24 }));
for (const b of bilan) {
  console.log(`\n- ${b.scenario}`);
  console.log(`  ${b.stats.collectes} articles collectés, ${b.stats.horsFenetre} hors période, ${b.stats.exclus} exclu, ${b.stats.doublonsFusionnes} doublons regroupés, ${b.stats.sujets} sujets envoyés à l'IA`);
  console.log(`  Adresse web ou donnée interne envoyée à l'IA : ${b.fuiteVersIA ? 'OUI (anomalie)' : 'non'}`);
  console.log(`  Réponse IA exploitable : ${b.iaOk ? 'oui' : 'non, rédaction manuelle'}. ${b.proposes} éléments proposés, ${b.aVerifier} à vérifier.`);
  b.alertes.forEach(a => console.log('  Signalé : ' + a.slice(0, 150)));
  if (b.publication) {
    console.log(`  Publication sans nom : ${b.sansValideur}`);
    console.log(`  Publication avec nom : ${b.publication.gardes} éléments publiés, ${b.publication.modifies} résumé corrigé, ligne de suivi ajoutée.`);
  }
}
console.log(`\nRésultats détaillés : ${path.join(ici, 'resultats')} (brouillons et version validée au format HTML).`);

// Vérifications : la commande échoue (et le badge GitHub passe au rouge) si un résultat attendu manque.
const fs = require('fs');
const lire = f => fs.readFileSync(path.join(ici, 'resultats', f), 'utf8');
const sc = nom => bilan.find(b => b.scenario === nom);
const verifs = [
  ['Préparation : 15 collectés, 1 hors période, 1 exclu, 3 doublons, 10 sujets', bilan.every(b => b.stats.collectes === 15 && b.stats.horsFenetre === 1 && b.stats.exclus === 1 && b.stats.doublonsFusionnes === 3 && b.stats.sujets === 10)],
  ['Aucune adresse web ni donnée interne envoyée à l’IA', bilan.every(b => !b.fuiteVersIA)],
  ['Chiffres divergents marqués « à arbitrer »', sc('1_nominal').alertes.some(a => a.includes('à arbitrer'))],
  ['Source unique peu fiable marquée « à confirmer »', sc('1_nominal').alertes.some(a => a.includes('à confirmer'))],
  ['Consigne cachée écartée', sc('1_nominal').alertes.some(a => a.startsWith('Non | écarté'))],
  ['Chiffre et nom inventés par l’IA signalés', sc('3_ia_invente').alertes.some(a => a.includes('14') && a.includes('Jean'))],
  ['Réponse IA illisible : bascule en rédaction manuelle', sc('2_ia_illisible').iaOk === false && lire('2_ia_illisible_brouillon.html').includes('résumé à rédiger à la main')],
  ['Audiences absentes : aucun ancien chiffre affiché', lire('4_audiences_absentes_brouillon.html').includes('Indisponibles cette semaine')],
  ['Réponse IA entourée de balises de code : lue correctement', sc('5_ia_bloc_json').iaOk === true],
  ['Publication sans nom refusée', String(sc('1_nominal').sansValideur).startsWith('refusé')],
  ['Publication avec nom : version validée et ligne de suivi', sc('1_nominal').publication.gardes > 0 && lire('1_nominal_valide.html').includes('Validée par Camille Martin')],
  ['Titre conforme à l’exemple du rapport', lire('1_nominal_valide.html').includes('Veille médias, semaine du 16 au 22 novembre 2026')],
];
console.log('\nVérifications :');
verifs.forEach(([nom, ok]) => console.log(`  ${ok ? 'OK    ' : 'ÉCHEC '} ${nom}`));
const echecs = verifs.filter(v => !v[1]).length;
console.log(echecs ? `\n${echecs} vérification(s) en échec.` : `\nToutes les vérifications passent (${verifs.length}/${verifs.length}).`);
process.exit(echecs ? 1 : 0);
