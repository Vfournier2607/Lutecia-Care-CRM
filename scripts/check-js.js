// Vérifie la syntaxe JS de tous les scripts inline des pages HTML + fichiers js/.
// Utilisé en CI avant le déploiement Azure : une page cassée ne part pas en prod.
const fs = require('fs');
let fail = false;
for (const f of ['index.html', 'fiche.html']) {
  const c = fs.readFileSync(f, 'utf8');
  const blocks = c.match(/<script(?![^>]*src)[^>]*>([\s\S]*?)<\/script>/g) || [];
  blocks.forEach((s, i) => {
    try { new Function(s.replace(/<\/?script[^>]*>/g, '')); }
    catch (e) { console.error(`${f} script#${i}: ${e.message}`); fail = true; }
  });
}
for (const f of ['js/config.js', 'js/auth.js', 'js/graph.js']) {
  try { new Function(fs.readFileSync(f, 'utf8')); }
  catch (e) { console.error(`${f}: ${e.message}`); fail = true; }
}
if (fail) process.exit(1);
console.log('Syntaxe JS OK');
