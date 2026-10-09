const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const output = path.join(root, 'dist');
// Only public runtime files belong in the deployment artifact.
const assets = ['index.html', 'app.js', 'appearance.js', 'styles.css', 'manifest.json', 'sw.js'];
for (const asset of assets) fs.accessSync(path.join(root, asset), fs.constants.R_OK);
fs.rmSync(output, { recursive: true, force: true });
fs.mkdirSync(output, { recursive: true });
for (const asset of assets) fs.copyFileSync(path.join(root, asset), path.join(output, asset));
console.log(`Prepared ${assets.length} public files in dist/`);
