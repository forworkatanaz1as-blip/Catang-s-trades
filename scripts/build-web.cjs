const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const output = path.join(root, 'www');
fs.mkdirSync(output, { recursive: true });
fs.copyFileSync(path.join(root, 'MAGAZIN_mobile.html'), path.join(output, 'index.html'));
