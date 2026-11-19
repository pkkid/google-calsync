const fs = require('node:fs');
const vm = require('node:vm');
for (const file of fs.readdirSync('src').filter((file) => file.endsWith('.js'))) {
  new vm.Script(fs.readFileSync('src/' + file, 'utf8'), { filename: file });
}
const manifest = JSON.parse(fs.readFileSync('src/appsscript.json', 'utf8'));
if (manifest.runtimeVersion !== 'V8') throw new Error('V8 runtime is required.');
console.log('Apps Script source syntax and manifest checked.');
