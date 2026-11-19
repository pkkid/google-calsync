// Syntax-check every Apps Script source in the given folder(s) and verify
// the manifest. Also fails if a private Outlook calendar URL slips into
// source, since this repo is public. Usage: node scripts/check.js html ical
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const dirs = process.argv.slice(2);
if (!dirs.length) throw new Error('Usage: node scripts/check.js <dir> [dir...]');
for (const dir of dirs) {
  for (const file of fs.readdirSync(dir).filter((f) => f.endsWith('.js'))) {
    const code = fs.readFileSync(path.join(dir, file), 'utf8');
    new vm.Script(code, { filename: path.join(dir, file) });
    if (/outlook\.office365\.com\/owa\/calendar\/[0-9a-f]{16,}/i.test(code)) {
      throw new Error(`${dir}/${file} contains a private Outlook calendar URL; use Script Properties.`);
    }
  }
  const manifest = JSON.parse(fs.readFileSync(path.join(dir, 'appsscript.json'), 'utf8'));
  if (manifest.runtimeVersion !== 'V8') throw new Error(`${dir}: V8 runtime is required.`);
  console.log(`${dir}: syntax and manifest OK`);
}
