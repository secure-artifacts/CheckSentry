const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
for (const name of ['report_template.html', 'management_template.html']) {
  const html = fs.readFileSync(path.join(root, name), 'utf8');
  const scripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)].map(match => match[1]);
  if (!scripts.length) throw new Error(`${name}: no inline script found`);
  for (const script of scripts) new Function(script);
  console.log(`${name}: JavaScript syntax OK`);
}
