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

const reportHtml = fs.readFileSync(path.join(root, 'report_template.html'), 'utf8');
if (!/id="cloudSyncProgress"[\s\S]*cloud-sync-spinner/.test(reportHtml)) {
  throw new Error('report_template.html: cloud save progress indicator missing');
}
if (!/setCloudSaveBusy\(true,[\s\S]*apiPost\('\/api\/manage\/cloudSettings'/.test(reportHtml)) {
  throw new Error('report_template.html: cloud save must enter busy state before request');
}
