const fs = require('fs');

const content = [
  'BT',
  '/F1 11 Tf',
  '50 760 Td (Daily income / expense) Tj',
  '0 -22 Td (Date Time Type Payment Description Amount) Tj',
  '0 -22 Td (21.10.2025 14:32:01 Income Cash Playstation 1h 8,000 TND) Tj',
  '0 -22 Td (21.10.2025 15:10:22 Expense Cash Coca Cola 3,500 TND) Tj',
  '0 -22 Td (21.10.2025 16:03:11 Income Credit Card Member topup 25,000 TND (evening)) Tj',
  'ET',
].join('\n');

const stream = `<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream`;
const objs = [
  '1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj',
  '2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj',
  '3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >> endobj',
  `4 0 obj ${stream} endobj`,
  '5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj',
];

let body = '%PDF-1.4\n';
const offsets = [0];
for (const o of objs) {
  offsets.push(Buffer.byteLength(body));
  body += o + '\n';
}
const xrefStart = Buffer.byteLength(body);
let xref = `xref\n0 ${objs.length + 1}\n`;
xref += '0000000000 65535 f \n';
for (let i = 1; i < offsets.length; i++) {
  xref += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`;
}
body += xref;
body += `trailer << /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`;
fs.writeFileSync('21-10.pdf', body);
console.log('wrote 21-10.pdf', Buffer.byteLength(body));
