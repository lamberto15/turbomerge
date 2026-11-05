/**
 * Builds a mail-merge template .docx from scratch - no Word needed.
 *
 *   node examples/invoice-template.ts            -> examples/invoice-template.docx + invoice-output.docx
 *
 * Shows: complex and simple MERGEFIELDs, dot paths, \@ date and \# number pictures,
 * a repeating table row, an IF field and a custom \* formatter.
 */
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strToU8, zipSync } from 'fflate';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

const text = (t: string, bold = false) =>
  `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;

/** What Word writes for Ctrl+F9 fields: begin / instrText / separate / result / end. */
const field = (instr: string, bold = false) => {
  const rPr = bold ? '<w:rPr><w:b/></w:rPr>' : '';
  const name = /MERGEFIELD\s+(\S+)/.exec(instr)?.[1] ?? 'field';
  return (
    `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r>${rPr}<w:instrText xml:space="preserve"> ${esc(instr)} </w:instrText></w:r>` +
    `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r>${rPr}<w:t>«${esc(name)}»</w:t></w:r>` +
    `<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`
  );
};

/** The compact form Word also uses. */
const simpleField = (instr: string) =>
  `<w:fldSimple w:instr=" ${esc(instr)} "><w:r><w:t>«field»</w:t></w:r></w:fldSimple>`;

const p = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`;
const cell = (content: string, width: number) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${p(content)}</w:tc>`;
const row = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`;

export function buildInvoiceTemplate(): Uint8Array {
  const body = [
    p(text('INVOICE ', true), field('MERGEFIELD number', true)),
    p(text('Date: '), field('MERGEFIELD date \\@ "d MMMM yyyy"')),
    p(
      text('Bill to: '),
      field('MERGEFIELD customer.name \\* Caps'),
      text(', '),
      simpleField('MERGEFIELD customer.city'),
    ),
    `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9000" w:type="dxa"/>` +
      `<w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/>` +
      `<w:insideH w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="4500"/><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
      row(
        cell(text('SKU', true), 1500),
        cell(text('Description', true), 4500),
        cell(text('Qty', true), 1000),
        cell(text('Price', true), 2000),
      ) +
      row(
        cell(field('MERGEFIELD sku'), 1500),
        cell(field('MERGEFIELD description'), 4500),
        cell(field('MERGEFIELD qty'), 1000),
        cell(field('MERGEFIELD price \\* Money'), 2000),
      ) +
      `</w:tbl>`,
    p(text('Total: ', true), field('MERGEFIELD total \\# "#,##0.00"', true), text(' EUR', true)),
    p(field('IF "{ MERGEFIELD paid }" = "yes" "Thank you for your payment." "Payment due within 30 days."')),
  ].join('');

  // The IF above contains a nested MERGEFIELD: expand the "{ … }" into a real nested field.
  const withNested = body.replace(
    /\{ MERGEFIELD paid \}/,
    `</w:instrText></w:r>${field('MERGEFIELD paid')}<w:r><w:instrText xml:space="preserve">`,
  );

  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${withNested}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;

  return zipSync({
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    'word/_rels/document.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`),
    'word/settings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="708"/><w:compat/></w:settings>`),
    'word/document.xml': strToU8(document),
  });
}

export const sampleInvoices = [
  {
    number: 'INV-1001',
    date: '2026-11-05',
    customer: { name: 'ada lovelace', city: 'London' },
    sku: [
      { sku: 'A-1', description: 'Analytical engine tuning', qty: 1, price: 1200 },
      { sku: 'B-7', description: 'Punch cards (box)', qty: 3, price: 15.5 },
    ],
    total: 1246.5,
    paid: 'yes',
  },
  {
    number: 'INV-1002',
    date: '2026-11-06',
    customer: { name: 'grace hopper', city: 'Arlington' },
    sku: [{ sku: 'C-3', description: 'Compiler consulting', qty: 10, price: 250 }],
    total: 2500,
    paid: 'no',
  },
];

export const money = (v: unknown) =>
  new Intl.NumberFormat('en-IE', { style: 'currency', currency: 'EUR' }).format(Number(v));

if (import.meta.main) {
  const { DocxTemplate } = await import('../src/index.ts');
  const bytes = buildInvoiceTemplate();
  writeFileSync(join(import.meta.dirname, 'invoice-template.docx'), bytes);

  const template = DocxTemplate.fromBuffer(bytes, {
    mergeIfFields: true,
    strict: true,
    formatters: { money },
  });
  console.log('fields:', [...template.fields].join(', '));
  writeFileSync(
    join(import.meta.dirname, 'invoice-output.docx'),
    template.renderAll(sampleInvoices, { separator: 'page_break' }),
  );
  for (const { index, bytes: doc } of template.renderEach(sampleInvoices)) {
    writeFileSync(join(import.meta.dirname, `invoice-${index + 1}.docx`), doc);
  }
  console.log('wrote invoice-template.docx, invoice-output.docx, invoice-1.docx, invoice-2.docx');
}
