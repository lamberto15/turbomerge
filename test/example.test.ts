import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { buildInvoiceTemplate, money, sampleInvoices } from '../examples/invoice-template.ts';
import { DocxTemplate } from '../src/index.ts';
import { Output } from './helpers.ts';

describe('template generated from scratch (examples/invoice-template.ts)', () => {
  const template = DocxTemplate.fromBuffer(buildInvoiceTemplate(), {
    mergeIfFields: true,
    strict: true,
    formatters: { money },
    onWarning: (w) => assert.fail(w.message),
  });

  it('discovers every field, including nested and dot-path ones', () => {
    assert.deepEqual([...template.fields].sort(), [
      'customer.city',
      'customer.name',
      'date',
      'description',
      'number',
      'paid',
      'price',
      'qty',
      'sku',
      'total',
    ]);
    assert.ok(template.describeFields().some((f) => f.type === 'IF'));
    assert.ok(template.describeFields().some((f) => f.name === 'paid' && f.nested));
  });

  it('renders a full invoice', () => {
    const text = new Output(template.render(sampleInvoices[0]!)).text();
    for (const expected of [
      'INVOICE INV-1001',
      'Date: 5 November 2026',
      'Bill to: Ada Lovelace, London',
      'A-1Analytical engine tuning1€1,200.00',
      'B-7Punch cards (box)3€15.50',
      'Total: 1,246.50 EUR',
      'Thank you for your payment.',
    ])
      assert.ok(text.includes(expected), `${expected}\n  in: ${text}`);
  });

  it('renders all invoices into one document and one per record', () => {
    const all = new Output(template.renderAll(sampleInvoices));
    assert.equal(all.pageBreaks(), 1);
    assert.ok(all.text().includes('Payment due within 30 days.'));
    assert.equal(all.liveFields(), 0);
    assert.equal([...template.renderEach(sampleInvoices)].length, 2);
  });
});
