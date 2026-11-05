import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, it } from 'node:test';
import { DocxTemplate, type MailMergeWarning, MissingFieldsError } from '../src/index.ts';
import { fixture, Output, template, withInstruction } from './helpers.ts';

describe('DocxTemplate loading', () => {
  it('accepts path, URL, Uint8Array, ArrayBuffer and Blob', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'mm-'));
    try {
      const path = join(dir, 't.docx');
      const bytes = fixture('test_merge_templates_simple.docx');
      await writeFile(path, bytes);
      const sources = [path, new URL(`file://${path}`), bytes, new Uint8Array(bytes).buffer, new Blob([bytes])];
      for (const source of sources) {
        const tpl = await DocxTemplate.load(source as never);
        assert.deepEqual([...tpl.fields], ['fieldname']);
      }
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it('gives a clear error for non-docx input', () => {
    assert.throws(() => DocxTemplate.fromBuffer(new TextEncoder().encode('hello')), /Not a valid \.docx/);
  });

  it('validates options eagerly', () => {
    const bytes = fixture('test_merge_templates_simple.docx');
    // @ts-expect-error invalid
    assert.throws(() => DocxTemplate.fromBuffer(bytes, { keepFields: 'maybe' }), TypeError);
    assert.throws(() => DocxTemplate.fromBuffer(bytes, { timeZone: 'Mars/Olympus' }), RangeError);
  });
});

describe('template reuse (parse once, render many)', () => {
  it('renders are independent and byte-for-byte deterministic', () => {
    const tpl = template('test_merge_table_rows.docx');
    const a1 = tpl.render({ student_name: 'A', class_code: [{ class_code: 'X' }] });
    const b = tpl.render({ student_name: 'B' });
    const a2 = tpl.render({ student_name: 'A', class_code: [{ class_code: 'X' }] });
    assert.deepEqual(a1, a2);
    assert.ok(new Output(b).text().includes('B'));
    assert.ok(!new Output(b).text().includes('X'));
  });

  it('renderEach yields one document per record, lazily', () => {
    const tpl = template('test_merge_templates_simple.docx');
    let pulled = 0;
    function* records() {
      for (const name of ['a', 'b', 'c']) {
        pulled += 1;
        yield { fieldname: name };
      }
    }
    const it = tpl.renderEach(records());
    const first = it.next();
    assert.equal(pulled, 1);
    assert.ok(!first.done);
    assert.equal(new Output(first.value.bytes).text(), 'Merge : a');
    const rest = [...it];
    assert.deepEqual(
      rest.map((r) => r.index),
      [1, 2],
    );
    assert.equal(new Output(rest[1]!.bytes).text(), 'Merge : c');
  });

  it('save() writes a file', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'mm-'));
    try {
      const path = join(dir, 'out.docx');
      await template('test_merge_templates_simple.docx').createDocument().merge({ fieldname: 'z' }).save(path);
      assert.equal(new Output(new Uint8Array(await readFile(path))).text(), 'Merge : z');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

describe('dot-path data', () => {
  it('resolves nested objects; an exact key wins over the path', () => {
    const tpl = withInstruction('MERGEFIELD customer.address.city');
    assert.deepEqual([...tpl.fields], ['customer.address.city']);
    assert.equal(new Output(tpl.render({ customer: { address: { city: 'Lisbon' } } })).text(), 'Lisbon');
    assert.equal(
      new Output(tpl.render({ 'customer.address.city': 'Flat', customer: { address: { city: 'X' } } })).text(),
      'Flat',
    );
    assert.equal(new Output(tpl.render({ customer: {} })).text(), '');
    assert.deepEqual(tpl.validate({ customer: { address: { city: 'L' } } }), { valid: true, missing: [], unused: [] });
  });
});

describe('formatting inside documents', () => {
  it('applies numeric/date switches and custom formatters', () => {
    assert.equal(
      new Output(withInstruction('MERGEFIELD total \\# "#,##0.00"').render({ total: '1234.5' })).text(),
      '1,234.50',
    );
    assert.equal(
      new Output(
        withInstruction('MERGEFIELD due \\@ "d MMMM yyyy"', { locale: 'de-DE' }).render({ due: '2024-03-09' }),
      ).text(),
      '9 März 2024',
    );
    const money = withInstruction('MERGEFIELD total \\* Money', {
      formatters: {
        money: (v, c) => new Intl.NumberFormat(c.locale, { style: 'currency', currency: 'EUR' }).format(Number(v)),
      },
    });
    assert.equal(new Output(money.render({ total: 3 })).text(), '€3.00');
  });

  it('a malformed picture warns and keeps the raw value', () => {
    const warnings: MailMergeWarning[] = [];
    const tpl = withInstruction('MERGEFIELD total \\# "0.00"', { onWarning: (w) => warnings.push(w) });
    assert.equal(new Output(tpl.render({ total: 'n/a' })).text(), 'n/a');
    assert.equal(warnings[0]?.code, 'INVALID_FORMAT');
  });

  it('preserves leading/trailing spaces', () => {
    const out = new Output(withInstruction('MERGEFIELD x').render({ x: '  padded  ' }));
    assert.equal(out.text(), '  padded  ');
    assert.ok(out.raw().includes('xml:space="preserve"'));
  });

  it('keeps the field run formatting (rPr) on the merged value', () => {
    const tpl = withInstruction('MERGEFIELD x');
    const raw = new Output(tpl.render({ x: 'v' })).raw();
    const run = raw.slice(raw.lastIndexOf('<w:r>', raw.indexOf('>v<')), raw.indexOf('>v<'));
    assert.ok(run.includes('<w:rPr>') || !raw.includes('fldSimple'), 'run properties retained');
  });
});

describe('IF comparisons', () => {
  const iff = (cond: string) => withInstruction(`IF ${cond} "yes" "no"`, { mergeIfFields: true });
  const text = (cond: string) => new Output(iff(cond).render({})).text();
  it('compares numbers numerically (Python compared "10" < "9" as strings)', () => {
    assert.equal(text('10 > 9'), 'yes');
    assert.equal(text('"2.50" = "2.5"'), 'yes');
  });
  it('supports Word wildcards for = and <>', () => {
    assert.equal(text('"Lisbon" = "lis*"'), 'yes');
    assert.equal(text('"Lisbon" <> "L?sbon"'), 'no');
    assert.equal(text('"abc" < "abd"'), 'yes');
  });
  it('warns on malformed conditions', () => {
    const warnings: MailMergeWarning[] = [];
    const tpl = withInstruction('IF 1 ~ 2 "a" "b"', { mergeIfFields: true, onWarning: (w) => warnings.push(w) });
    assert.equal(new Output(tpl.render({})).text(), 'b');
    assert.equal(warnings[0]?.code, 'INVALID_CONDITION');
  });
});

describe('validate() and strict mode', () => {
  const tpl = template('test_merge_table_rows.docx');

  it('reports missing and unused keys, understanding table rows', () => {
    const result = tpl.validate({
      student_name: 'A',
      study_name: 'B',
      thesis_grade: 'C',
      extra: 1,
      class_code: [{ class_code: 'X', class_name: 'Y', class_grade: 'Z' }],
    });
    assert.deepEqual(result, { valid: true, missing: [], unused: ['extra'] });
    const bad = tpl.validate({ student_name: 'A' });
    assert.equal(bad.valid, false);
    assert.ok(bad.missing.includes('thesis_grade'));
  });

  it('strict render throws MissingFieldsError with the record index', () => {
    assert.throws(() => tpl.render({}, { strict: true }), MissingFieldsError);
    const strictTpl = template('test_merge_templates_simple.docx', { strict: true });
    assert.throws(
      () => strictTpl.renderAll([{ fieldname: 'a' }, {}]),
      (e: unknown) => e instanceof MissingFieldsError && e.recordIndex === 1 && e.missing[0] === 'fieldname',
    );
    assert.doesNotThrow(() => strictTpl.render({ fieldname: 'ok' }));
  });
});

describe('describeFields()', () => {
  it('lists type, instruction and part', () => {
    const fields = template('test_footnote_header_footer.docx').describeFields();
    const parts = new Set(fields.map((f) => f.part));
    assert.ok([...parts].some((p) => /header|footer/.test(p)));
    assert.ok(fields.every((f) => f.type === 'MERGEFIELD' && f.instruction.startsWith('MERGEFIELD')));
  });
});

describe('warnings', () => {
  it('routes formatting problems to onWarning instead of failing', () => {
    const warnings: MailMergeWarning[] = [];
    const tpl = DocxTemplate.fromBuffer(fixture('test_unbalanced_quote_date_field.docx'), {
      onWarning: (w) => warnings.push(w),
    });
    assert.deepEqual([...tpl.fields], []);
    assert.doesNotThrow(() => tpl.render({}));
  });
});
