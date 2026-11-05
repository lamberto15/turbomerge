import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { NS } from '../src/constants.ts';
import { descendants } from '../src/xml.ts';
import { Output, template } from './helpers.ts';

const values = ['one', 'two', 'three', 'four', 'five'];

describe('merge()', () => {
  it('replaces simple and complex fields and removes the mail-merge data source', () => {
    const tpl = template('test_merge_templates_simple.docx');
    assert.deepEqual([...tpl.fields], ['fieldname']);
    const out = new Output(tpl.render({ fieldname: 'hello' }));
    assert.equal(out.text(), 'Merge : hello');
    assert.equal(out.liveFields(), 0);
    assert.ok(!out.raw('word/settings.xml').includes('w:mailMerge'));
    assert.ok(!out.raw().includes('MergeField'));
  });

  it('turns newlines into line breaks', () => {
    const out = new Output(template('test_merge_templates_simple.docx').render({ fieldname: 'a\nb' }));
    assert.deepEqual(out.texts().slice(-2), ['a', 'b']);
    assert.equal(descendants(out.xml(), NS.w, 'cr').length, 1);
  });

  it('merges headers, footers and footnotes', () => {
    const tpl = template('test_footnote_header_footer.docx');
    const record = Object.fromEntries([...tpl.fields].map((f) => [f, `<${f}>`]));
    const out = new Output(tpl.render(record));
    const all = Object.keys(out.files).filter((f) => /word\/(header|footer|footnotes|endnotes)\d*\.xml$/.test(f));
    for (const part of all) assert.equal(out.liveFields(part), 0, part);
    const text = all.map((p) => out.text(p)).join('|') + out.text();
    for (const f of tpl.fields) assert.ok(text.includes(`<${f}>`), f);
  });

  it('can be called incrementally', () => {
    const doc = template('test_merge_table_rows.docx').createDocument();
    doc.merge({ student_name: 'Bouke Haarsma' });
    assert.ok(!doc.getMergeFields().has('student_name'));
    assert.ok(doc.getMergeFields().has('thesis_grade'));
  });
});

describe('table rows', () => {
  const rows = [
    { class_code: 'ECON101', class_name: 'Economics 101', class_grade: 'A' },
    { class_code: 'ECONADV', class_name: 'Economics Advanced', class_grade: 'B' },
    { class_code: 'OPRES', class_name: 'Operations Research', class_grade: 'A' },
  ];

  it('mergeRows() repeats the anchor row (Python test_merge_rows)', () => {
    const doc = template('test_merge_table_rows.docx').createDocument();
    doc.merge({ student_name: 'Bouke Haarsma', thesis_grade: 'A' }).mergeRows('class_code', rows);
    const text = new Output(doc.toBuffer()).text();
    assert.match(text, /ECON101Economics 101A.*ECONADVEconomics AdvancedB.*OPRESOperations ResearchA/s);
  });

  it('array values in merge() do the same', () => {
    const a = new Output(template('test_merge_table_rows.docx').render({ class_code: rows })).text();
    const doc = template('test_merge_table_rows.docx').createDocument().mergeRows('class_code', rows);
    assert.equal(a, new Output(doc.toBuffer()).text());
  });

  it('removeEmptyTables', () => {
    const keep = new Output(template('test_merge_table_rows.docx').render({ class_code: [] }));
    const drop = new Output(template('test_merge_table_rows.docx', { removeEmptyTables: true }).render({ class_code: [] }));
    assert.equal(descendants(drop.xml(), NS.w, 'tbl').length, descendants(keep.xml(), NS.w, 'tbl').length - 1);
  });

  it('tableRowsReplaceMode overwrites rows and throws on overflow', () => {
    const tpl = template('test_merge_table_rows.docx', { tableRowsReplaceMode: true });
    const many = Array.from({ length: 50 }, (_, i) => ({ class_code: `C${i}` }));
    assert.throws(() => tpl.render({ class_code: many }), RangeError);
    const out = new Output(tpl.render({ class_code: rows.slice(0, 1) }));
    assert.ok(out.text().includes('ECON101'));
  });

  it('nested tables', () => {
    const tpl = template('test_merge_nested_table_rows.docx');
    assert.ok(tpl.fields.size > 0);
    const out = new Output(tpl.render({}));
    assert.equal(out.liveFields(), 0);
  });
});

describe('mergeTemplates() / renderAll()', () => {
  it('page_break between records (Python test_break_page)', () => {
    const out = new Output(template('test_merge_templates_simple.docx').renderAll(
      [{ fieldname: 'Test with page_break' }, { fieldname: 'abc' }, { fieldname: '2b v ~2b' }],
    ));
    assert.deepEqual(out.texts().filter((t) => !['Merge :', ' '].includes(t)), ['Test with page_break', 'abc', '2b v ~2b']);
    assert.equal(out.pageBreaks(), 2);
    assert.equal(descendants(out.xml(), NS.w, 'sectPr').length, 1);
  });

  it('section separators duplicate headers/footers with their relationships', () => {
    const tpl = template('test_footnote_header_footer.docx');
    const records = [0, 1, 2].map((i) => Object.fromEntries([...tpl.fields].map((f) => [f, `${f}-${i}`])));
    const out = new Output(tpl.renderAll(records, { separator: 'nextPage_section' }));
    const doc = out.xml();
    assert.equal(descendants(doc, NS.w, 'sectPr').length, 3);
    const rels = out.xml('word/_rels/document.xml.rels');
    const ids = new Map(descendants(rels, NS.rr, 'Relationship').map((r) => [r.getAttribute('Id'), r.getAttribute('Target')]));
    const refs = descendants(doc, NS.w, 'headerReference').map((r) => r.getAttributeNS(NS.r, 'id'));
    for (const id of refs) {
      const target = ids.get(id);
      assert.ok(target && out.files[`word/${target}`], `header ${id} -> ${target} exists`);
      assert.ok(out.raw('[Content_Types].xml').includes(`/word/${target}`), `${target} has a content type`);
    }
    assert.equal(new Set(refs).size, refs.length, 'each section has its own header copy');
  });

  it('drawing ids stay unique when content is duplicated', () => {
    const out = new Output(template('test_input_duplicate_id.docx').renderAll([{}, {}, {}]));
    const ids = descendants(out.xml(), NS.wp, 'docPr').map((d) => d.getAttribute('id'));
    assert.ok(ids.length >= 3);
    assert.equal(new Set(ids).size, ids.length);
  });

  it('NEXT moves to the next record (Python test_next_record)', () => {
    const out = new Output(template('test_next_record.docx').renderAll(values.map((field) => ({ field }))));
    const expected = [...values, '', '', ''].flatMap((v) => [v, '/', v]);
    assert.deepEqual(out.texts(), expected);
  });

  it('NEXT outside mergeTemplates is ignored instead of crashing (Python AssertionError)', () => {
    assert.doesNotThrow(() => template('test_next_record.docx').render({ field: 'x' }));
  });

  it('NEXTIF (Python test_next_if_record)', () => {
    const out = new Output(template('test_nextif.docx').renderAll(values.map((fieldname) => ({ fieldname }))));
    assert.deepEqual(out.texts(), ['one', 'two', 'three', 'three', 'four', 'four', 'five', 'five']);
  });

  it('SKIPIF drops records - without a leading blank page (Python bug)', () => {
    const out = new Output(template('test_skipif.docx').renderAll(values.map((fieldname) => ({ fieldname }))));
    assert.deepEqual(out.texts(), ['two', 'two', 'three', 'three', 'four', 'four', 'five', 'five']);
    const first = descendants(out.xml(), NS.w, 'p')[0]!;
    assert.equal(descendants(first, NS.w, 'br').length, 0, 'document must not start with a page break');
    assert.equal(out.pageBreaks(), 3);
  });

  it('rejects bad input', () => {
    const tpl = template('test_merge_templates_simple.docx');
    assert.throws(() => tpl.renderAll([]), RangeError);
    // @ts-expect-error invalid separator
    assert.throws(() => tpl.renderAll([{}], { separator: 'nope' }), TypeError);
    const doc = tpl.createDocument().mergeTemplates([{}]);
    assert.throws(() => doc.mergeTemplates([{}]), /only be called once/);
  });
});

describe('IF fields (mergeIfFields)', () => {
  it('evaluates conditions (Python test_if_record)', () => {
    const out = new Output(template('test_if_with_paragraph.docx', { mergeIfFields: true })
      .renderAll(values.map((fieldname) => ({ fieldname }))));
    assert.deepEqual(out.texts(), values.flatMap((v) => [v, v + (v === 'one' ? ' is ' : ' is not '), 'one']));
  });

  it('missing values merge as empty (Python test_if_record_with_missing_value)', () => {
    const out = new Output(template('test_if_with_paragraph.docx', { mergeIfFields: true })
      .renderAll([{}, ...values.map((fieldname) => ({ fieldname })), {}]));
    assert.deepEqual(out.texts(), ['', ...values, ''].flatMap((v) => [v, v + (v === 'one' ? ' is ' : ' is not '), 'one']));
  });

  it('left to Word when disabled', () => {
    const tpl = template('test_if_with_paragraph.docx');
    assert.ok(!tpl.describeFields().some((f) => f.type === 'IF'));
  });
});

describe('keepFields', () => {
  it("'some' keeps fields without data as live fields", () => {
    const tpl = template('test_keep_fields.docx', { keepFields: 'some' });
    assert.deepEqual([...tpl.fields].sort(), ['first', 'four_simple', 'second', 'three_simple']);
    assert.equal(new Output(tpl.render({})).liveFields(), 4);
    const out = new Output(tpl.render({ second: 'VALUE' }));
    assert.ok(out.text().includes('VALUE'));
    assert.equal(out.liveFields(), 3);
  });

  it("'all' keeps every field and updates its result", () => {
    const tpl = template('test_keep_fields.docx', { keepFields: 'all' });
    const out = new Output(tpl.render({ first: '<1>', second: '<2>', three_simple: '<3>', four_simple: '<4>' }));
    assert.equal(out.liveFields(), 4);
    for (const v of ['<1>', '<2>', '<3>', '<4>']) assert.ok(out.text().includes(v));
  });

  it("'none' uses emptyValue", () => {
    const out = new Output(template('test_merge_templates_simple.docx').render({}, { emptyValue: 'N/A' }));
    assert.equal(out.text(), 'Merge : N/A');
  });
});

describe('nested fields', () => {
  it('evaluates IF with nested MERGEFIELD (Python test_field_outside)', () => {
    const out = new Output(template('test_nested_if_outside.docx', { mergeIfFields: true }).render({ fieldname: 'one' }));
    assert.deepEqual(out.texts(), ['true', 'one']);
    assert.equal(out.liveFields(), 0);
  });

  it('fills MERGEFIELDs inside other fields and sets updateFields when asked', () => {
    const tpl = template('test_nested_if_inside.docx', { autoUpdateFieldsOnOpen: 'auto' });
    assert.ok(tpl.describeFields().some((f) => f.nested && f.name === 'fieldname'));
    const out = new Output(tpl.render({ fieldname: 'one' }));
    assert.ok(out.raw().includes('one'), 'value written into the outer field instruction');
    const settings = out.raw('word/settings.xml');
    assert.equal(settings.match(/w:updateFields/g)?.length, 1);
    // schema order: updateFields must precede compat/rsids
    const pos = settings.indexOf('w:updateFields');
    for (const later of ['<w:compat', '<w:rsids']) {
      const i = settings.indexOf(later);
      if (i >= 0) assert.ok(pos < i, `updateFields before ${later}`);
    }
  });
  it("'auto' does nothing without nested fields; 'no' leaves settings alone", () => {
    const auto = new Output(template('test_nested_if_outside.docx', { autoUpdateFieldsOnOpen: 'auto' }).render({ fieldname: 'one' }));
    assert.ok(!auto.raw('word/settings.xml').includes('updateFields'));
    const out = new Output(template('test_nested_if_inside.docx').render({ fieldname: 'one' }));
    assert.ok(!out.raw('word/settings.xml').includes('updateFields'));
  });
});
