import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { applySwitches, formatDate, formatNumber, toText, type FormatContext } from '../src/format.ts';
import { resolveOptions, type MailMergeOptions } from '../src/options.ts';
import { tokenizeInstruction } from '../src/tokenize.ts';

const ctx = (o: MailMergeOptions = {}): FormatContext => {
  const options = resolveOptions({ timeZone: 'UTC', ...o });
  return { options, locale: options.locale, timeZone: options.timeZone, field: 'f' };
};
const fmt = (value: unknown, instr: string, o?: MailMergeOptions): unknown =>
  applySwitches(value, tokenizeInstruction(instr).slice(2), ctx(o));

describe('tokenizeInstruction (shlex-compatible)', () => {
  it('splits, unquotes and keeps empty strings', () => {
    assert.deepEqual(tokenizeInstruction('IF "a b" = "" "yes" "no"'), ['IF', 'a b', '=', '', 'yes', 'no']);
    assert.deepEqual(tokenizeInstruction('MERGEFIELD a"b c"d'), ['MERGEFIELD', 'ab cd']);
    assert.deepEqual(tokenizeInstruction(' DATE @"MMMM d\', \'yyyy" '), ['DATE', "@MMMM d', 'yyyy"]);
  });
  it('recovers from an unbalanced quote instead of throwing', () => {
    const warnings: string[] = [];
    assert.deepEqual(tokenizeInstruction('MERGEFIELD d \\@ "dd.MM', (w) => warnings.push(w)), ['MERGEFIELD', 'd', '\\@', 'dd.MM']);
    assert.equal(warnings.length, 1);
  });
});

describe('\\# numeric pictures (vectors from the Python test-suite)', () => {
  const cases: Record<string, [unknown, string][]> = {
    '0.00': [[25, '25.00'], [0, '0.00'], [345.12314, '345.12'], [null, '0.00']],
    '#,#00': [[25, '25'], [0, '00'], [32345.12314, '32,345']],
    '#,###.##': [[0, '0.00'], [23423, '23,423.00']],
    "#'###.##": [[0, '0.00'], [23423, "23'423.00"]],
    '#,##0.00': [[0, '0.00'], [23423, '23,423.00'], [-1234.5, '-1,234.50']],
    N3: [[25, '25.000'], [0, '0.000'], [345.12314, '345.123']],
    P3: [[25, '2500.000%'], [0, '0.000%'], [345.12314, '34512.314%']],
    '##%': [[0.25, '25%'], [0, '0%'], [0.034512314, '3%']],
    '$#,##0.00 USD': [[1234.5, '$1,234.50 USD']],
  };
  for (const [picture, vectors] of Object.entries(cases)) {
    it(picture, () => {
      for (const [value, expected] of vectors) assert.equal(formatNumber(value, picture), expected);
    });
  }
  it('accepts numeric strings and bigint (CSV/JSON friendly)', () => {
    assert.equal(formatNumber('1234.5', '#,##0.00'), '1,234.50');
    assert.equal(formatNumber(12n, '0.0'), '12.0');
  });
  it('rejects non-numeric values', () => assert.throws(() => formatNumber('abc', '0.00')));
});

describe('\\@ date pictures', () => {
  const d = new Date(Date.UTC(2022, 2, 9, 17, 7, 8));
  const c = ctx();
  const cases: Record<string, string> = {
    M: '3', MM: '03', MMM: 'Mar', MMMM: 'March', d: '9', dd: '09', ddd: 'Wed', dddd: 'Wednesday',
    D: '9', DD: '09', yy: '22', yyyy: '2022', YYYYMMDD: '20220309',
    h: '5', hh: '05', H: '17', HH: '17', m: '7', mm: '07', s: '8', ss: '08', 'am/pm': 'PM',
    "dddd, d 'de' MMMM": 'Wednesday, 9 de March',
  };
  for (const [picture, expected] of Object.entries(cases)) {
    it(picture, () => assert.equal(formatDate(d, picture, c), expected));
  }
  it('12 o\'clock is 12, not 0 (Python bug)', () => {
    assert.equal(formatDate(new Date(Date.UTC(2022, 0, 1, 12, 0)), 'h am/pm', c), '12 PM');
    assert.equal(formatDate(new Date(Date.UTC(2022, 0, 1, 0, 30)), 'h:mm am/pm', c), '12:30 AM');
  });
  it('is locale-aware without touching process locale', () => {
    assert.equal(formatDate(d, 'MMMM', ctx({ locale: 'de-DE' })), 'März');
    assert.equal(formatDate(d, 'dddd', ctx({ locale: 'fr-FR' })), 'mercredi');
  });
  it('honours timeZone', () => {
    assert.equal(formatDate(d, 'HH:mm', ctx({ timeZone: 'Asia/Tokyo' })), '02:07');
  });
  it('parses ISO strings; date-only strings are local dates', () => {
    assert.equal(formatDate('2022-03-09', 'dd/MM/yyyy', ctx({ timeZone: undefined })), '09/03/2022');
    assert.equal(formatDate('2022-03-09T17:07:08Z', 'HH:mm', c), '17:07');
    assert.equal(formatDate('not a date', 'yyyy', c), 'not a date');
    assert.equal(formatDate(null, 'yyyy', c), '');
  });
  it('formats Date values without a switch using the locale', () => {
    const c2 = ctx({ locale: 'en-GB' });
    assert.equal(toText(new Date(Date.UTC(2022, 2, 9)), c2), '09/03/2022');
    assert.equal(toText(new Date(Date.UTC(2022, 2, 9, 17, 7, 8)), c2), '09/03/2022, 17:07:08');
  });
});

describe('\\* text, \\b, \\f', () => {
  it('Caps / FirstCap / Upper / Lower / unknown', () => {
    assert.equal(fmt('handsome law group', 'MERGEFIELD f \\* Caps'), 'Handsome Law Group');
    assert.equal(fmt('HANDSOME  LAW GROUP', 'MERGEFIELD f \\* Caps'), 'Handsome  Law Group');
    assert.equal(fmt('handsome law group', 'MERGEFIELD f \\* FirstCap'), 'Handsome law group');
    assert.equal(fmt('Handsome Law Group', 'MERGEFIELD f \\* Upper'), 'HANDSOME LAW GROUP');
    assert.equal(fmt('Handsome Law Group', 'MERGEFIELD f \\* Lower'), 'handsome law group');
    assert.equal(fmt('handsome Law Group', 'MERGEFIELD f \\* Invalid'), 'handsome Law Group');
    assert.equal(fmt('x', 'MERGEFIELD f \\* MERGEFORMAT'), 'x');
  });
  it('locale-aware casing (Turkish dotted i)', () => {
    assert.equal(fmt('istanbul', 'MERGEFIELD f \\* Upper', { locale: 'tr-TR' }), 'İSTANBUL');
  });
  it('text before/after only when there is a value', () => {
    assert.equal(fmt('law', 'MERGEFIELD f \\b " - "'), ' - law');
    assert.equal(fmt('law', 'MERGEFIELD f \\f " - "'), 'law - ');
    assert.equal(fmt('', 'MERGEFIELD f \\b " - "'), null);
    assert.equal(fmt(2, 'MERGEFIELD f \\b "(" \\f ")"'), '(2)');
    assert.equal(fmt(0, 'MERGEFIELD f \\b "(" \\f ")"'), null);
  });
  it('custom formatters', () => {
    const formatters = { money: (v: unknown) => `€${Number(v).toFixed(2)}` };
    assert.equal(fmt(3, 'MERGEFIELD f \\* Money', { formatters }), '€3.00');
  });
  it('a switch without argument throws (caller warns and keeps the raw value)', () => {
    assert.throws(() => fmt(2, 'MERGEFIELD f \\b'));
  });
});
