# docx-mailmerge

Fast, dependency-light **mail merge for Word `.docx` files** in Node.js. Use real Word merge fields
(`MERGEFIELD`, `IF`, `NEXT`, `NEXTIF`, `SKIPIF`) and Word's formatting switches, without Microsoft Word.

- **Native Word fields.** Templates are made in Word (Insert → Field → MergeField), so no `{{custom}}` syntax.
- **Parse once, render many.** One template renders thousands of documents (~0.6 ms per document).
- **Tables, multi-record documents and one file per record**
- **Formatting:** `\# "#,##0.00"`, `\@ "d MMMM yyyy"`, `\* Upper`, `\b`/`\f`, plus custom formatters
- **Dot paths** for nested JSON (`MERGEFIELD customer.address.city`), and validation with a `strict` mode
- Headers, footers, footnotes and endnotes
- ESM, TypeScript types included, **Node.js ≥ 24**, one runtime dependency (`@xmldom/xmldom`)

A Node.js port of the Python [docx-mailmerge2](https://github.com/iulica/docx-mailmerge). Its output is
checked against the Python implementation on every fixture of the Python test suite.

## Install

```sh
npm install docx-mailmerge
```

## Quick start

```js
import { writeFile } from 'node:fs/promises';
import { DocxTemplate } from 'docx-mailmerge';

const template = await DocxTemplate.load('letter.docx');

console.log(template.fields); // Set { 'name', 'amount', 'items', ... }

const bytes = template.render({
  name: 'Ada Lovelace',
  amount: 1234.5,
  items: [
    // an array fills the table row containing «items»
    { items: 'A-1', qty: 2 },
    { items: 'B-7', qty: 1 },
  ],
});
await writeFile('letter-ada.docx', bytes);
```

CommonJS works too, because Node ≥ 24 can `require()` ES modules: `const { DocxTemplate } = require('docx-mailmerge')`.

## Rendering modes

| Method                                       | Output                                                                                          |
| -------------------------------------------- | ----------------------------------------------------------------------------------------------- |
| `template.render(record)`                    | one `.docx` for one record                                                                      |
| `template.renderAll(records, { separator })` | one `.docx` that repeats the body once per record                                               |
| `template.renderEach(records)`               | a lazy generator with one `.docx` per record                                                    |
| `template.createDocument()`                  | an incremental `MailMergeDocument` (`merge`, `mergeRows`, `mergeTemplates`, `toBuffer`, `save`) |

```js
// one file per customer, streamed: memory stays flat for any number of records
for (const { index, bytes } of template.renderEach(customers)) {
  await writeFile(`out/letter-${index + 1}.docx`, bytes);
}

// a single document with a page per customer
await writeFile('all.docx', template.renderAll(customers, { separator: 'page_break' }));
```

Separators: `page_break`, `column_break`, `textWrapping_break`, `continuous_section`, `evenPage_section`,
`nextColumn_section`, `nextPage_section`, `oddPage_section`. With a `*_section` separator, headers and
footers that contain fields are copied per record, so each record gets its own header.

## Data

- **Scalars** (`string`, `number`, `bigint`, `boolean`, `Date`) fill fields. `\n` becomes a line break.
- **Arrays of objects** repeat a table row. The key must be the name of a field inside that row. Rows can contain nested tables.
- **Nested objects** are reachable with dot paths: `{ customer: { name: 'Ada' } }` fills `MERGEFIELD customer.name`. An exact key `'customer.name'` takes precedence.
- **ISO date strings** (`'2026-11-05'`, `'2026-11-05T10:00:00Z'`) work with `\@` pictures, which helps with JSON data. A date-only string means a local calendar date.

### Validation

```js
template.validate({ name: 'Ada', extra: 1 });
// { valid: false, missing: ['amount', 'items'], unused: ['extra'] }

template.render(record, { strict: true }); // throws MissingFieldsError (with .missing, .recordIndex)
```

## Field formatting

| Switch       | Example                                 | Result                                                   |
| ------------ | --------------------------------------- | -------------------------------------------------------- |
| `\#` numeric | `MERGEFIELD total \# "#,##0.00"`        | `1,234.50`                                               |
|              | `\# "N2"` / `\# "P1"` / `\# "0%"`       | fixed decimals / percent                                 |
| `\@` date    | `MERGEFIELD due \@ "dddd, d MMMM yyyy"` | `Thursday, 5 November 2026`                              |
| `\*` text    | `Upper`, `Lower`, `Caps`, `FirstCap`    | case conversion (locale-aware)                           |
| `\b` / `\f`  | `MERGEFIELD title \b "Re: "`            | adds text before/after, only when the value is non-empty |

Dates use `Intl` with the `locale` and `timeZone` options. Nothing depends on the process locale.

### Custom formatters

```js
const template = await DocxTemplate.load('invoice.docx', {
  formatters: {
    money: (value, { locale }) =>
      new Intl.NumberFormat(locale, { style: 'currency', currency: 'EUR' }).format(Number(value)),
  },
});
// MERGEFIELD price \* Money   ->   €1,200.00
```

## Conditional fields

- `NEXT` continues with the next record on the same page (labels, lists). Use it with `renderAll`.
- `NEXTIF a = b` / `SKIPIF a = b` move to the next record, or drop the current record, when the condition holds.
- `IF a op b "then" "else"` is evaluated when `mergeIfFields: true`. Otherwise it is left for Word to compute.

Operators: `=`, `<>`, `<`, `<=`, `>`, `>=`. Numbers compare numerically. `=` and `<>` accept Word wildcards
`*` and `?`. Merge fields can be nested inside other fields, for example
`{ IF "{ MERGEFIELD paid }" = "yes" "Thanks" "Please pay" }` or `{ INCLUDEPICTURE "{ MERGEFIELD photo }" }`.
For fields Word has to compute (like `INCLUDEPICTURE`), set `autoUpdateFieldsOnOpen: 'auto'`.

## Options

```ts
DocxTemplate.load(source, {
  keepFields: 'none', // 'none' | 'some' (keep fields without data) | 'all' (keep every field live)
  mergeIfFields: false, // evaluate IF fields
  autoUpdateFieldsOnOpen: 'no', // 'no' | 'auto' (only if nested fields) | 'always'
  removeEmptyTables: false, // drop a table whose rows array is empty
  tableRowsReplaceMode: false, // overwrite existing rows instead of inserting (throws on overflow)
  locale: 'en-US',
  timeZone: undefined, // e.g. 'Europe/Lisbon'; default: process time zone
  formatters: {},
  strict: false,
  onWarning: (w) => {}, // default: process.emitWarning
});
```

`source` can be a file path, `URL`, `Uint8Array`/`Buffer`, `ArrayBuffer` or `Blob`. Use
`DocxTemplate.fromBuffer(bytes)` for synchronous creation. `toBuffer({ emptyValue: 'N/A' })` sets the text
used for fields without data.

## Inspecting a template

```js
template.describeFields();
// [{ name: 'total', type: 'MERGEFIELD', instruction: 'MERGEFIELD total \\# "#,##0.00"',
//    part: 'word/document.xml', nested: false }, ...]
```

## Example without Word

[`examples/invoice-template.ts`](examples/invoice-template.ts) builds a complete invoice template in
code (fields, nested IF, table row, formatting) and renders it:

```sh
node examples/invoice-template.ts
```

## Performance

Parsing happens once per template. Each render copies only the XML parts that contain fields, and
unchanged zip entries are copied byte-for-byte without being recompressed. The remaining compression uses
Node's native zlib. Output is deterministic: the same input gives byte-identical files.

```
$ npm run bench -- 1000      # 1000 separate documents, Apple M-series
node   1000 docs: 585 ms (0.58 ms/doc)
python 1000 docs: 2580 ms (2.58 ms/doc)   # docx-mailmerge2 1.0.3 + lxml
```

## Differences from the Python library

The API is idiomatic JS (`render*`, camelCase options). These Python behaviours are fixed:

- `<`, `>`, `<=`, `>=` in `IF`/`NEXTIF`/`SKIPIF` compare numbers numerically. Python compared strings, so `"10" < "9"` was true.
- A template with `NEXT` used with plain `merge()` works. Python raised `AssertionError`.
- `SKIPIF` on the first record no longer leaves a blank first page.
- Per-record header/footer copies also copy their relationships, so images in headers keep working.
- `h` / `hh` show 12 for noon and midnight. Python showed `0`.
- Leading and trailing spaces in values are preserved (`xml:space="preserve"`), and a simple field keeps its run formatting (bold, font…).
- `w:type` / `w:updateFields` are inserted in schema order, and `updateFields` is never duplicated.
- A malformed instruction (unbalanced quote) produces a warning instead of an exception.
- Field type names are case-insensitive (`mergefield` works).
- Dates render through `Intl` with explicit `locale`/`timeZone` instead of the process-global C locale.

## Development

```sh
npm install
npm test            # node:test, runs the TypeScript sources directly
npm run typecheck
npm run build       # dist/ (ESM + .d.ts)
npm run parity      # differential test vs the Python implementation (needs ../.venv)
npm run bench
```

## License

MIT. Includes work derived from docx-mailmerge (© 2012 Bouke Haarsma) and docx-mailmerge2 (© 2023 Iulian Ciorăscu).
