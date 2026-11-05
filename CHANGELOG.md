# Changelog

## 0.1.1

### Patch Changes

- 526d629: Add npm, CI and license badges to the README.

## 0.1.0 — unreleased

First release: a Node.js port of the Python docx-mailmerge2 1.0.3.

- `DocxTemplate` (parse once, render many) with `render`, `renderAll`, `renderEach`, `createDocument`
- MERGEFIELD, NEXT, NEXTIF, SKIPIF, IF (opt-in), nested fields, simple and complex field forms
- Table rows (insert or replace mode, nested tables, empty-table removal)
- Multi-record documents with break or section separators, and per-record headers/footers
- `\#`, `\@`, `\*`, `\b`, `\f` switches; locale and time zone through `Intl`; custom `\*` formatters
- Dot-path data, ISO date strings, `validate()`, `strict` mode with `MissingFieldsError`
- `describeFields()`, `onWarning`
- Native zlib zip I/O with passthrough of unchanged entries; deterministic output
