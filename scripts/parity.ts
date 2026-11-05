/**
 * Differential test against the original Python implementation.
 *
 *   npm run parity            (needs ../.venv with the Python package installed)
 *
 * 1. renders every fixture x scenario with Python (reference) and with this library
 * 2. extracts what a reader sees from each output: body text, the text of the headers/footers
 *    actually referenced, the number of live fields left, and the page/section breaks
 * 3. reports every difference
 */
import { execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync } from 'fflate';
import { DocxTemplate } from '../src/index.ts';
import { NS } from '../src/constants.ts';
import { parseXml, descendants, wAttr } from '../src/xml.ts';

const root = join(import.meta.dirname, '..');
const fixtures = join(root, 'test/fixtures');
const outDir = join(root, '.parity');
const python = join(root, '../.venv/bin/python');
const scenarios = JSON.parse(readFileSync(join(import.meta.dirname, 'parity/scenarios.json'), 'utf8')) as {
  id: string; mode: 'merge' | 'templates'; records?: number; partial?: boolean;
  separator?: 'page_break'; options: Record<string, unknown>;
}[];

rmSync(outDir, { recursive: true, force: true });
mkdirSync(join(outDir, 'node'), { recursive: true });
console.log(execFileSync(python, [join(import.meta.dirname, 'parity/run_python.py'), fixtures, join(outDir, 'python')], { encoding: 'utf8' }).trim());

const record = (names: Iterable<string>, i: number, partial?: boolean) => {
  let sorted = [...names].sort();
  if (partial) sorted = sorted.filter((_, k) => k % 2 === 0);
  return Object.fromEntries(sorted.map((n) => [n, `${n}#${i}`]));
};

const nodeErrors: Record<string, string> = {};
for (const file of readdirSync(fixtures).filter((f) => f.endsWith('.docx')).sort()) {
  for (const sc of scenarios) {
    const key = `${file.replace(/\.docx$/, '')}__${sc.id}`;
    try {
      const tpl = DocxTemplate.fromBuffer(readFileSync(join(fixtures, file)), { ...sc.options, onWarning: () => {} });
      const doc = tpl.createDocument();
      const names = doc.getMergeFields();
      if (sc.mode === 'merge') doc.merge(record(names, 0, sc.partial));
      else doc.mergeTemplates(Array.from({ length: sc.records! }, (_, i) => record(names, i, sc.partial)), sc.separator);
      writeFileSync(join(outDir, 'node', `${key}.docx`), doc.toBuffer());
    } catch (error) {
      nodeErrors[key] = `${(error as Error).name}: ${(error as Error).message}`;
    }
  }
}

interface Summary { body: string[]; headers: string[]; liveFields: number; breaks: string[]; }

function summarize(path: string): Summary {
  const zip = unzipSync(readFileSync(path));
  const xml = (name: string) => parseXml(zip[name]!, name).documentElement!;
  const doc = xml('word/document.xml');
  const texts = (el: ReturnType<typeof xml>) => descendants(el, NS.w, 't').map((t) => t.textContent ?? '');
  const live = (el: ReturnType<typeof xml>) =>
    descendants(el, NS.w, 'fldSimple').length +
    descendants(el, NS.w, 'fldChar').filter((f) => wAttr(f, 'fldCharType') === 'begin').length;

  const rels = zip['word/_rels/document.xml.rels'] ? xml('word/_rels/document.xml.rels') : null;
  const target = new Map(rels ? descendants(rels, NS.rr, 'Relationship').map((r) => [r.getAttribute('Id'), r.getAttribute('Target')]) : []);
  const headers: string[] = [];
  let liveFields = live(doc);
  for (const ref of [...descendants(doc, NS.w, 'headerReference'), ...descendants(doc, NS.w, 'footerReference')]) {
    const t = target.get(ref.getAttributeNS(NS.r, 'id'));
    const part = t && zip[`word/${t}`] ? xml(`word/${t}`) : null;
    if (part) {
      headers.push(texts(part).join(''));
      liveFields += live(part);
    }
  }
  const breaks = [
    ...descendants(doc, NS.w, 'br').map((b) => `br:${wAttr(b, 'type') ?? 'line'}`),
    ...descendants(doc, NS.w, 'sectPr').map((s) => `sect:${wAttr(descendants(s, NS.w, 'type')[0] ?? s, 'val') ?? '-'}`),
  ];
  return { body: texts(doc), headers, liveFields, breaks };
}

const pyErrors = JSON.parse(readFileSync(join(outDir, 'python/errors.json'), 'utf8')) as Record<string, string>;
const keys = new Set([...readdirSync(join(outDir, 'python')), ...readdirSync(join(outDir, 'node'))]
  .filter((f) => f.endsWith('.docx')).map((f) => f.slice(0, -5)));
for (const k of [...Object.keys(pyErrors), ...Object.keys(nodeErrors)]) keys.add(k);

let same = 0;
const diffs: string[] = [];
for (const key of [...keys].sort()) {
  const py = existsSync(join(outDir, 'python', `${key}.docx`)) ? summarize(join(outDir, 'python', `${key}.docx`)) : pyErrors[key];
  const nd = existsSync(join(outDir, 'node', `${key}.docx`)) ? summarize(join(outDir, 'node', `${key}.docx`)) : nodeErrors[key];
  const a = JSON.stringify(py);
  const b = JSON.stringify(nd);
  if (a === b) { same += 1; continue; }
  const lines = [`✗ ${key}`];
  if (typeof py !== 'object' || typeof nd !== 'object') lines.push(`    python: ${typeof py === 'string' ? py : 'ok'}`, `    node:   ${typeof nd === 'string' ? nd : 'ok'}`);
  else for (const f of ['body', 'headers', 'liveFields', 'breaks'] as const) {
    if (JSON.stringify(py[f]) !== JSON.stringify(nd[f])) lines.push(`    ${f}\n      python: ${JSON.stringify(py[f])}\n      node:   ${JSON.stringify(nd[f])}`);
  }
  diffs.push(lines.join('\n'));
}
console.log(diffs.join('\n'));
console.log(`\nparity: ${same}/${keys.size} identical, ${diffs.length} different`);
