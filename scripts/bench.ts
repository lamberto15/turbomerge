/**
 * npm run bench [-- <count>]
 * Generates <count> separate documents from one template. If ../.venv exists, runs the same
 * workload with the Python docx-mailmerge2 for comparison.
 */
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { DocxTemplate } from '../src/index.ts';

const count = Number(process.argv[2] ?? 500);
const fixture = join(import.meta.dirname, '../test/fixtures/test_merge_table_rows.docx');
const record = (i: number) => ({
  student_name: `S${i}`,
  thesis_grade: 'A',
  class_code: [
    { class_code: 'C1', class_name: 'N', class_grade: 'B' },
    { class_code: 'C2', class_name: 'M', class_grade: 'A' },
  ],
});

const t0 = performance.now();
const template = DocxTemplate.fromBuffer(readFileSync(fixture));
let bytes = 0;
for (const doc of template.renderEach(Array.from({ length: count }, (_, i) => record(i)))) bytes += doc.bytes.length;
const nodeMs = performance.now() - t0;
console.log(
  `node   ${count} docs: ${nodeMs.toFixed(0)} ms (${(nodeMs / count).toFixed(2)} ms/doc, ${(bytes / 1e6).toFixed(1)} MB)`,
);

const python = join(import.meta.dirname, '../../.venv/bin/python');
if (existsSync(python)) {
  const script = `
import io, time, warnings, json, sys
warnings.simplefilter("ignore")
from mailmerge import MailMerge
n = int(sys.argv[2]); t = time.perf_counter()
for i in range(n):
    with MailMerge(sys.argv[1]) as d:
        d.merge(student_name=f"S{i}", thesis_grade="A", class_code=[{"class_code":"C1","class_name":"N","class_grade":"B"},{"class_code":"C2","class_name":"M","class_grade":"A"}])
        d.write(io.BytesIO())
print((time.perf_counter() - t) * 1000)`;
  const pyMs = Number(execFileSync(python, ['-c', script, fixture, String(count)], { encoding: 'utf8' }));
  console.log(`python ${count} docs: ${pyMs.toFixed(0)} ms (${(pyMs / count).toFixed(2)} ms/doc)`);
  console.log(`node is ${(pyMs / nodeMs).toFixed(1)}x faster`);
}
