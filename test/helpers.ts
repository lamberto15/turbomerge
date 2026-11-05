import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync, zipSync } from 'fflate';
import { NS } from '../src/constants.ts';
import { DocxTemplate, type MailMergeOptions } from '../src/index.ts';
import { descendants, type Element, parseXml, wAttr } from '../src/xml.ts';

export const fixture = (name: string): Uint8Array => readFileSync(join(import.meta.dirname, 'fixtures', name));

export const template = (name: string, options: MailMergeOptions = {}): DocxTemplate =>
  DocxTemplate.fromBuffer(fixture(name), { onWarning: () => {}, ...options });

export class Output {
  readonly files: Record<string, Uint8Array>;
  constructor(bytes: Uint8Array) {
    this.files = unzipSync(bytes);
  }
  xml(name = 'word/document.xml'): Element {
    const data = this.files[name];
    if (!data) throw new Error(`missing part ${name}`);
    return parseXml(data, name).documentElement!;
  }
  raw(name = 'word/document.xml'): string {
    return new TextDecoder().decode(this.files[name]);
  }
  texts(name?: string): string[] {
    return descendants(this.xml(name), NS.w, 't').map((t) => t.textContent ?? '');
  }
  text(name?: string): string {
    return this.texts(name).join('');
  }
  liveFields(name?: string): number {
    const root = this.xml(name);
    return (
      descendants(root, NS.w, 'fldSimple').length +
      descendants(root, NS.w, 'fldChar').filter((f) => wAttr(f, 'fldCharType') === 'begin').length
    );
  }
  pageBreaks(name?: string): number {
    return descendants(this.xml(name), NS.w, 'br').filter((b) => wAttr(b, 'type') === 'page').length;
  }
}

/** test_one_simple_field.docx with its single field instruction replaced. */
export function withInstruction(instr: string, options: MailMergeOptions = {}): DocxTemplate {
  const files = unzipSync(fixture('test_one_simple_field.docx'));
  const xml = new TextDecoder().decode(files['word/document.xml']);
  const escaped = instr.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  files['word/document.xml'] = new TextEncoder().encode(xml.replace(/w:instr="[^"]*"/, `w:instr="${escaped}"`));
  return DocxTemplate.fromBuffer(zipSync(files), { onWarning: () => {}, ...options });
}
