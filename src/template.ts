import { readFile } from 'node:fs/promises';
import type { Separator } from './constants.ts';
import { MailMergeDocument, type WriteOptions } from './document.ts';
import type { FieldType } from './fields.ts';
import { type MailMergeOptions, type MergeRecord, type ResolvedOptions, resolveOptions } from './options.ts';
import { DocxPackage } from './package.ts';
import { FieldRegistry, parsePart } from './parse.ts';

export type TemplateSource = string | URL | Uint8Array | ArrayBuffer | Blob;

export interface FieldInfo {
  /** Data key the field reads (empty for NEXT / IF / NEXTIF / SKIPIF). */
  readonly name: string;
  readonly type: FieldType;
  /** Raw field instruction, e.g. `MERGEFIELD amount \# "#,##0.00"`. */
  readonly instruction: string;
  /** Zip path of the part containing the field, e.g. `word/header1.xml`. */
  readonly part: string;
  /** True when the field sits inside another field (e.g. IF, INCLUDEPICTURE). */
  readonly nested: boolean;
}

export interface ValidationResult {
  readonly valid: boolean;
  /** Merge fields the record does not provide (directly, via dot path, or in table rows). */
  readonly missing: string[];
  /** Top-level record keys no field uses. */
  readonly unused: string[];
}

export interface RenderOptions extends WriteOptions {
  /** Overrides the template's `strict` option for this call. */
  strict?: boolean;
}

export interface RenderAllOptions extends RenderOptions {
  /** How records are joined. Default `'page_break'`. */
  separator?: Separator;
}

export interface RenderedDocument {
  readonly index: number;
  readonly record: MergeRecord;
  readonly bytes: Uint8Array;
}

export class MissingFieldsError extends Error {
  override readonly name = 'MissingFieldsError';
  readonly missing: readonly string[];
  readonly recordIndex: number | undefined;

  constructor(missing: readonly string[], recordIndex?: number) {
    super(
      `${recordIndex === undefined ? 'Record' : `Record #${recordIndex}`} is missing merge field(s): ${missing.join(', ')}`,
    );
    this.missing = missing;
    this.recordIndex = recordIndex;
  }
}

const isPlainObject = (v: unknown): v is Record<string, unknown> =>
  v !== null && typeof v === 'object' && !Array.isArray(v) && !(v instanceof Date) && !ArrayBuffer.isView(v);

function collectPaths(record: Record<string, unknown>, prefix: string, out: Set<string>): void {
  for (const [key, value] of Object.entries(record)) {
    out.add(prefix + key);
    if (Array.isArray(value)) {
      for (const item of value) if (isPlainObject(item)) collectPaths(item, '', out); // table rows
    } else if (isPlainObject(value)) {
      collectPaths(value, `${prefix}${key}.`, out);
    }
  }
}

async function readSource(source: TemplateSource): Promise<Uint8Array> {
  if (typeof source === 'string' || source instanceof URL) return new Uint8Array(await readFile(source));
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  if (typeof Blob !== 'undefined' && source instanceof Blob) return new Uint8Array(await source.arrayBuffer());
  throw new TypeError('Template source must be a path, URL, Uint8Array, ArrayBuffer or Blob');
}

/**
 * A parsed .docx mail-merge template. Parsing (unzip, XML parse, field discovery) happens once;
 * every render works on a cheap in-memory clone, so one template can produce thousands of documents.
 *
 * @example
 * const template = await DocxTemplate.load('letter.docx');
 * const bytes = template.render({ name: 'Ada', items: [{ sku: 'A1', qty: 2 }] });
 */
export class DocxTemplate {
  readonly #pkg: DocxPackage;
  readonly #registry: FieldRegistry;
  readonly #options: ResolvedOptions;
  readonly #fieldNames: ReadonlySet<string>;

  private constructor(bytes: Uint8Array, options: MailMergeOptions) {
    this.#options = resolveOptions(options);
    this.#pkg = DocxPackage.open(bytes);
    this.#registry = new FieldRegistry(this.#options.mergeIfFields, (message) =>
      this.#options.onWarning({ code: 'INVALID_INSTRUCTION', message }),
    );
    for (const part of this.#pkg.partsOf('main', 'header_footer', 'notes')) {
      parsePart(part.read.documentElement!, part.path, this.#registry); // the template owns these docs
    }
    this.#pkg.indexFields();
    const names = new Set<string>();
    for (const f of this.#registry.all()) if (f.type === 'MERGEFIELD' && f.name) names.add(f.name);
    this.#fieldNames = names;
  }

  /** Loads a template from a file path, URL, bytes or Blob. */
  static async load(source: TemplateSource, options: MailMergeOptions = {}): Promise<DocxTemplate> {
    return new DocxTemplate(await readSource(source), options);
  }

  /** Synchronously creates a template from .docx bytes. */
  static fromBuffer(bytes: Uint8Array | ArrayBuffer, options: MailMergeOptions = {}): DocxTemplate {
    return new DocxTemplate(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), options);
  }

  /** Names of all MERGEFIELDs in the body, headers, footers, footnotes and endnotes. */
  get fields(): ReadonlySet<string> {
    return this.#fieldNames;
  }

  /** Every supported field with its type, instruction and location. */
  describeFields(): FieldInfo[] {
    return Array.from(this.#registry.all(), (f) => ({
      name: f.name,
      type: f.type,
      instruction: f.instr.trim(),
      part: f.part,
      nested: f.nested,
    }));
  }

  /** Checks a record against the template's fields without rendering. */
  validate(record: MergeRecord): ValidationResult {
    const available = new Set<string>();
    collectPaths(record as Record<string, unknown>, '', available);
    const missing = [...this.#fieldNames].filter((name) => !available.has(name));
    const names = [...this.#fieldNames];
    const unused = Object.keys(record).filter((key) => !names.some((n) => n === key || n.startsWith(`${key}.`)));
    return { valid: missing.length === 0, missing, unused };
  }

  /** Starts an output document for incremental merging. */
  createDocument(): MailMergeDocument {
    return new MailMergeDocument(this.#pkg.fork(), this.#registry, this.#options);
  }

  /** Renders one record into a complete .docx. */
  render(record: MergeRecord, options: RenderOptions = {}): Uint8Array {
    this.#check(record, options.strict);
    return this.createDocument().merge(record).toBuffer(options);
  }

  /** Renders all records into ONE .docx, one copy of the body per record. */
  renderAll(records: Iterable<MergeRecord>, options: RenderAllOptions = {}): Uint8Array {
    const rows = Array.from(records);
    rows.forEach((record, i) => {
      this.#check(record, options.strict, i);
    });
    return this.createDocument().mergeTemplates(rows, options.separator).toBuffer(options);
  }

  /** Lazily renders one separate .docx per record. Works with any iterable, including generators. */
  *renderEach(
    records: Iterable<MergeRecord>,
    options: RenderOptions = {},
  ): Generator<RenderedDocument, void, undefined> {
    let index = 0;
    for (const record of records) {
      this.#check(record, options.strict, index);
      yield { index, record, bytes: this.createDocument().merge(record).toBuffer(options) };
      index += 1;
    }
  }

  #check(record: MergeRecord, strict: boolean | undefined, index?: number): void {
    if (!(strict ?? this.#options.strict)) return;
    const { missing } = this.validate(record);
    if (missing.length > 0) throw new MissingFieldsError(missing, index);
  }
}
