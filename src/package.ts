import { CONTENT_TYPE_CATEGORIES, NS, STORED_EXTENSIONS, type PartCategory } from './constants.ts';
import { childrenNamed, cloneDocument, parseXml, placeholders, serializeXml, type Document, type Element } from './xml.ts';
import { inflate, makeEntry, readZip, writeZip, type ZipEntry } from './zip.ts';

const CONTENT_TYPES = '[Content_Types].xml';

export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash + 1)}_rels/${partPath.slice(slash + 1)}.rels`;
}

/** Copy-on-write view of a parsed XML part: reads hit the template, the first write clones it. */
export class XmlPart {
  readonly path: string;
  readonly #template: Document;
  #copy: Document | null = null;

  constructor(path: string, template: Document) {
    this.path = path;
    this.#template = template;
  }

  /** Read-only access (shared with the template - do not mutate). */
  get read(): Document {
    return this.#copy ?? this.#template;
  }

  /** Writable access; clones the template on first use. */
  get doc(): Document {
    return (this.#copy ??= cloneDocument(this.#template));
  }

  get modified(): boolean {
    return this.#copy !== null;
  }

  fork(): XmlPart {
    return new XmlPart(this.path, this.read);
  }
}

export class PartInfo extends XmlPart {
  readonly category: PartCategory;
  /** The part's <Override> in [Content_Types].xml (read-only; cloned for new parts). */
  readonly override: Element;
  /** Contains merge placeholders in the template. Parts without fields are never copied. */
  hasFields = false;

  constructor(path: string, template: Document, category: PartCategory, override: Element) {
    super(path, template);
    this.category = category;
    this.override = override;
  }

  override fork(): PartInfo {
    const p = new PartInfo(this.path, this.read, this.category, this.override);
    p.hasFields = this.hasFields;
    return p;
  }
}

export interface NewPart {
  readonly path: string;
  readonly doc: Document;
  readonly rels: ZipEntry | undefined;
}

/**
 * An opened .docx. Zip entries are kept compressed and shared between renders; XML parts are
 * parsed once and copied lazily. Writing re-compresses only the parts a render modified.
 */
export class DocxPackage {
  readonly #entries: ReadonlyMap<string, ZipEntry>;
  readonly contentTypes: XmlPart;
  readonly parts: readonly PartInfo[];

  private constructor(entries: ReadonlyMap<string, ZipEntry>, contentTypes: XmlPart, parts: readonly PartInfo[]) {
    this.#entries = entries;
    this.contentTypes = contentTypes;
    this.parts = parts;
  }

  static open(bytes: Uint8Array): DocxPackage {
    let list: ZipEntry[];
    try {
      list = readZip(bytes);
    } catch (error) {
      throw new Error(`Not a valid .docx (zip) file: ${(error as Error).message}`, { cause: error });
    }
    const entries = new Map(list.map((e) => [e.name, e]));
    const ct = entries.get(CONTENT_TYPES);
    if (!ct) throw new Error(`Not a valid .docx file: missing ${CONTENT_TYPES}`);
    const contentTypes = parseXml(inflate(ct), CONTENT_TYPES);

    const parts: PartInfo[] = [];
    for (const override of childrenNamed(contentTypes.documentElement!, NS.ct, 'Override')) {
      const category = CONTENT_TYPE_CATEGORIES[override.getAttribute('ContentType') ?? ''];
      if (!category) continue;
      const path = (override.getAttribute('PartName') ?? '').replace(/^\//, '');
      const entry = entries.get(path);
      if (!entry) continue; // dangling override: ignore like Word does
      parts.push(new PartInfo(path, parseXml(inflate(entry), path), category, override));
    }
    return new DocxPackage(entries, new XmlPart(CONTENT_TYPES, contentTypes), parts);
  }

  /** Call once after field parsing: records which parts contain fields. */
  indexFields(): void {
    for (const p of this.parts) p.hasFields = placeholders(p.read.documentElement!).length > 0;
  }

  fork(): DocxPackage {
    return new DocxPackage(this.#entries, this.contentTypes.fork(), this.parts.map((p) => p.fork()));
  }

  partsOf(...categories: PartCategory[]): PartInfo[] {
    return this.parts.filter((p) => categories.includes(p.category));
  }

  hasEntry(path: string): boolean {
    return this.#entries.has(path);
  }

  entry(path: string): ZipEntry | undefined {
    return this.#entries.get(path);
  }

  readXml(path: string): Document | null {
    const entry = this.#entries.get(path);
    return entry ? parseXml(inflate(entry), path) : null;
  }

  write(extraDocs: ReadonlyMap<string, Document>, newParts: readonly NewPart[]): Uint8Array {
    const modified = new Map<string, Document>(extraDocs);
    if (this.contentTypes.modified) modified.set(CONTENT_TYPES, this.contentTypes.read);
    for (const p of this.parts) if (p.modified) modified.set(p.path, p.read);

    const fresh = (path: string, data: Uint8Array): ZipEntry => {
      const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
      return makeEntry(path, data, STORED_EXTENSIONS.has(ext));
    };
    const out: ZipEntry[] = [];
    for (const [path, entry] of this.#entries) {
      const doc = modified.get(path);
      out.push(doc ? fresh(path, serializeXml(doc)) : entry);
    }
    for (const part of newParts) {
      out.push(fresh(part.path, serializeXml(part.doc)));
      if (part.rels) out.push({ ...part.rels, name: relsPathFor(part.path) });
    }
    return writeZip(out);
  }
}
