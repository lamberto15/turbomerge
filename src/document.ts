import { writeFile } from 'node:fs/promises';
import {
  NS, SECTPR_AFTER_TYPE, SEPARATORS, SETTINGS_AFTER_UPDATE_FIELDS, type Separator,
} from './constants.ts';
import { NextRecord, SkipRecord, type Field, type FillContext } from './fields.ts';
import type { MergeRecord, ResolvedOptions } from './options.ts';
import { relsPathFor, type DocxPackage, type NewPart, type PartInfo } from './package.ts';
import type { FieldRegistry } from './parse.ts';
import {
  child, childrenNamed, clone, cloneDocument, createW, descendants, insertOrdered, nextElement,
  placeholders, remove, removeChildren, setWAttr, type Document, type Element,
} from './xml.ts';

export interface WriteOptions {
  /** Text used for fields without data when `keepFields` is `'none'`. Default `''`. */
  emptyValue?: string;
}

interface HeaderTemplate {
  readonly part: PartInfo;
  readonly dir: string;
  readonly target: string;
  readonly stem: string;
}

const PART_NAME = /^(.*?)(\d*)\.xml$/;

/**
 * A single output document being built from a template. Use it for incremental work
 * (several `merge()` calls, `mergeRows()`, `mergeTemplates()`), then `toBuffer()` / `save()`.
 * `DocxTemplate.render*()` are one-call shortcuts over this class.
 */
export class MailMergeDocument {
  readonly #pkg: DocxPackage;
  readonly #registry: FieldRegistry;
  readonly #options: ResolvedOptions;
  readonly #ctx: FillContext;
  readonly #rels = new Map<string, Document | null>();
  readonly #newParts: NewPart[] = [];
  readonly #partCounters = new Map<string, number>();
  #rows: readonly MergeRecord[] | null = null;
  #index = 0;
  #replaceMissing = false;
  #docPrIds: Set<number> | null = null;
  #maxDocPrId = 0;
  readonly #maxRelId = new Map<Document, number>();
  #templatesMerged = false;

  /** @internal use `DocxTemplate.createDocument()` */
  constructor(pkg: DocxPackage, registry: FieldRegistry, options: ResolvedOptions) {
    this.#pkg = pkg;
    this.#registry = registry;
    this.#options = options;
    this.#ctx = {
      format: { options, locale: options.locale, timeZone: options.timeZone, field: '' },
      field: (key) => registry.get(key),
      instrText: (elements) => registry.instrText(elements),
      warn: (code, message, field) => options.onWarning(field === undefined ? { code, message } : { code, message, field }),
    };
  }

  // ---- public API ----------------------------------------------------------------------

  /** Names of the merge fields still present (not yet merged) in the body, headers, footers and notes. */
  getMergeFields(): Set<string> {
    const names = new Set<string>();
    for (const root of this.#roots(false, false)) {
      for (const el of placeholders(root)) {
        const name = el.getAttribute('name');
        if (name) names.add(name);
      }
    }
    return names;
  }

  /** Fills fields from one record. Array values fill table rows (the key names a field in the row). */
  merge(record: MergeRecord): this {
    for (const root of this.#roots(true)) this.#replace(root, record);
    return this;
  }

  /** Repeats the table row that contains field `anchor` once per record. */
  mergeRows(anchor: string, rows: readonly MergeRecord[]): this {
    for (const root of this.#roots(false)) this.#replaceTableRows(root, anchor, rows);
    return this;
  }

  /**
   * Repeats the whole document body once per record, joined by `separator`.
   * With a `*_section` separator, headers/footers containing fields are duplicated per record.
   * Supports NEXT, NEXTIF and SKIPIF.
   */
  mergeTemplates(records: Iterable<MergeRecord>, separator: Separator = 'page_break'): this {
    if (!SEPARATORS.includes(separator)) {
      throw new TypeError(`Invalid separator "${separator}". Expected one of: ${SEPARATORS.join(', ')}`);
    }
    if (this.#templatesMerged) throw new Error('mergeTemplates() can only be called once per document');
    const rows = Array.from(records);
    if (rows.length === 0) throw new RangeError('mergeTemplates() needs at least one record');
    const main = this.#pkg.partsOf('main')[0];
    if (!main) throw new Error('Document has no main part');
    this.#templatesMerged = true;

    const [sepType = '', sepClass] = separator.split('_');
    const body = child(main.doc.documentElement!, NS.w, 'body');
    if (!body) throw new Error('Document has no body');

    if (sepClass === 'section') {
      const firstSection = this.#firstSection(body);
      for (const t of childrenNamed(firstSection, NS.w, 'type')) firstSection.removeChild(t);
      const type = createW(main.doc, 'type');
      setWAttr(type, 'val', sepType);
      insertOrdered(firstSection, type, (el) => SECTPR_AFTER_TYPE.has(el.localName ?? ''));
    }
    const lastSection = child(body, NS.w, 'sectPr');
    if (!lastSection) throw new Error('Document body has no final section properties (w:sectPr)');
    remove(lastSection);
    const bodyTemplate = clone(body);
    removeChildren(body);

    const separatorEl = createW(main.doc, 'p');
    if (sepClass === 'section') {
      separatorEl.appendChild(createW(main.doc, 'pPr')).appendChild(clone(lastSection));
    } else {
      const run = createW(main.doc, 'r');
      const br = createW(main.doc, 'br');
      setWAttr(br, 'type', sepType);
      run.appendChild(br);
      separatorEl.appendChild(run);
    }

    const headers = this.#pkg
      .partsOf('header_footer')
      .filter((p) => p.hasFields)
      .map((part): HeaderTemplate => {
        const slash = part.path.lastIndexOf('/');
        const target = part.path.slice(slash + 1);
        return { part, dir: part.path.slice(0, slash + 1), target, stem: PART_NAME.exec(target)?.[1] ?? target };
      });
    const mainRels = this.#relations(main.path);
    this.#ensureDocPrIds();

    type Copy = { header: HeaderTemplate; doc: Document };
    let pending: Copy[] = [];
    let appended = 0;

    this.#rows = rows;
    this.#index = 0;
    let row: MergeRecord | undefined = rows[0];
    try {
      while (this.#index < rows.length) {
        const current = clone(bodyTemplate);
        this.#fixDocPrIds(current);
        const copies: Copy[] = headers.map((header) => ({ header, doc: cloneDocument(header.part.read) }));
        try {
          for (const c of copies) this.#replace(c.doc.documentElement!, row);
          this.#replace(current, row);
          if (appended > 0) {
            const sep = clone(separatorEl);
            if (sepClass === 'section') {
              this.#commitHeaders(pending, mainRels, sep);
            }
            body.appendChild(sep);
          }
          while (current.firstChild) body.appendChild(current.firstChild);
          pending = copies;
          appended += 1;
        } catch (error) {
          if (!(error instanceof SkipRecord)) throw error;
        }
        row = this.#nextRow();
      }
    } finally {
      this.#rows = null;
    }
    this.#commitHeaders(pending, mainRels, lastSection);
    body.appendChild(lastSection);
    return this;
  }

  /** Finishes the document (fills remaining fields per `keepFields`) and returns the .docx bytes. */
  toBuffer(options: WriteOptions = {}): Uint8Array {
    const hasUnmerged = this.#roots(false, false).some((root) => placeholders(root).length > 0);
    if (this.#options.keepFields === 'none') {
      const empty = options.emptyValue ?? '';
      this.merge(Object.fromEntries([...this.getMergeFields()].map((name) => [name, empty])));
    } else {
      this.#replaceMissing = true;
      try {
        this.merge({});
      } finally {
        this.#replaceMissing = false;
      }
    }
    // never let an internal placeholder reach the output
    for (const root of this.#roots(true)) for (const el of placeholders(root)) remove(el);
    this.#fixSettings(hasUnmerged);

    const extra = new Map<string, Document>();
    for (const [path, doc] of this.#rels) if (doc) extra.set(path, doc);
    return this.#pkg.write(extra, this.#newParts);
  }

  /** Writes the finished document to `path`. */
  async save(path: string | URL, options?: WriteOptions): Promise<void> {
    await writeFile(path, this.toBuffer(options));
  }

  // ---- merge engine --------------------------------------------------------------------

  /** Part roots to merge into. `write` gives private copies of the parts that have fields. */
  #roots(includeNewParts: boolean, write = true): Element[] {
    const parts = this.#pkg.partsOf('main', 'header_footer', 'notes');
    const roots = write
      ? parts.filter((p) => p.hasFields).map((p) => p.doc.documentElement!)
      : parts.map((p) => p.read.documentElement!);
    if (includeNewParts) for (const p of this.#newParts) roots.push(p.doc.documentElement!);
    return roots;
  }

  #replace(body: Element, record: MergeRecord | undefined): void {
    if (record) {
      for (const [key, value] of Object.entries(record)) {
        if (Array.isArray(value)) this.#replaceTableRows(body, key, value as MergeRecord[]);
      }
    }
    this.#replaceInBody(body, record);
  }

  #replaceInBody(body: Element, record: MergeRecord | undefined): void {
    const found = placeholders(body);
    if (found.length === 0) return;
    let row = record;
    for (const el of found) {
      const field = this.#registry.get(el.getAttribute('merge_key')!);
      try {
        field.reset();
        if (field.hasValue(this.#ctx, row)) {
          field.fill(this.#ctx, row);
          this.#replaceField(el, field);
        } else if (this.#replaceMissing) {
          this.#replaceField(el, field, true);
        }
      } catch (error) {
        if (!(error instanceof NextRecord)) throw error;
        this.#replaceField(el, field);
        if (this.#rows) row = this.#nextRow(); // outside mergeTemplates NEXT has nothing to advance
      }
    }
    if (this.#replaceMissing) this.#replaceInBody(body, row);
  }

  #nextRow(): MergeRecord | undefined {
    this.#index += 1;
    return this.#rows?.[this.#index];
  }

  #replaceField(el: Element, field: Field, forceKeep = false): void {
    const parent = el.parentNode;
    if (!parent) return;
    const keep = forceKeep || this.#options.keepFields === 'all';
    for (const node of field.elementsToReplace(keep)) parent.insertBefore(node, el);
    parent.removeChild(el);
  }

  #replaceTableRows(body: Element, anchor: string, rows: readonly MergeRecord[]): void {
    const anchors = placeholders(body).filter((el) => el.getAttribute('name') === anchor);
    for (const fieldEl of anchors) {
      // nearest enclosing table and the row (direct child of the table) holding the field
      let row: Element = fieldEl;
      let table = row.parentNode as Element | null;
      while (table && !(table.localName === 'tbl' && table.namespaceURI === NS.w)) {
        row = table;
        table = table.parentNode as Element | null;
      }
      if (!table) continue; // not in a table, or its row was already expanded

      if (rows.length === 0) {
        if (this.#options.removeEmptyTables) remove(table);
        continue;
      }
      this.#ensureDocPrIds();
      const template = row;
      let cursor: Element | null = template;
      if (!this.#options.tableRowsReplaceMode) {
        cursor = nextElement(template);
        table.removeChild(template);
      }
      rows.forEach((data, i) => {
        let ref: Element | null = cursor;
        if (this.#options.tableRowsReplaceMode) {
          if (!cursor) throw new RangeError(`Table row ${i} does not fit into the table (tableRowsReplaceMode)`);
          ref = nextElement(cursor);
          table.removeChild(cursor);
          cursor = ref;
        }
        const copy = clone(template);
        if (i > 0) this.#fixDocPrIds(copy); // row 0 takes the place of the removed template
        this.#replace(copy, data);
        table.insertBefore(copy, ref);
      });
    }
  }

  // ---- package plumbing ----------------------------------------------------------------

  #firstSection(body: Element): Element {
    for (const p of childrenNamed(body, NS.w, 'p')) {
      const pPr = child(p, NS.w, 'pPr');
      const sectPr = pPr && child(pPr, NS.w, 'sectPr');
      if (sectPr) return sectPr;
    }
    const sectPr = child(body, NS.w, 'sectPr');
    if (!sectPr) throw new Error('Document body has no section properties (w:sectPr)');
    return sectPr;
  }

  #relations(partPath: string): Document | null {
    const path = relsPathFor(partPath);
    if (!this.#rels.has(path)) {
      this.#rels.set(path, this.#pkg.readXml(path));
    }
    return this.#rels.get(path)!;
  }

  #nextRelId(rels: Document): string {
    let max = this.#maxRelId.get(rels);
    if (max === undefined) {
      max = 0;
      for (const r of childrenNamed(rels.documentElement!, NS.rr, 'Relationship')) {
        const m = /^rId(\d+)$/.exec(r.getAttribute('Id') ?? '');
        if (m) max = Math.max(max, Number(m[1]));
      }
    }
    max += 1;
    this.#maxRelId.set(rels, max);
    return `rId${max}`;
  }

  #newPartPath(header: HeaderTemplate): string {
    const key = header.dir + header.stem;
    let n = this.#partCounters.get(key) ?? 0;
    let path: string;
    do {
      n += 1;
      path = `${header.dir}${header.stem}${n}.xml`;
    } while (this.#pkg.hasEntry(path) || this.#newParts.some((p) => p.path === path));
    this.#partCounters.set(key, n);
    return path;
  }

  /** Adds per-record header/footer copies as new parts and points `sectPr` at them. */
  #commitHeaders(copies: readonly { header: HeaderTemplate; doc: Document }[], rels: Document | null, sectPr: Element): void {
    for (const { header, doc } of copies) {
      const path = this.#newPartPath(header);
      const override = clone(header.part.override);
      override.setAttribute('PartName', `/${path}`);
      this.#pkg.contentTypes.doc.documentElement!.appendChild(override);
      // copy the header's own relationships (images, hyperlinks) - otherwise those break
      this.#newParts.push({ path, doc, rels: this.#pkg.entry(relsPathFor(header.part.path)) });

      if (!rels) continue;
      const oldRel = childrenNamed(rels.documentElement!, NS.rr, 'Relationship')
        .find((r) => r.getAttribute('Target') === header.target);
      if (!oldRel) continue;
      const oldId = oldRel.getAttribute('Id');
      const newRel = clone(oldRel);
      const newId = this.#nextRelId(rels);
      newRel.setAttribute('Id', newId);
      newRel.setAttribute('Target', path.slice(header.dir.length));
      rels.documentElement!.appendChild(newRel);
      for (const el of Array.from(sectPr.getElementsByTagName('*'))) {
        if (el.getAttributeNS(NS.r, 'id') === oldId) el.setAttributeNS(NS.r, 'r:id', newId);
      }
    }
  }

  #ensureDocPrIds(): void {
    if (this.#docPrIds) return;
    this.#docPrIds = new Set();
    for (const root of this.#roots(true, false)) {
      for (const el of descendants(root, NS.wp, 'docPr')) {
        const id = Number(el.getAttribute('id'));
        if (Number.isInteger(id)) {
          this.#docPrIds.add(id);
          this.#maxDocPrId = Math.max(this.#maxDocPrId, id);
        }
      }
    }
  }

  /** Drawing ids (wp:docPr/@id) must be unique; duplicated content gets fresh ids. */
  #fixDocPrIds(root: Element): void {
    const ids = this.#docPrIds!;
    for (const el of descendants(root, NS.wp, 'docPr')) {
      const id = Number(el.getAttribute('id'));
      if (!Number.isInteger(id)) continue;
      if (ids.has(id)) {
        const fresh = ++this.#maxDocPrId;
        ids.add(fresh);
        el.setAttribute('id', String(fresh));
      } else {
        ids.add(id);
        this.#maxDocPrId = Math.max(this.#maxDocPrId, id);
      }
    }
  }

  #fixSettings(hasUnmerged: boolean): void {
    const part = this.#pkg.partsOf('settings')[0];
    if (!part) return;
    // a leftover mail-merge data source makes Word prompt (and fail) on open
    if (!hasUnmerged && child(part.read.documentElement!, NS.w, 'mailMerge')) {
      const settings = part.doc.documentElement!;
      for (const mm of childrenNamed(settings, NS.w, 'mailMerge')) settings.removeChild(mm);
    }

    const mode = this.#options.autoUpdateFieldsOnOpen;
    if (mode === 'always' || (mode === 'auto' && this.#registry.hasNestedFields)) {
      const settings = part.doc.documentElement!;
      let update = child(settings, NS.w, 'updateFields');
      if (!update) {
        update = createW(settings.ownerDocument!, 'updateFields');
        insertOrdered(settings, update, (el) => SETTINGS_AFTER_UPDATE_FIELDS.has(el.localName ?? ''));
      }
      setWAttr(update, 'val', 'true');
    }
  }
}
