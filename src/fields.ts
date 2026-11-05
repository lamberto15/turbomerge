import { NS, PLACEHOLDER } from './constants.ts';
import { applySwitches, isNumeric, toNumber, toText, type FormatContext } from './format.ts';
import type { MergeRecord } from './options.ts';
import {
  child, clone, childrenNamed, createW, isPlaceholder, ownerDoc, remove, setWAttr,
  type Element,
} from './xml.ts';

/** Thrown by NEXT / NEXTIF: continue the current page with the next record. */
export class NextRecord extends Error {}
/** Thrown by SKIPIF: drop the current record entirely. */
export class SkipRecord extends Error {}

export type FieldType = 'MERGEFIELD' | 'NEXT' | 'NEXTIF' | 'SKIPIF' | 'IF';

const MISSING = Symbol('missing');

/** Looks up `name` in a record: exact key first, then a dot path through nested objects. */
export function lookup(record: MergeRecord | null | undefined, name: string): unknown {
  if (record == null) return MISSING;
  if (Object.hasOwn(record, name)) return record[name];
  if (!name.includes('.')) return MISSING;
  let current: unknown = record;
  for (const segment of name.split('.')) {
    if (current !== null && typeof current === 'object' && Object.hasOwn(current, segment)) {
      current = (current as Record<string, unknown>)[segment];
    } else {
      return MISSING;
    }
  }
  return current;
}

export function hasKey(record: MergeRecord | null | undefined, name: string): boolean {
  return lookup(record, name) !== MISSING;
}

/** What a field needs from the render session while filling. */
export interface FillContext {
  readonly format: FormatContext;
  field(key: string): Field;
  instrText(elements: readonly Element[]): string;
  warn(code: 'INVALID_FORMAT' | 'INVALID_CONDITION', message: string, field?: string): void;
}

export interface FieldInit {
  key: string;
  type: FieldType;
  instr: string;
  tokens: string[];
  parent: Element;
  nested: boolean;
  part: string;
  simple: boolean;
  allElements: Element[];
  instrElements: Element[];
  showElements: Element[];
}

/**
 * A Word field found in the template. The original XML nodes are detached from the tree and
 * kept here; the tree gets a `<MergeField merge_key=…>` placeholder. Rendering never mutates the
 * stored nodes (it clones them), so one parsed template can be rendered any number of times.
 */
export class Field {
  readonly key: string;
  readonly type: FieldType;
  readonly instr: string;
  readonly tokens: readonly string[];
  readonly name: string;
  readonly part: string;
  readonly simple: boolean;
  nested: boolean;
  protected readonly parent: Element;
  protected readonly allElements: Element[];
  protected readonly instrElements: Element[];
  protected readonly showElements: Element[];
  protected readonly nestedKeys: string[];
  /** Per-fill state. */
  protected currentTokens: readonly string[];
  protected filled: Element[] = [];

  constructor(init: FieldInit) {
    this.key = init.key;
    this.type = init.type;
    this.instr = init.instr;
    this.tokens = init.tokens;
    this.currentTokens = init.tokens;
    this.parent = init.parent;
    this.nested = init.nested;
    this.part = init.part;
    this.simple = init.simple;
    this.allElements = init.allElements;
    this.instrElements = init.instrElements;
    this.showElements = init.showElements;
    this.nestedKeys = init.instrElements.filter(isPlaceholder).map((el) => el.getAttribute('merge_key')!);
    this.name = init.type === 'MERGEFIELD' ? (init.tokens[1] ?? '') : '';
  }

  get nestedFieldKeys(): readonly string[] {
    return this.nestedKeys;
  }

  reset(): void {
    this.filled = [];
    this.currentTokens = this.tokens;
  }

  hasValue(ctx: FillContext, record: MergeRecord | null | undefined): boolean {
    for (const key of this.nestedKeys) if (!ctx.field(key).hasValue(ctx, record)) return false;
    return !(this.name && !hasKey(record, this.name));
  }

  /** Names of fields (this one or nested) that the record does not provide. */
  missingNames(ctx: FillContext, record: MergeRecord | null | undefined, out: Set<string>): void {
    for (const key of this.nestedKeys) ctx.field(key).missingNames(ctx, record, out);
    if (this.name && !hasKey(record, this.name)) out.add(this.name);
  }

  protected fillNested(ctx: FillContext, record: MergeRecord | null | undefined): void {
    if (this.nestedKeys.length === 0) return;
    const values = new Map<string, string>();
    for (const key of this.nestedKeys) {
      const nested = ctx.field(key);
      nested.fill(ctx, record);
      values.set(key, ctx.instrText(nested.elementsToReplace(false)));
    }
    this.currentTokens = this.tokens.map((t) =>
      t.includes('{') ? t.replace(/\{(field_\d+)\}/g, (m, k: string) => values.get(k) ?? m) : t,
    );
  }

  /** Computes the replacement elements for this record. May throw NextRecord / SkipRecord. */
  fill(ctx: FillContext, record: MergeRecord | null | undefined): void {
    this.fillNested(ctx, record);
    this.filled = [];
    const raw = lookup(record, this.name);
    let value: unknown = raw === MISSING ? `«${this.name}»` : raw;
    const fctx: FormatContext = { ...ctx.format, field: this.name };
    try {
      value = applySwitches(value, this.currentTokens.slice(2), fctx);
    } catch (error) {
      ctx.warn('INVALID_FORMAT', `Invalid formatting for field <${this.instr.trim()}>: ${(error as Error).message}`, this.name);
    }
    this.fillValue(toText(value, fctx));
  }

  protected fillValue(text: string): void {
    const base = this.instrElements[0];
    const doc = ownerDoc(this.parent);
    const run = base ? clone(base) : createW(doc, 'r');
    for (const name of ['instrText', 't', 'fldChar', 'cr', 'br', 'tab']) {
      for (const c of childrenNamed(run, NS.w, name)) run.removeChild(c);
    }
    const lines = text.replace(/\r/g, '').split('\n');
    lines.forEach((line, i) => {
      if (i > 0) run.appendChild(createW(doc, 'cr'));
      const t = createW(doc, this.nested ? 'instrText' : 't');
      t.setAttributeNS(NS.xml, 'xml:space', 'preserve');
      t.appendChild(doc.createTextNode(line));
      run.appendChild(t);
    });
    this.filled.push(run);
  }

  /**
   * - keep=false: just the result runs
   * - keep=true, nothing filled: the original field, untouched
   * - keep=true, filled: the original field with its result replaced
   */
  elementsToReplace(keep: boolean): Element[] {
    if (!keep) return this.filled;
    if (this.filled.length === 0) return this.allElements.map(clone);
    return this.withFilledResult();
  }

  protected withFilledResult(): Element[] {
    if (this.simple) {
      const fld = clone(this.allElements[0]!);
      while (fld.firstChild) fld.removeChild(fld.firstChild);
      for (const el of this.filled) fld.appendChild(el);
      return [fld];
    }
    const all = this.allElements.map(clone);
    if (this.showElements.length === 0) {
      const separate = clone(this.allElements.at(-1)!);
      const fldChar = child(separate, NS.w, 'fldChar');
      if (fldChar) setWAttr(fldChar, 'fldCharType', 'separate');
      all.splice(all.length - 1, 0, separate, ...this.filled);
    } else {
      const index = this.allElements.indexOf(this.showElements[0]!);
      all.splice(index, this.showElements.length, ...this.filled);
    }
    return all;
  }

  /** Detaches the field's nodes from the tree and puts a placeholder where it started. */
  insertIntoTree(): Element {
    const toDrop: Element[] = [];
    let currentParent = this.parent;
    for (const el of this.allElements.slice(1)) {
      const parent = el.parentNode as Element | null;
      remove(el);
      if (parent && parent !== currentParent) {
        toDrop.push(parent);
        currentParent = parent;
      }
    }
    // a field spanning paragraphs: the continuation paragraphs are merged away
    for (const p of toDrop) remove(p);
    const placeholder = ownerDoc(this.parent).createElementNS(null, PLACEHOLDER);
    placeholder.setAttribute('merge_key', this.key);
    placeholder.setAttribute('name', this.name);
    this.parent.replaceChild(placeholder, this.allElements[0]!);
    return placeholder;
  }
}

export class NextField extends Field {
  override fill(): void {
    throw new NextRecord();
  }
}

const numericOrString = (a: string, b: string): [number, number] | [string, string] =>
  isNumeric(a) && isNumeric(b) ? [toNumber(a), toNumber(b)] : [a, b];

function wildcardEquals(a: string, b: string): boolean {
  if (isNumeric(a) && isNumeric(b)) return toNumber(a) === toNumber(b);
  const pattern = b.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
  return new RegExp(`^(?:${pattern})$`, 'is').test(a);
}

const OPERATORS: Record<string, (a: string, b: string) => boolean> = {
  '=': wildcardEquals,
  '<>': (a, b) => !wildcardEquals(a, b),
  '!=': (a, b) => !wildcardEquals(a, b),
  '<': (a, b) => { const [x, y] = numericOrString(a, b); return x < y; },
  '<=': (a, b) => { const [x, y] = numericOrString(a, b); return x <= y; },
  '>': (a, b) => { const [x, y] = numericOrString(a, b); return x > y; },
  '>=': (a, b) => { const [x, y] = numericOrString(a, b); return x >= y; },
};

export abstract class ConditionalField extends Field {
  protected evaluate(ctx: FillContext): boolean {
    const [left, op, right] = this.currentTokens.slice(1, 4);
    if (left === undefined || op === undefined || right === undefined) {
      ctx.warn('INVALID_CONDITION', `Incomplete condition in <${this.instr.trim()}>`);
      return false;
    }
    const fn = OPERATORS[op];
    if (!fn) {
      ctx.warn('INVALID_CONDITION', `Invalid operator ${op} in <${this.instr.trim()}>`);
      return false;
    }
    return fn(left, right);
  }

  override fill(ctx: FillContext, record: MergeRecord | null | undefined): void {
    this.fillNested(ctx, record);
    this.filled = [];
    if (this.evaluate(ctx)) this.onTrue();
    else this.onFalse();
  }

  protected onTrue(): void {}
  protected onFalse(): void {}
}

export class NextIfField extends ConditionalField {
  protected override onTrue(): void {
    throw new NextRecord();
  }
}

export class SkipIfField extends ConditionalField {
  protected override onTrue(): void {
    throw new SkipRecord();
  }
}

export class IfField extends ConditionalField {
  protected override onTrue(): void {
    this.fillValue(this.currentTokens[4] ?? '');
  }
  protected override onFalse(): void {
    this.fillValue(this.currentTokens[5] ?? '');
  }
}

export const FIELD_CLASSES: Record<FieldType, new (init: FieldInit) => Field> = {
  MERGEFIELD: Field,
  NEXT: NextField,
  NEXTIF: NextIfField,
  SKIPIF: SkipIfField,
  IF: IfField,
};
