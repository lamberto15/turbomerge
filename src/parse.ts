import { NS } from './constants.ts';
import { FIELD_CLASSES, type Field, type FieldType } from './fields.ts';
import { tokenizeInstruction } from './tokenize.ts';
import {
  child,
  childrenNamed,
  clone,
  descendants,
  type Element,
  isPlaceholder,
  nextElement,
  textOf,
  wAttr,
} from './xml.ts';

const BASE_TYPES: ReadonlySet<string> = new Set(['MERGEFIELD', 'NEXT', 'NEXTIF', 'SKIPIF']);

/** Parse-time catalogue of all fields of a template, keyed by placeholder key. */
export class FieldRegistry {
  readonly #fields = new Map<string, Field>();
  readonly #supported: ReadonlySet<string>;
  readonly #warn: (message: string) => void;
  #nextKey = 0;
  hasNestedFields = false;

  constructor(mergeIfFields: boolean, warn: (message: string) => void) {
    this.#supported = mergeIfFields ? new Set([...BASE_TYPES, 'IF']) : BASE_TYPES;
    this.#warn = warn;
  }

  get(key: string): Field {
    const field = this.#fields.get(key);
    if (!field) throw new Error(`Unknown merge field key ${key}`);
    return field;
  }

  all(): IterableIterator<Field> {
    return this.#fields.values();
  }

  markNested(key: string): void {
    this.hasNestedFields = true;
    this.get(key).nested = true;
  }

  /** Text of a field instruction; nested placeholders appear as `{key}` (or their own instruction when recursive). */
  instrText(elements: readonly Element[], recursive = false): string {
    let text = '';
    let currentParent: unknown;
    let first = true;
    for (const el of elements) {
      const parent = el.parentNode;
      if (first) {
        currentParent = parent;
        first = false;
      } else if (parent !== currentParent) {
        currentParent = parent;
        text += '\n';
      }
      for (const it of childrenNamed(el, NS.w, 'instrText')) text += textOf(it);
      if (isPlaceholder(el)) {
        const key = el.getAttribute('merge_key')!;
        text += recursive ? this.get(key).instr : `{${key}}`;
      }
    }
    return text;
  }

  make(init: {
    parent: Element;
    part: string;
    nested?: boolean;
    simple?: boolean;
    instr?: string | null;
    allElements: Element[];
    instrElements: Element[];
    showElements: Element[];
  }): Field | null {
    const instr = init.instr ?? this.instrText(init.instrElements);
    const tokens = tokenizeInstruction(instr, this.#warn);
    const type = (tokens[0] ?? '').toUpperCase();
    if (!this.#supported.has(type)) return null;
    const key = `field_${this.#nextKey++}`;
    const nested = init.nested ?? false;
    const Cls = FIELD_CLASSES[type as FieldType];
    const field = new Cls({
      key,
      type: type as FieldType,
      instr,
      tokens,
      parent: init.parent,
      nested,
      part: init.part,
      simple: init.simple ?? false,
      allElements: init.allElements,
      instrElements: init.instrElements,
      showElements: init.showElements,
    });
    if (nested) this.hasNestedFields = true;
    this.#fields.set(key, field);
    return field;
  }
}

interface Step {
  element: Element;
  charType: string | null;
}

/** Next run of a complex field, continuing into following paragraphs when the field spans them. */
function nextStep(current: Element): Step | null {
  let next = nextElement(current);
  let paragraph: Element | null = current.parentNode as Element | null;
  while (next === null) {
    paragraph = paragraph ? nextElement(paragraph) : null;
    if (paragraph === null) return null;
    next = child(paragraph, NS.w, 'r');
  }
  const fldChar = child(next, NS.w, 'fldChar');
  return { element: next, charType: fldChar ? wAttr(fldChar, 'fldCharType') : null };
}

/** Replaces every supported field of `root` with a placeholder and registers it. */
export function parsePart(root: Element, partName: string, registry: FieldRegistry): void {
  // simple fields: <w:fldSimple w:instr="MERGEFIELD name"><w:r>…</w:r></w:fldSimple>
  for (const fld of descendants(root, NS.w, 'fldSimple')) {
    const firstRun = child(fld, NS.w, 'r');
    if (!firstRun || !fld.parentNode) continue;
    const run = clone(firstRun);
    const field = registry.make({
      parent: fld.parentNode as Element,
      part: partName,
      simple: true,
      instr: wAttr(fld, 'instr') ?? '',
      allElements: [fld],
      instrElements: [run],
      showElements: [run],
    });
    field?.insertIntoTree();
  }

  // complex fields: runs with <w:fldChar begin/> … instrText … <separate/> … result … <end/>
  const begins = descendants(root, NS.w, 'fldChar')
    .filter((fc) => wAttr(fc, 'fldCharType') === 'begin')
    .map((fc) => fc.parentNode as Element)
    .filter((run, i, arr) => run && arr.indexOf(run) === i);

  const pull = (nested: boolean): [Field | null, Element] => {
    let current = begins.shift()!;
    const parent = current.parentNode as Element;
    const allElements: Element[] = [current];
    const instrElements: Element[] = [];
    const showElements: Element[] = [];
    let list = instrElements;
    let charType: string | null = null;

    while (charType !== 'end') {
      const step = nextStep(current);
      if (!step) throw new Error(`Field begin without end near: ${registry.instrText(instrElements, true)}`);
      let next = step.element;
      charType = step.charType;
      if (charType === 'begin') {
        if (begins[0] !== next) throw new Error('Malformed nested field structure');
        const [sub, end] = pull(true);
        next = sub ? sub.insertIntoTree() : end;
      } else if (charType === 'separate') {
        list = showElements;
      } else if (isPlaceholder(next)) {
        registry.markNested(next.getAttribute('merge_key')!); // simple field nested in a complex one
      }
      if (charType !== 'end' && charType !== 'separate') list.push(next);
      allElements.push(next);
      current = next;
    }
    const field = registry.make({ parent, part: partName, nested, allElements, instrElements, showElements });
    return [field, current];
  };

  while (begins.length > 0) {
    const [field] = pull(false);
    field?.insertIntoTree();
  }
}
