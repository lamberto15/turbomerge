import { DOMParser, XMLSerializer, type Document, type Element, type Node } from '@xmldom/xmldom';
import { NS, PLACEHOLDER } from './constants.ts';

export type { Document, Element, Node };

const ELEMENT_NODE = 1;
const decoder = new TextDecoder('utf-8');
const encoder = new TextEncoder();
const serializer = new XMLSerializer();

export function parseXml(input: Uint8Array | string, name = 'xml'): Document {
  const text = typeof input === 'string' ? input : decoder.decode(input);
  const parser = new DOMParser({
    onError(level, message) {
      if (level === 'fatalError') throw new Error(`Invalid XML in ${name}: ${message}`);
    },
  });
  return parser.parseFromString(text, 'text/xml');
}

export function serializeXml(doc: Document): Uint8Array {
  let text = serializer.serializeToString(doc);
  if (!text.startsWith('<?xml')) {
    text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${text}`;
  }
  return encoder.encode(text);
}

export function isElement(node: Node | null | undefined): node is Element {
  return node != null && node.nodeType === ELEMENT_NODE;
}

export function is(node: Node | null | undefined, ns: string | null, localName: string): node is Element {
  return isElement(node) && node.localName === localName && (node.namespaceURI ?? null) === ns;
}

export function isPlaceholder(node: Node | null | undefined): node is Element {
  return is(node, null, PLACEHOLDER);
}

export function children(el: Element): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (isElement(n)) out.push(n);
  return out;
}

export function child(el: Element, ns: string, localName: string): Element | null {
  for (let n = el.firstChild; n; n = n.nextSibling) if (is(n, ns, localName)) return n;
  return null;
}

export function childrenNamed(el: Element, ns: string, localName: string): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) if (is(n, ns, localName)) out.push(n);
  return out;
}

export function nextElement(node: Node): Element | null {
  for (let n = node.nextSibling; n; n = n.nextSibling) if (isElement(n)) return n;
  return null;
}

/** Snapshot of all descendants (document order, excluding `root` itself). */
export function descendants(root: Element, ns: string, localName: string): Element[] {
  return Array.from(root.getElementsByTagNameNS(ns, localName));
}

export function placeholders(root: Element): Element[] {
  return Array.from(root.getElementsByTagName(PLACEHOLDER)).filter((el) => el.namespaceURI == null);
}

export function wAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(NS.w, name) || el.getAttribute(`w:${name}`) || null;
}

export function setWAttr(el: Element, name: string, value: string): void {
  el.setAttributeNS(NS.w, `w:${name}`, value);
}

export function createW(doc: Document, localName: string): Element {
  return doc.createElementNS(NS.w, `w:${localName}`);
}

export function ownerDoc(node: Node): Document {
  const doc = node.ownerDocument;
  if (!doc) throw new Error('Node has no owner document');
  return doc;
}

function deepCopy(node: Node, doc: Document): Node {
  if (node.nodeType === ELEMENT_NODE) {
    const src = node as Element;
    const el = doc.createElementNS(src.namespaceURI, src.nodeName);
    const attrs = src.attributes;
    for (let i = 0; i < attrs.length; i++) {
      const a = attrs[i]!;
      el.setAttributeNS(a.namespaceURI, a.nodeName, a.value);
    }
    for (let c = src.firstChild; c; c = c.nextSibling) el.appendChild(deepCopy(c, doc));
    return el;
  }
  if (node.nodeType === 3) return doc.createTextNode((node as unknown as { data: string }).data);
  return node.cloneNode(true);
}

/**
 * Deep copy of an element. ~8x faster than xmldom's generic `cloneNode(true)`, which dominates
 * render time when a template is rendered many times.
 */
export function clone<T extends Element>(node: T): T {
  return deepCopy(node, ownerDoc(node)) as T;
}

/** Deep copy of a whole document (keeps the XML declaration). */
export function cloneDocument(doc: Document): Document {
  const copy = doc.cloneNode(false) as Document;
  for (let c = doc.firstChild; c; c = c.nextSibling) copy.appendChild(deepCopy(c, copy));
  return copy;
}

export function remove(node: Node): void {
  node.parentNode?.removeChild(node);
}

export function removeChildren(el: Element): void {
  while (el.firstChild) el.removeChild(el.firstChild);
}

/** Insert `node` as a child of `parent` before the first child element accepted by `isAfter`. */
export function insertOrdered(parent: Element, node: Element, isAfter: (el: Element) => boolean): void {
  const ref = children(parent).find(isAfter) ?? null;
  parent.insertBefore(node, ref);
}

export function textOf(el: Element): string {
  return el.textContent ?? '';
}
