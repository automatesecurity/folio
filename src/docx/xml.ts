// A small XML reader that keeps source offsets. Everything the editor cannot
// model is carried as an exact slice of the original file, so a DOM parser
// (which would normalise namespaces, entities and whitespace) is not usable.

export interface XEl {
  kind: 'el';
  name: string;
  attrs: Record<string, string>;
  children: XNode[];
  start: number;
  end: number;
  innerStart: number;
  innerEnd: number;
}
export interface XText { kind: 'text'; text: string; start: number; end: number }
export type XNode = XEl | XText;

const NAME_END = /[\s/>=]/;

export function decodeEntities(s: string): string {
  if (!s.includes('&')) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e: string) => {
    if (e === 'amp') return '&';
    if (e === 'lt') return '<';
    if (e === 'gt') return '>';
    if (e === 'quot') return '"';
    if (e === 'apos') return "'";
    const code = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return code <= 0x10ffff ? String.fromCodePoint(code) : m;
  });
}

export function parseXml(src: string): XEl {
  let i = 0;
  const n = src.length;
  let root: XEl | null = null;
  const stack: XEl[] = [];

  const fail = (msg: string): never => { throw new Error(`Invalid XML at offset ${i}: ${msg}`); };

  while (i < n) {
    if (src[i] !== '<') {
      const j = src.indexOf('<', i);
      const end = j < 0 ? n : j;
      if (stack.length) {
        stack[stack.length - 1].children.push({ kind: 'text', text: decodeEntities(src.slice(i, end)), start: i, end });
      }
      i = end;
      continue;
    }
    if (src.startsWith('<?', i)) {
      const j = src.indexOf('?>', i);
      if (j < 0) fail('unterminated processing instruction');
      i = j + 2;
      continue;
    }
    if (src.startsWith('<!--', i)) {
      const j = src.indexOf('-->', i);
      if (j < 0) fail('unterminated comment');
      i = j + 3;
      continue;
    }
    if (src.startsWith('<![CDATA[', i)) {
      const j = src.indexOf(']]>', i);
      if (j < 0) fail('unterminated CDATA');
      if (stack.length) stack[stack.length - 1].children.push({ kind: 'text', text: src.slice(i + 9, j), start: i, end: j + 3 });
      i = j + 3;
      continue;
    }
    // DTDs are never legitimate in OOXML and are how entity-expansion attacks arrive.
    if (src.startsWith('<!', i)) fail('document type declarations are not allowed');
    if (src[i + 1] === '/') {
      const j = src.indexOf('>', i);
      if (j < 0) fail('unterminated end tag');
      const name = src.slice(i + 2, j).trim();
      const el = stack.pop();
      if (!el || el.name !== name) fail(`mismatched end tag ${name}`);
      el!.innerEnd = i;
      el!.end = j + 1;
      i = j + 1;
      continue;
    }
    const start = i;
    i++;
    let k = i;
    while (k < n && !NAME_END.test(src[k])) k++;
    const name = src.slice(i, k);
    if (!name) fail('empty tag name');
    i = k;
    const attrs: Record<string, string> = Object.create(null);
    let selfClosed = false;
    for (;;) {
      while (i < n && /\s/.test(src[i])) i++;
      if (i >= n) fail('unterminated tag');
      if (src[i] === '>') { i++; break; }
      if (src[i] === '/' && src[i + 1] === '>') { i += 2; selfClosed = true; break; }
      let e = i;
      while (e < n && !NAME_END.test(src[e])) e++;
      const an = src.slice(i, e);
      i = e;
      while (i < n && /\s/.test(src[i])) i++;
      if (src[i] !== '=') fail('expected =');
      i++;
      while (i < n && /\s/.test(src[i])) i++;
      const q = src[i];
      if (q !== '"' && q !== "'") fail('expected quoted attribute');
      const close = src.indexOf(q, i + 1);
      if (close < 0) fail('unterminated attribute');
      attrs[an] = decodeEntities(src.slice(i + 1, close));
      i = close + 1;
    }
    const el: XEl = { kind: 'el', name, attrs, children: [], start, end: i, innerStart: i, innerEnd: i };
    if (stack.length) stack[stack.length - 1].children.push(el);
    else if (root) fail('multiple root elements');
    else root = el;
    if (!selfClosed) stack.push(el);
  }
  if (stack.length) fail(`unclosed element ${stack[stack.length - 1].name}`);
  if (!root) fail('no root element');
  return root!;
}

const ILLEGAL = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g;
export const escText = (s: string) =>
  s.replace(ILLEGAL, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
export const escAttr = (s: string) => escText(s).replace(/"/g, '&quot;').replace(/\n/g, '&#10;').replace(/\t/g, '&#9;');

export const elements = (el: XEl): XEl[] => el.children.filter((c): c is XEl => c.kind === 'el');
export const local = (name: string) => name.slice(name.indexOf(':') + 1);
export const child = (el: XEl, localName: string): XEl | undefined =>
  elements(el).find((c) => local(c.name) === localName);
export const attr = (el: XEl | undefined, localName: string): string | undefined => {
  if (!el) return undefined;
  for (const k in el.attrs) if (local(k) === localName && !k.startsWith('xmlns')) return el.attrs[k];
  return undefined;
};
export function* walk(el: XEl): Generator<XEl> {
  yield el;
  for (const c of elements(el)) yield* walk(c);
}
/** All w:t style text under an element, in document order. */
export function textOf(el: XEl, source?: Set<string>): string {
  const names = source ?? new Set(['t', 'delText']);
  let out = '';
  for (const c of walk(el)) {
    const l = local(c.name);
    if (names.has(l)) for (const t of c.children) if (t.kind === 'text') out += t.text;
    if (l === 'tab') out += '\t';
  }
  return out;
}
export const rawAttrs = (el: XEl, skip?: (k: string) => boolean): string =>
  Object.keys(el.attrs).filter((k) => !skip?.(k)).map((k) => ` ${k}="${escAttr(el.attrs[k])}"`).join('');
