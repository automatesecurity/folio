import { Attrs, DOMOutputSpec, DOMSerializer, Fragment, Mark, Node as PMNode, Schema } from 'prosemirror-model';
import { cssIdent, fontStack, jcToCss, parseParaFragment, paraCss, tw } from '../docx/styles';
import { session } from './session';

const safeHex = (v: unknown) => (typeof v === 'string' && /^[0-9a-fA-F]{6}$/.test(v) ? `#${v}` : null);
export const HIGHLIGHTS: Record<string, string> = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#00008b', darkCyan: '#008b8b',
  darkGreen: '#006400', darkMagenta: '#8b008b', darkRed: '#8b0000', darkYellow: '#808000', darkGray: '#a9a9a9', lightGray: '#d3d3d3', black: '#000000', white: '#ffffff',
};

function paragraphDom(node: PMNode): DOMOutputSpec {
  const a = node.attrs;
  const attrs: Record<string, string> = {};
  if (a.pStyle) attrs.class = `ps-${cssIdent(a.pStyle)}`;
  const props = (a.pExtra as string[]).length ? parseParaFragment((a.pExtra as string[]).join('')) : {};
  const styles: string[] = [paraCss(props)];
  const list = a.numId && a.numId !== '0' ? session.numbering?.level(a.numId, a.ilvl ?? 0) : undefined;
  if (list) {
    if (list.left !== undefined) styles.push(`margin-left:${tw(list.left)}px`);
    if (list.first !== undefined) { styles.push(`text-indent:${tw(list.first)}px`); styles.push(`--hang:${tw(-list.first)}px`); }
  }
  const ind = a.ind as Record<string, string> | null;
  if (ind) {
    const left = ind.left ?? ind.start, right = ind.right ?? ind.end;
    if (left !== undefined) styles.push(`margin-left:${tw(+left)}px`);
    if (right !== undefined) styles.push(`margin-right:${tw(+right)}px`);
    if (ind.firstLine !== undefined) styles.push(`text-indent:${tw(+ind.firstLine)}px`);
    if (ind.hanging !== undefined) { styles.push(`text-indent:${tw(-ind.hanging)}px`); styles.push(`--hang:${tw(+ind.hanging)}px`); }
  }
  if (a.jc) styles.push(`text-align:${jcToCss(a.jc)}`);
  const style = styles.filter(Boolean).join(';');
  if (style) attrs.style = style;
  return ['p', attrs, 0];
}

const STYLE_TO_JC: Record<string, string> = { center: 'center', right: 'right', justify: 'both' };

/** Formatting that survives copy and paste inside one document; see clipboardSerializer. */
function carried(el: HTMLElement): Record<string, unknown> | null {
  const raw = el.dataset.folio;
  if (!raw || Number(el.dataset.session) !== session.id) return null;
  try { return JSON.parse(raw); } catch { return null; }
}

function headingStyle(level: number): string | null {
  const sheet = session.styles;
  if (!sheet) return null;
  for (const s of sheet.styles.values()) if (s.type === 'paragraph' && s.name.toLowerCase() === `heading ${level}`) return s.id;
  return null;
}
function listDepth(el: HTMLElement) { let d = -1; for (let p: HTMLElement | null = el; p; p = p.parentElement) if (p.tagName === 'UL' || p.tagName === 'OL') d++; return Math.max(0, Math.min(8, d)); }

export const schema = new Schema({
  nodes: {
    doc: { content: 'block+', attrs: { pgMar: { default: null }, titlePg: { default: false } } },
    paragraph: {
      group: 'block', content: 'inline*',
      attrs: {
        srcId: { default: null }, pStyle: { default: null }, jc: { default: null }, numId: { default: null }, ilvl: { default: null },
        ind: { default: null }, pExtra: { default: [] }, rawAttrs: { default: '' }, pmark: { default: null },
      },
      parseDOM: [
        { tag: 'p[data-folio]', getAttrs: (el) => carried(el as HTMLElement) ?? {} },
        { tag: 'p', getAttrs: (el) => { const ta = (el as HTMLElement).style.textAlign; return { jc: STYLE_TO_JC[ta] ?? null }; } },
        ...[1, 2, 3, 4, 5, 6].map((l) => ({ tag: `h${l}`, getAttrs: () => ({ pStyle: headingStyle(l) }) })),
        { tag: 'li', getAttrs: (el) => {
          const e = el as HTMLElement;
          const kind = e.parentElement?.tagName === 'OL' ? 'ordered' : 'bullet';
          return { numId: session.listFor(kind, e.parentElement), ilvl: listDepth(e) };
        } },
      ],
      toDOM: paragraphDom,
    },
    text: { group: 'inline' },
    hard_break: {
      inline: true, group: 'inline', selectable: false, attrs: { xml: { default: null } },
      parseDOM: [{ tag: 'br' }], toDOM: () => ['br'],
    },
    page_break: {
      inline: true, group: 'inline', atom: true, attrs: { xml: { default: null } },
      parseDOM: [{ tag: 'span.pm-page-break' }],
      toDOM: () => ['span', { class: 'pm-page-break', contenteditable: 'false' }, 'Page break'],
    },
    opaque_inline: {
      inline: true, group: 'inline', atom: true, marks: '',
      attrs: { kind: {}, xml: {}, label: { default: '' }, img: { default: null }, w: { default: 0 }, h: { default: 0 }, session: { default: 0 }, instr: { default: null } },
      parseDOM: [{
        tag: 'span[data-opaque]',
        getAttrs: (el) => {
          const e = el as HTMLElement;
          // Raw XML can carry relationship ids that only exist in its own document.
          if (Number(e.dataset.session) !== session.id) return false;
          const v = session.pasteVault.get(e.dataset.opaque!);
          return v ?? false;
        },
      }],
      toDOM: (node) => {
        const a = node.attrs;
        const key = session.stash(a);
        const base: Record<string, string> = { 'data-opaque': key, 'data-session': String(session.id), contenteditable: 'false' };
        if (a.instr) base['data-field'] = a.instr;
        if (a.kind === 'drawing') {
          const url = a.img ? session.assets.get(a.img) : null;
          const w = a.w ? Math.max(8, a.w) : 160, h = a.h ? Math.max(8, a.h) : 100;
          if (url) return ['span', { ...base, class: 'ob ob-img' }, ['img', { src: url, style: `width:${w}px;height:${h}px`, alt: a.label || 'Image' }]];
          return ['span', { ...base, class: 'ob ob-box', style: `width:${w}px;height:${h}px` }, a.label || 'Graphic'];
        }
        const hidden = ['bookmark', 'range', 'fieldPart', 'comment'].includes(a.kind);
        return ['span', { ...base, class: `ob ob-${a.kind}${hidden ? ' ob-hidden' : ''}`, title: a.title ?? '' }, hidden ? '\u200b' : a.label || ''];
      },
    },
    opaque_block: {
      group: 'block', atom: true, selectable: true,
      attrs: { srcId: { default: null }, kind: {}, xml: {}, preview: { default: [] }, session: { default: 0 } },
      parseDOM: [{ tag: 'div[data-folio]', getAttrs: (el) => carried(el as HTMLElement) ?? false }],
      toDOM: (node) => {
        const preview = node.attrs.preview as { text: string; cls: string | null }[];
        return ['div', { class: `ob-block ob-block-${node.attrs.kind}`, contenteditable: 'false' },
          ...(preview.length ? preview : [{ text: '\u00a0', cls: null }]).map((p) => ['p', p.cls ? { class: `ps-${cssIdent(p.cls)}` } : {}, p.text || '\u00a0'] as DOMOutputSpec)];
      },
    },
    table: {
      group: 'block', content: 'table_row+', tableRole: 'table', isolating: true,
      attrs: { srcId: { default: null }, tblPr: { default: '' }, rawAttrs: { default: '' }, gridRaw: { default: '' } },
      parseDOM: [{ tag: 'table', getAttrs: (el) => carried(el as HTMLElement) ?? {} }],
      toDOM: (node): DOMOutputSpec => {
        const first = node.firstChild;
        const cols: DOMOutputSpec[] = [];
        first?.forEach((c) => { const w = (c.attrs.colwidth as number[] | null) ?? []; for (let i = 0; i < c.attrs.colspan; i++) cols.push(['col', w[i] ? { style: `width:${tw(w[i])}px` } : {}]); });
        let total = 0;
        first?.forEach((c) => { for (const w of (c.attrs.colwidth as number[] | null) ?? []) total += w; });
        return ['table', { class: tableClass(node.attrs.tblPr), style: total ? `width:${tw(total)}px` : 'width:100%' }, ['colgroup', ...cols], ['tbody', 0]];
      },
    },
    table_row: {
      content: 'table_cell+', tableRole: 'row',
      attrs: { srcId: { default: null }, trPr: { default: '' }, rawAttrs: { default: '' }, rowChange: { default: null } },
      parseDOM: [{ tag: 'tr', getAttrs: (el) => carried(el as HTMLElement) ?? {} }], toDOM: () => ['tr', 0],
    },
    table_cell: {
      content: 'block+', tableRole: 'cell', isolating: true,
      attrs: {
        srcId: { default: null }, colspan: { default: 1 }, rowspan: { default: 1 }, colwidth: { default: null },
        tcPr: { default: [] }, rawAttrs: { default: '' },
      },
      parseDOM: [{ tag: 'td' }, { tag: 'th' }].map((r) => ({
        ...r, getAttrs: (el: HTMLElement | string) => {
          const e = el as HTMLElement;
          const kept = carried(e);
          if (kept) return kept;
          return { colspan: Number(e.getAttribute('colspan') ?? 1) || 1, rowspan: Number(e.getAttribute('rowspan') ?? 1) || 1 };
        },
      })),
      toDOM: (node): DOMOutputSpec => {
        const a: Record<string, string> = {};
        if (node.attrs.colspan > 1) a.colspan = String(node.attrs.colspan);
        if (node.attrs.rowspan > 1) a.rowspan = String(node.attrs.rowspan);
        const fill = /<[\w:]*shd\b[^>]*\bw?:?fill="([0-9a-fA-F]{6})"/.exec((node.attrs.tcPr as string[]).join(''));
        if (fill) a.style = `background:#${fill[1]}`;
        return ['td', a, 0];
      },
    },
  },
  marks: {
    tracked_ins: {
      inclusive: false, attrs: { author: { default: '' }, date: { default: '' }, raw: { default: null } },
      parseDOM: [{ tag: 'ins[data-track]', getAttrs: (el) => trackedAttrs(el as HTMLElement) }],
      toDOM: (m) => ['ins', trackedDom('t-ins', m), 0],
    },
    tracked_del: {
      inclusive: false, attrs: { author: { default: '' }, date: { default: '' }, raw: { default: null } },
      parseDOM: [{ tag: 'del[data-track]', getAttrs: (el) => trackedAttrs(el as HTMLElement) }],
      toDOM: (m) => ['del', trackedDom('t-del', m), 0],
    },
    link: {
      inclusive: false,
      attrs: { href: { default: null }, rId: { default: null }, anchor: { default: null }, rawAttrs: { default: '' } },
      parseDOM: ['a[href]', 'a[data-href]'].map((tag) => ({ tag, getAttrs: (el: HTMLElement | string) => {
        const e = el as HTMLElement;
        const href = e.getAttribute('href') ?? e.getAttribute('data-href') ?? '';
        return /^(https?:|mailto:)/i.test(href) ? { href } : false;
      } })),
      toDOM: (m) => ['a', { 'data-href': m.attrs.href ?? m.attrs.anchor ?? '', title: 'Ctrl-click to open' }, 0],
    },
    bold: {
      parseDOM: [{ tag: 'b' }, { tag: 'strong' }, { style: 'font-weight', getAttrs: (v) => (/^(bold|[6-9]00)$/.test(v as string) ? null : false) }],
      toDOM: () => ['strong', { style: 'font-weight:700' }, 0],
    },
    italic: { parseDOM: [{ tag: 'i' }, { tag: 'em' }, { style: 'font-style=italic' }], toDOM: () => ['em', { style: 'font-style:italic' }, 0] },
    underline: {
      attrs: { val: { default: 'single' }, raw: { default: null } },
      parseDOM: [{ tag: 'u' }, { style: 'text-decoration', getAttrs: (v) => ((v as string).includes('underline') ? {} : false) }],
      toDOM: () => ['u', 0],
    },
    strike: { parseDOM: [{ tag: 's' }, { tag: 'strike' }, { tag: 'del' }, { style: 'text-decoration', getAttrs: (v) => ((v as string).includes('line-through') ? {} : false) }], toDOM: () => ['s', 0] },
    vert: {
      excludes: 'vert', attrs: { val: {} },
      parseDOM: [{ tag: 'sub', attrs: { val: 'subscript' } }, { tag: 'sup', attrs: { val: 'superscript' } }],
      toDOM: (m) => [m.attrs.val === 'subscript' ? 'sub' : 'sup', 0],
    },
    color: {
      excludes: 'color', attrs: { val: {}, raw: { default: null } },
      parseDOM: [{ style: 'color', getAttrs: (v) => { const h = cssToHex(v as string); return h ? { val: h } : false; } }],
      toDOM: (m) => ['span', { style: `color:${safeHex(m.attrs.val) ?? 'inherit'}` }, 0],
    },
    highlight: {
      excludes: 'highlight', attrs: { val: {} },
      parseDOM: [{ tag: 'span[data-hl]', getAttrs: (el) => ({ val: (el as HTMLElement).dataset.hl }) }],
      toDOM: (m) => ['span', { 'data-hl': m.attrs.val, style: `background:${HIGHLIGHTS[m.attrs.val] ?? 'transparent'}` }, 0],
    },
    font: {
      excludes: 'font', attrs: { name: {}, raw: { default: null } },
      parseDOM: [{ tag: 'span[data-font]', getAttrs: (el) => ({ name: (el as HTMLElement).dataset.font }) }],
      toDOM: (m) => ['span', { 'data-font': m.attrs.name, style: `font-family:${fontStack(m.attrs.name)}` }, 0],
    },
    size: {
      excludes: 'size', attrs: { val: {} },
      parseDOM: [{ tag: 'span[data-size]', getAttrs: (el) => ({ val: Number((el as HTMLElement).dataset.size) || 22 }) }],
      toDOM: (m) => ['span', { 'data-size': String(m.attrs.val), style: `font-size:${Number(m.attrs.val) / 2}pt` }, 0],
    },
    rstyle: {
      excludes: 'rstyle', attrs: { id: {} },
      parseDOM: [{ tag: 'span[data-rstyle]', getAttrs: (el) => ({ id: (el as HTMLElement).dataset.rstyle }) }],
      toDOM: (m) => ['span', { 'data-rstyle': m.attrs.id, class: `cs-${cssIdent(m.attrs.id)}` }, 0],
    },
    rextra: {
      excludes: 'rextra', attrs: { xml: {} },
      parseDOM: [{ tag: 'span[data-rextra]', getAttrs: (el) => {
        const e = el as HTMLElement;
        if (Number(e.dataset.session) !== session.id) return false;
        try { return { xml: JSON.parse(e.dataset.rextra!) }; } catch { return false; }
      } }],
      toDOM: (m) => {
        const x = (m.attrs.xml as string[]).join('');
        const s: string[] = [];
        if (/:caps\b(?![^>]*val="(0|false)")/.test(x)) s.push('text-transform:uppercase');
        if (/:smallCaps\b(?![^>]*val="(0|false)")/.test(x)) s.push('font-variant:small-caps');
        if (/:vanish\b(?![^>]*val="(0|false)")/.test(x)) s.push('opacity:.45');
        if (/:b\b[^>]*val="(0|false)"/.test(x)) s.push('font-weight:400');
        if (/:i\b[^>]*val="(0|false)"/.test(x)) s.push('font-style:normal');
        const shd = /:shd\b[^>]*fill="([0-9a-fA-F]{6})"/.exec(x);
        if (shd) s.push(`background:#${shd[1]}`);
        const attrs: Record<string, string> = { 'data-rextra': JSON.stringify(m.attrs.xml), 'data-session': String(session.id) };
        if (s.length) attrs.style = s.join(';');
        return ['span', attrs, 0];
      },
    },
  },
});

export function authorColor(author: string) {
  let h = 0;
  for (let i = 0; i < author.length; i++) h = (h * 31 + author.charCodeAt(i)) >>> 0;
  return h % 6;
}
function trackedDom(cls: string, m: Mark) {
  return {
    class: `${cls} t-a${authorColor(m.attrs.author)}`, title: m.attrs.author,
    'data-track': JSON.stringify({ author: m.attrs.author, date: m.attrs.date }),
  };
}
function trackedAttrs(el: HTMLElement) {
  try { return JSON.parse(el.dataset.track!); } catch { return false; }
}

function tableClass(tblPr: string) {
  return /:tblBorders\b/.test(tblPr) || /:tblStyle\b[^>]*val="(TableGrid|[^"]*Grid[^"]*)"/.test(tblPr) ? 'tb-grid' : 'tb-plain';
}
function cssToHex(v: string): string | null {
  const m = /^#([0-9a-f]{6})$/i.exec(v.trim());
  if (m) return m[1].toUpperCase();
  const r = /^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/.exec(v.trim());
  if (!r) return null;
  return [r[1], r[2], r[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** Adds the raw Word attributes to clipboard HTML only, so ordinary rendering stays light. */
export function clipboardSerializer() {
  const nodes = DOMSerializer.nodesFromSchema(schema);
  for (const name of ['paragraph', 'table', 'table_row', 'table_cell', 'opaque_block']) {
    const base = nodes[name];
    nodes[name] = (n: PMNode) => {
      const spec = [...(base(n) as unknown as unknown[])];
      const { srcId: _drop, ...attrs } = n.attrs;
      const extra = { 'data-folio': JSON.stringify(attrs), 'data-session': String(session.id) };
      if (spec[1] && typeof spec[1] === 'object' && !Array.isArray(spec[1])) spec[1] = { ...(spec[1] as object), ...extra };
      else spec.splice(1, 0, extra);
      return spec as unknown as DOMOutputSpec;
    };
  }
  return new DOMSerializer(nodes, DOMSerializer.marksFromSchema(schema));
}

export type { Attrs, Fragment, Mark };
