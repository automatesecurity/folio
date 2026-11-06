import { Node as PMNode } from 'prosemirror-model';
import { EditorState, Plugin, PluginKey, TextSelection, Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, EditorView } from 'prosemirror-view';
import type { WordDoc } from '../docx/document';
import { ListCounter, parseParaFragment, tw } from '../docx/styles';
import { schema } from './schema';

// List markers ---------------------------------------------------------------

export function effectiveNum(doc: WordDoc, p: PMNode): { numId: string; ilvl: number } | null {
  const a = p.attrs;
  if (a.numId) return a.numId === '0' ? null : { numId: a.numId, ilvl: a.ilvl ?? 0 };
  if (a.pStyle) {
    const s = doc.styles.resolvedPara(a.pStyle);
    if (s.numId && s.numId !== '0') return { numId: s.numId, ilvl: s.ilvl ?? 0 };
  }
  return null;
}

export function listPlugin(getDoc: () => WordDoc) {
  const build = (state: { doc: PMNode }) => {
    const d = getDoc();
    const counter = new ListCounter(d.numbering);
    const decos: Decoration[] = [];
    state.doc.descendants((n, pos) => {
      if (n.type !== schema.nodes.paragraph) return true;
      const num = effectiveNum(d, n);
      if (!num) return false;
      const r = counter.next(num.numId, num.ilvl);
      if (r) {
        const fromStyle = !n.attrs.numId;
        const style = fromStyle && r.level.left !== undefined
          ? `margin-left:${tw(r.level.left)}px;text-indent:${tw(r.level.first ?? 0)}px;--hang:${tw(-(r.level.first ?? 0))}px` : undefined;
        decos.push(Decoration.node(pos, pos + n.nodeSize, { 'data-marker': r.label, ...(style ? { style } : {}) }));
      }
      return false;
    });
    return DecorationSet.create(state.doc, decos);
  };
  return new Plugin<DecorationSet>({
    key: new PluginKey('lists'),
    state: { init: (_, s) => build(s), apply: (tr, old, _, s) => (tr.docChanged || tr.getMeta('refreshLists') ? build(s) : old) },
    props: { decorations(s) { return this.getState(s); } },
  });
}

// Pagination -------------------------------------------------------------------

export interface PageInfo { top: number; height: number }
export interface Layout { pushes: Map<number, number>; pages: PageInfo[]; blockPage: number[] }
export const paginateKey = new PluginKey<DecorationSet>('paginate');
export const PAGE_GAP = 20;

export function computeLayout(view: EditorView, doc: WordDoc): Layout {
  const pageH = tw(doc.page.height), mTop = tw(doc.effTop), mBot = tw(doc.effBottom), body = pageH - mTop - mBot;
  const kids = view.dom.children;
  const pushes = new Map<number, number>();
  const pages: PageInfo[] = [];
  const blockPage: number[] = [];
  let y = 0, pageTop = 0, forceNext = false, pos = 0, index = 0;
  const finish = () => { const h = Math.max(pageH, y + mTop + mBot); pages.push({ top: pageTop, height: h }); return h; };
  view.state.doc.forEach((node) => {
    const el = kids[index] as HTMLElement | undefined;
    const h = el?.offsetHeight ?? 0;
    const props = node.type === schema.nodes.paragraph ? doc.styles.resolvedPara(node.attrs.pStyle) : {};
    const own = node.type === schema.nodes.paragraph && (node.attrs.pExtra as string[]).length ? parseParaFragment((node.attrs.pExtra as string[]).join('')) : {};
    const keepNext = (own.keepNext ?? props.keepNext) === true;
    const nextEl = kids[index + 1] as HTMLElement | undefined;
    const needed = h + (keepNext && nextEl ? Math.min(nextEl.offsetHeight, 60) : 0);
    const leadingBreak = node.firstChild?.type === schema.nodes.page_break;
    const forced = forceNext || leadingBreak || (own.pageBreakBefore ?? props.pageBreakBefore) === true;
    forceNext = false;
    if ((forced && (y > 0 || pages.length > 0)) || (y > 0 && y + needed > body)) {
      const used = finish();
      pushes.set(pos, used - y + PAGE_GAP);
      pageTop += used + PAGE_GAP;
      y = 0;
    }
    blockPage.push(pages.length);
    y += h;
    let endsWithBreak = false;
    node.descendants((c, at) => { if (c.type === schema.nodes.page_break && !(leadingBreak && at === 0)) endsWithBreak = true; return !endsWithBreak; });
    forceNext = endsWithBreak;
    pos += node.nodeSize;
    index++;
  });
  finish();
  return { pushes, pages, blockPage };
}

/** Forces a layout pass when something outside the document, like margins, changed. */
export let requestLayout: () => void = () => undefined;

export function paginatePlugin(onLayout: (l: Layout) => void, getDoc: () => WordDoc) {
  let pending = 0;
  let last = '';
  return new Plugin<DecorationSet>({
    key: paginateKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, old) {
        const next = tr.getMeta(paginateKey) as DecorationSet | undefined;
        return next ?? old.map(tr.mapping, tr.doc);
      },
    },
    props: { decorations(s) { return paginateKey.getState(s); } },
    view(view) {
      const run = () => {
        pending = 0;
        const layout = computeLayout(view, getDoc());
        onLayout(layout);
        const sig = [...layout.pushes].join(';');
        if (sig === last) return;
        last = sig;
        const decos: Decoration[] = [];
        view.state.doc.forEach((n, pos) => {
          const m = layout.pushes.get(pos);
          if (m !== undefined) decos.push(Decoration.node(pos, pos + n.nodeSize, { style: `margin-top:${m}px` }));
        });
        view.dispatch(view.state.tr.setMeta(paginateKey, DecorationSet.create(view.state.doc, decos)).setMeta('addToHistory', false));
      };
      const schedule = () => { if (!pending) pending = requestAnimationFrame(run); };
      requestLayout = () => { last = ''; schedule(); };
      schedule();
      const fonts = document.fonts?.ready.then(schedule);
      void fonts;
      return {
        update(v, prev) { if (v.state.doc !== prev.doc) { last = ''; schedule(); } },
        destroy() { if (pending) cancelAnimationFrame(pending); },
      };
    },
  });
}

// Hidden anchors ------------------------------------------------------------------

const isHidden = (n: PMNode | null | undefined) =>
  !!n && n.type === schema.nodes.opaque_inline && ['bookmark', 'range', 'fieldPart'].includes(n.attrs.kind);

/** Bookmarks have no width, so without this the caret would appear to get stuck on them. */
export function hiddenAnchorPlugin() {
  const skip = (state: EditorState, dir: -1 | 1) => {
    const sel = state.selection;
    if (!sel.empty) return null;
    let $p = sel.$from, moved = false;
    for (;;) {
      const n = dir < 0 ? $p.nodeBefore : $p.nodeAfter;
      if (!isHidden(n)) break;
      $p = state.doc.resolve($p.pos + dir);
      moved = true;
    }
    return moved ? $p.pos : null;
  };
  return new Plugin({
    props: {
      handleKeyDown(view, e) {
        if (e.shiftKey || e.ctrlKey || e.metaKey || e.altKey) return false;
        const dir = e.key === 'ArrowLeft' || e.key === 'Backspace' ? -1 : e.key === 'ArrowRight' || e.key === 'Delete' ? 1 : 0;
        if (!dir) return false;
        const pos = skip(view.state, dir);
        if (pos !== null) view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
        return false;
      },
    },
  });
}

// Find ----------------------------------------------------------------------------

export interface FindState { query: string; matchCase: boolean; matches: { from: number; to: number }[]; index: number; deco: DecorationSet }
export const findKey = new PluginKey<FindState>('find');

function search(doc: PMNode, query: string, matchCase: boolean) {
  const out: { from: number; to: number }[] = [];
  if (!query) return out;
  const q = matchCase ? query : query.toLowerCase();
  doc.descendants((n, pos) => {
    if (!n.isTextblock) return true;
    let buf = '', start = -1;
    const flush = () => {
      const hay = matchCase ? buf : buf.toLowerCase();
      for (let i = hay.indexOf(q); i >= 0; i = hay.indexOf(q, i + q.length)) out.push({ from: start + i, to: start + i + q.length });
      buf = ''; start = -1;
    };
    n.forEach((c, off) => {
      if (c.isText) { if (start < 0) start = pos + 1 + off; buf += c.text; } else flush();
    });
    flush();
    return false;
  });
  return out;
}
function decorate(doc: PMNode, f: Omit<FindState, 'deco'>) {
  return DecorationSet.create(doc, f.matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === f.index ? 'find-current' : 'find-hit' })));
}

export function findPlugin() {
  return new Plugin<FindState>({
    key: findKey,
    state: {
      init: () => ({ query: '', matchCase: false, matches: [], index: -1, deco: DecorationSet.empty }),
      apply(tr, old) {
        const meta = tr.getMeta(findKey) as Partial<FindState> | undefined;
        if (!meta && !tr.docChanged) return old;
        const query = meta?.query ?? old.query, matchCase = meta?.matchCase ?? old.matchCase;
        const matches = search(tr.doc, query, matchCase);
        let index = meta?.index ?? old.index;
        if (index >= matches.length) index = matches.length ? 0 : -1;
        if (index < 0 && matches.length) index = 0;
        const f = { query, matchCase, matches, index };
        return { ...f, deco: decorate(tr.doc, f) };
      },
    },
    props: { decorations(s) { return findKey.getState(s)?.deco; } },
  });
}
export function findGo(view: EditorView, dir: 1 | -1) {
  const f = findKey.getState(view.state)!;
  if (!f.matches.length) return;
  const index = (f.index + dir + f.matches.length) % f.matches.length;
  const m = f.matches[index];
  view.dispatch(view.state.tr.setMeta(findKey, { index }).setSelection(TextSelection.create(view.state.doc, m.from, m.to)).scrollIntoView());
}
export function findReplace(view: EditorView, text: string, all: boolean) {
  const f = findKey.getState(view.state)!;
  if (!f.matches.length) return 0;
  const tr: Transaction = view.state.tr;
  const list = all ? [...f.matches].reverse() : [f.matches[Math.max(0, f.index)]];
  for (const m of list) {
    const marks = view.state.doc.resolve(m.from + 1).marks();
    if (text) tr.replaceWith(m.from, m.to, schema.text(text, marks));
    else tr.delete(m.from, m.to);
  }
  view.dispatch(tr);
  return list.length;
}
