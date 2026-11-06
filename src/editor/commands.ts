import { Attrs, MarkType, Node as PMNode } from 'prosemirror-model';
import { Command, EditorState, TextSelection, Transaction } from 'prosemirror-state';
import { isInTable } from 'prosemirror-tables';
import type { WordDoc } from '../docx/document';
import { schema } from './schema';

export function selectedParagraphs(state: EditorState) {
  const out: { node: PMNode; pos: number }[] = [];
  for (const r of state.selection.ranges) {
    state.doc.nodesBetween(r.$from.pos, r.$to.pos, (n, pos) => {
      if (n.type === schema.nodes.paragraph) { out.push({ node: n, pos }); return false; }
      return true;
    });
  }
  return out;
}

function updateParagraphs(fn: (n: PMNode, i: number, tr: Transaction) => Attrs | null): Command {
  return (state, dispatch) => {
    const ps = selectedParagraphs(state);
    if (!ps.length) return false;
    if (dispatch) {
      const tr = state.tr;
      ps.forEach((p, i) => { const a = fn(p.node, i, tr); if (a) tr.setNodeMarkup(p.pos, undefined, { ...p.node.attrs, ...a }, p.node.marks); });
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export const setStyle = (id: string | null) => updateParagraphs(() => ({ pStyle: id }));
export const setAlign = (jc: string) => updateParagraphs(() => ({ jc }));

export function toggleList(doc: WordDoc, kind: 'bullet' | 'ordered'): Command {
  return (state, dispatch, view) => {
    const ps = selectedParagraphs(state);
    if (!ps.length) return false;
    const on = ps.every((p) => doc.numbering.kindOf(p.node.attrs.numId) === kind);
    if (on) return updateParagraphs(() => ({ numId: null, ilvl: null, ind: null }))(state, dispatch, view);
    const before = state.doc.resolve(ps[0].pos).nodeBefore;
    const after = state.doc.resolve(ps[ps.length - 1].pos + ps[ps.length - 1].node.nodeSize).nodeAfter;
    const near = [before, after].find((n) => n?.type === schema.nodes.paragraph && doc.numbering.kindOf(n.attrs.numId) === kind);
    const numId = near?.attrs.numId ?? (dispatch ? doc.numbering.newList(kind, doc.w, true) : '0');
    return updateParagraphs((n) => ({ numId, ilvl: n.attrs.ilvl ?? 0, ind: null }))(state, dispatch, view);
  };
}

export function indent(dir: 1 | -1): Command {
  return updateParagraphs((n) => {
    if (n.attrs.numId) return { ilvl: Math.min(8, Math.max(0, (n.attrs.ilvl ?? 0) + dir)) };
    const ind = { ...(n.attrs.ind ?? {}) } as Record<string, string>;
    const left = Math.max(0, Number(ind.left ?? ind.start ?? 0) + dir * 720);
    delete ind.start;
    if (left) ind.left = String(left); else delete ind.left;
    return { ind: Object.keys(ind).length ? ind : null };
  });
}

/** Enter: continues the paragraph's formatting, follows the style's "next", and ends empty lists. */
export function splitParagraph(doc: WordDoc): Command {
  return (state, dispatch, view) => {
    const sel = state.selection;
    if (!(sel instanceof TextSelection)) return false;
    const { $from } = sel;
    const p = $from.parent;
    if (p.type !== schema.nodes.paragraph) return false;
    if (p.content.size === 0 && p.attrs.numId) return updateParagraphs(() => ({ numId: null, ilvl: null, ind: null }))(state, dispatch, view);
    if (dispatch) {
      const atEnd = $from.parentOffset === p.content.size;
      const style = atEnd && p.attrs.pStyle ? doc.styles.styles.get(p.attrs.pStyle)?.next ?? p.attrs.pStyle : p.attrs.pStyle;
      const pExtra = (p.attrs.pExtra as string[]).filter((x) => !/^<[\w.-]*:?(sectPr|pPrChange)\b/.test(x));
      const attrs = { ...p.attrs, srcId: null, rawAttrs: '', pExtra, pStyle: style, pmark: null };
      const tr = state.tr;
      if (!sel.empty) tr.deleteSelection();
      const marks = state.storedMarks ?? $from.marks();
      tr.split(tr.mapping.map($from.pos), 1, [{ type: schema.nodes.paragraph, attrs }]);
      tr.setStoredMarks(marks);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** Backspace at the start of a list item removes the marker before merging paragraphs. */
export const unlistAtStart: Command = (state, dispatch) => {
  const { $from, empty } = state.selection;
  if (!empty || $from.parentOffset !== 0 || !$from.parent.attrs.numId) return false;
  if (dispatch) dispatch(state.tr.setNodeMarkup($from.before(), undefined, { ...$from.parent.attrs, numId: null, ilvl: null, ind: null }));
  return true;
};

export function setMark(type: MarkType, attrs: Attrs | null): Command {
  return (state, dispatch) => {
    const { empty, $cursor, ranges } = state.selection as TextSelection & { $cursor?: unknown };
    if (dispatch) {
      const tr = state.tr;
      if (empty && $cursor) {
        tr.removeStoredMark(type);
        if (attrs) tr.addStoredMark(type.create(attrs));
      } else {
        for (const { $from, $to } of ranges) {
          tr.removeMark($from.pos, $to.pos, type);
          if (attrs) tr.addMark($from.pos, $to.pos, type.create(attrs));
        }
      }
      dispatch(tr);
    }
    return true;
  };
}

export function insertTable(doc: WordDoc, rows: number, cols: number): Command {
  return (state, dispatch) => {
    if (dispatch) {
      const w = doc.w;
      const total = doc.page.width - doc.page.left - doc.page.right;
      const each = Math.floor(total / cols);
      const b = (s: string) => `<${w}${s} ${w}val="single" ${w}sz="4" ${w}space="0" ${w}color="auto"/>`;
      const tblPr = `<${w}tblPr><${w}tblW ${w}w="${each * cols}" ${w}type="dxa"/><${w}tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(b).join('')}</${w}tblBorders><${w}tblLook ${w}val="04A0"/></${w}tblPr>`;
      const mk = () => schema.nodes.table_cell.create({ colwidth: [each] }, schema.nodes.paragraph.create());
      const table = schema.nodes.table.create({ tblPr }, Array.from({ length: rows }, () =>
        schema.nodes.table_row.create(null, Array.from({ length: cols }, mk))));
      const tr = state.tr.replaceSelectionWith(table);
      const after = schema.nodes.paragraph.create();
      const end = tr.mapping.map(state.selection.to);
      if (!tr.doc.resolve(end).nodeAfter) tr.insert(end, after);
      let pos = -1;
      tr.doc.nodesBetween(state.selection.from, tr.doc.content.size, (n, p) => { if (pos < 0 && n.type === schema.nodes.table_cell) { pos = p + 2; } return pos < 0; });
      if (pos > 0) tr.setSelection(TextSelection.create(tr.doc, pos));
      dispatch(tr.scrollIntoView());
    }
    return !isInTable(state) || true;
  };
}

export const insertPageBreak: Command = (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.page_break.create(), false).scrollIntoView());
  return true;
};
export const insertHardBreak: Command = (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.hard_break.create(), true).scrollIntoView());
  return true;
};

export function linkRange(state: EditorState) {
  const { from, to, $from } = state.selection;
  const type = schema.marks.link;
  if (from !== to) { const m = state.doc.rangeHasMark(from, to, type); return { from, to, href: m ? findMark(state, from, to) : '' }; }
  const m = type.isInSet($from.marks());
  if (!m) return { from, to, href: '' };
  let start = from, end = from;
  const parent = $from.parent, base = $from.start();
  parent.forEach((c, off) => {
    if (m.isInSet(c.marks)) {
      if (base + off <= start && base + off + c.nodeSize >= start) { start = base + off; end = base + off + c.nodeSize; }
    }
  });
  return { from: start, to: end, href: (m.attrs.href ?? m.attrs.anchor ?? '') as string };
}
function findMark(state: EditorState, from: number, to: number) {
  let href = '';
  state.doc.nodesBetween(from, to, (n) => { const m = schema.marks.link.isInSet(n.marks); if (m) href = m.attrs.href ?? ''; });
  return href;
}
