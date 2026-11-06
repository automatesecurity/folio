import { Mark, Node as PMNode, Slice } from 'prosemirror-model';
import { EditorState, TextSelection, Transaction } from 'prosemirror-state';
import { ReplaceStep } from 'prosemirror-transform';
import { schema } from '../editor/schema';

export interface Reviewer { author: string; date: string }

const HIDDEN = new Set(['bookmark', 'range', 'fieldPart', 'comment']);
const isHidden = (n: PMNode) => n.type === schema.nodes.opaque_inline && HIDDEN.has(n.attrs.kind);
const insMark = (me: Reviewer) => schema.marks.tracked_ins.create({ author: me.author, date: me.date, raw: null });
const delMark = (me: Reviewer) => schema.marks.tracked_del.create({ author: me.author, date: me.date, raw: null });
const rev = (me: Reviewer) => ({ author: me.author, date: me.date, raw: null });
const mine = (m: Mark | undefined | null, me: Reviewer) => !!m && m.attrs.author === me.author;

interface Op { from: number; to: number; slice: Slice }

function contains(slice: Slice, test: (n: PMNode) => boolean) {
  let found = false;
  slice.content.descendants((n) => { if (test(n)) found = true; return !found; });
  return found;
}

function operations(state: EditorState, tr: Transaction): Op[] {
  const ops: Op[] = [];
  tr.steps.forEach((step, i) => {
    const s = step as ReplaceStep;
    // Mapping.slice().invert() walks every map in the transaction, so step back by hand.
    const back = (pos: number, assoc: number) => {
      for (let j = i - 1; j >= 0; j--) pos = tr.mapping.maps[j].invert().map(pos, assoc);
      return pos;
    };
    ops.push({ from: back(s.from, -1), to: back(s.to, 1), slice: s.slice });
  });
  ops.sort((a, b) => b.from - a.from);
  const independent = ops.every((op, i) => i === ops.length - 1 || ops[i + 1].to < op.from);
  if (independent) return ops;
  // Steps that overlap, like replacing a selection and then splitting, are read as one edit.
  const start = state.doc.content.findDiffStart(tr.doc.content);
  if (start === null) return [];
  let { a, b } = state.doc.content.findDiffEnd(tr.doc.content)!;
  const overlap = start - Math.min(a, b);
  if (overlap > 0) { a += overlap; b += overlap; }
  return [{ from: start, to: a, slice: tr.doc.slice(start, b) }];
}

function markInserted(out: Transaction, a: number, b: number, me: Reviewer) {
  if (b <= a) return;
  out.removeMark(a, b, schema.marks.tracked_del);
  out.removeMark(a, b, schema.marks.tracked_ins);
  out.addMark(a, b, insMark(me));
  const updates: { pos: number; node: PMNode }[] = [];
  out.doc.nodesBetween(a, b, (n, pos) => {
    if (n.type === schema.nodes.paragraph) {
      const end = pos + n.nodeSize - 1;
      if (end >= a && end < b) updates.push({ pos, node: n });
      return false;
    }
    if (n.type === schema.nodes.table_row && pos >= a && pos + n.nodeSize <= b) updates.push({ pos, node: n });
    return true;
  });
  for (const { pos, node } of updates) {
    if (node.type === schema.nodes.table_row) out.setNodeMarkup(pos, undefined, { ...node.attrs, srcId: null, rowChange: { kind: 'ins', rev: rev(me) } });
    else out.setNodeMarkup(pos, undefined, { ...node.attrs, pmark: { ...(node.attrs.pmark ?? {}), ins: rev(me) } });
  }
}

/** Marks a range deleted, except what this author inserted, which really goes away. */
function markDeleted(out: Transaction, doc: PMNode, from: number, to: number, me: Reviewer) {
  const real: [number, number][] = [];
  doc.nodesBetween(from, to, (n, pos) => {
    if (n.type === schema.nodes.table_row && pos >= from && pos + n.nodeSize <= to) {
      const change = n.attrs.rowChange;
      if (change?.kind === 'ins' && change.rev.author === me.author) { real.push([pos, pos + n.nodeSize]); return false; }
      if (!change) out.setNodeMarkup(pos, undefined, { ...n.attrs, srcId: null, rowChange: { kind: 'del', rev: rev(me) } });
      return true;
    }
    if (n.type === schema.nodes.opaque_block && pos >= from && pos + n.nodeSize <= to) { real.push([pos, pos + n.nodeSize]); return false; }
    if (n.type === schema.nodes.paragraph) {
      const end = pos + n.nodeSize - 1;
      const next = doc.resolve(end + 1).nodeAfter;
      if (end >= from && end < to && next?.type === schema.nodes.paragraph) {
        const pm = n.attrs.pmark;
        if (pm?.ins?.author === me.author) real.push([end, end + 2]);
        else if (!pm?.del) out.setNodeMarkup(pos, undefined, { ...n.attrs, pmark: { ...(pm ?? {}), del: rev(me) } });
      }
      return true;
    }
    if (!n.isInline) return true;
    if (isHidden(n)) return false;
    const s = Math.max(pos, from), e = Math.min(pos + n.nodeSize, to);
    const ins = schema.marks.tracked_ins.isInSet(n.marks);
    if (mine(ins, me)) real.push([s, e]);
    else if (!schema.marks.tracked_del.isInSet(n.marks)) out.addMark(s, e, delMark(me));
    return false;
  });
  real.sort((x, y) => y[0] - x[0]);
  for (const [s, e] of real) out.delete(s, e);
}

/**
 * Turns an ordinary edit into a tracked one. Deleted content stays in the document
 * with a mark, and new content is added after it with another mark.
 * Formatting and table column changes pass through unchanged.
 */
export function trackTransaction(state: EditorState, tr: Transaction, me: Reviewer): Transaction {
  if (!tr.docChanged || !tr.steps.every((s) => s instanceof ReplaceStep)) return tr;
  const ops = operations(state, tr);
  if (!ops.length) return tr;
  for (const op of ops) if (state.doc.resolve(op.from).parent.type === schema.nodes.table_row) return tr;
  try {
    const out = state.tr;
    let caret: number | null = null;
    for (const op of ops) {
      let slice = op.slice;
      const rows = contains(slice, (n) => n.type === schema.nodes.table_row) || slice.content.firstChild?.type === schema.nodes.table_row;
      if (op.to > op.from && slice.openStart === 0 && !contains(slice, (n) => n.isInline) && !rows) slice = Slice.empty;
      let insertedEnd: number | null = null;
      if (slice.size) {
        const index = out.steps.length;
        out.replace(op.to, op.to, slice);
        let a = op.to, b = op.to;
        out.mapping.maps[index]?.forEach((_s, _e, ns, ne) => { a = ns; b = ne; });
        markInserted(out, a, b, me);
        insertedEnd = b;
      }
      if (op.to > op.from) {
        const at = out.steps.length;
        markDeleted(out, state.doc, op.from, op.to, me);
        if (insertedEnd !== null) insertedEnd = out.mapping.slice(at).map(insertedEnd);
      }
      if (ops.length === 1) {
        const old = state.selection;
        caret = insertedEnd ?? out.mapping.map(old.empty && old.from === op.to ? op.from : op.to, -1);
      }
    }
    if (caret === null) caret = out.mapping.map(state.selection.to);
    out.setSelection(TextSelection.near(out.doc.resolve(Math.min(caret, out.doc.content.size)), -1));
    if (tr.storedMarksSet) out.setStoredMarks(tr.storedMarks);
    if (tr.scrolledIntoView) out.scrollIntoView();
    return out;
  } catch {
    return tr;
  }
}
