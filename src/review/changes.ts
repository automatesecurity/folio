import { Mark, Node as PMNode } from 'prosemirror-model';
import { EditorState, Transaction } from 'prosemirror-state';
import type { Revision, WordDoc } from '../docx/document';
import { schema } from '../editor/schema';

export type ChangeKind = 'ins' | 'del' | 'mark-ins' | 'mark-del' | 'row-ins' | 'row-del' | 'format' | 'para-format';

export interface Change {
  kind: ChangeKind;
  from: number;
  to: number;
  author: string;
  date: string;
  text: string;
}

const tagOf = (x: string) => /^<(?:[\w.-]+:)?([\w.-]+)/.exec(x)?.[1] ?? '';
const rExtra = (n: PMNode) => schema.marks.rextra.isInSet(n.marks);
const hasRunChange = (n: PMNode) => !!rExtra(n) && (rExtra(n)!.attrs.xml as string[]).some((x) => tagOf(x) === 'rPrChange');
const revOf = (xml: string) => ({ author: /\bauthor="([^"]*)"/.exec(xml)?.[1] ?? '', date: /\bdate="([^"]*)"/.exec(xml)?.[1] ?? '' });

export function collectChanges(doc: PMNode): Change[] {
  const out: Change[] = [];
  doc.descendants((n, pos) => {
    if (n.type === schema.nodes.table_row && n.attrs.rowChange) {
      const c = n.attrs.rowChange as { kind: string; rev: Revision };
      out.push({ kind: c.kind === 'ins' ? 'row-ins' : 'row-del', from: pos, to: pos + n.nodeSize, author: c.rev.author, date: c.rev.date, text: 'Table row' });
      return true;
    }
    if (n.type !== schema.nodes.paragraph) return true;
    const start = pos + 1;
    let open: Change | null = null;
    const flush = () => { if (open) out.push(open); open = null; };
    n.forEach((c, off) => {
      const del = schema.marks.tracked_del.isInSet(c.marks), ins = schema.marks.tracked_ins.isInSet(c.marks);
      const mark: Mark | null = del ?? ins ?? null;
      const kind: ChangeKind | null = del ? 'del' : ins ? 'ins' : hasRunChange(c) ? 'format' : null;
      if (!kind) { flush(); return; }
      const rev = mark ? { author: mark.attrs.author as string, date: mark.attrs.date as string } : revOf((rExtra(c)!.attrs.xml as string[]).find((x) => tagOf(x) === 'rPrChange')!);
      const text = c.isText ? c.text! : (c.attrs.label as string) || (c.type === schema.nodes.hard_break ? ' ' : '');
      const current = open as Change | null;
      if (current && current.kind === kind && current.author === rev.author && current.to === start + off) {
        current.to = start + off + c.nodeSize;
        current.text += text;
      } else {
        flush();
        open = { kind, from: start + off, to: start + off + c.nodeSize, author: rev.author, date: rev.date, text };
      }
    });
    flush();
    const end = pos + n.nodeSize - 1;
    const pm = n.attrs.pmark as { ins?: Revision; del?: Revision } | null;
    if (pm?.del) out.push({ kind: 'mark-del', from: end, to: end + 1, author: pm.del.author, date: pm.del.date, text: 'Paragraph break' });
    if (pm?.ins) out.push({ kind: 'mark-ins', from: end, to: end + 1, author: pm.ins.author, date: pm.ins.date, text: 'Paragraph break' });
    const pc = (n.attrs.pExtra as string[]).find((x) => tagOf(x) === 'pPrChange');
    if (pc) out.push({ kind: 'para-format', from: pos, to: end + 1, ...revOf(pc), text: 'Paragraph formatting' });
    return false;
  });
  return out.sort((a, b) => a.from - b.from);
}

/** Merging keeps the following paragraph's mark, so its formatting is what survives. */
function joinNext(tr: Transaction, end: number) {
  const $end = tr.doc.resolve(end);
  const para = $end.parent;
  const next = tr.doc.resolve(end + 1).nodeAfter;
  const pos = $end.before();
  if (next?.type !== schema.nodes.paragraph) {
    tr.setNodeMarkup(pos, undefined, { ...para.attrs, pmark: null });
    return;
  }
  tr.setNodeMarkup(pos, undefined, { ...next.attrs, srcId: null, rawAttrs: '' });
  tr.join(end + 1);
}

function clearMark(tr: Transaction, end: number, key: 'ins' | 'del') {
  const $end = tr.doc.resolve(end);
  const para = $end.parent;
  const pm = { ...(para.attrs.pmark ?? {}) };
  delete pm[key];
  tr.setNodeMarkup($end.before(), undefined, { ...para.attrs, pmark: Object.keys(pm).length ? pm : null });
}

function deleteRow(tr: Transaction, from: number) {
  const $row = tr.doc.resolve(from);
  if ($row.parent.childCount === 1) tr.delete($row.before(), $row.after());
  else tr.delete(from, from + tr.doc.nodeAt(from)!.nodeSize);
}

function clearRow(tr: Transaction, from: number) {
  const row = tr.doc.nodeAt(from)!;
  tr.setNodeMarkup(from, undefined, { ...row.attrs, rowChange: null });
}

function resolveFormat(tr: Transaction, d: WordDoc, c: Change, accept: boolean) {
  tr.doc.nodesBetween(c.from, c.to, (n, pos) => {
    if (!n.isInline || !hasRunChange(n)) return true;
    const from = Math.max(pos, c.from), to = Math.min(pos + n.nodeSize, c.to);
    const extra = rExtra(n)!;
    if (accept) {
      const rest = (extra.attrs.xml as string[]).filter((x) => tagOf(x) !== 'rPrChange');
      tr.removeMark(from, to, schema.marks.rextra);
      if (rest.length) tr.addMark(from, to, schema.marks.rextra.create({ xml: rest }));
      return false;
    }
    const old = d.marksFromChange(extra.attrs.xml);
    for (const m of n.marks) if (![schema.marks.link, schema.marks.tracked_ins, schema.marks.tracked_del].includes(m.type)) tr.removeMark(from, to, m.type);
    for (const m of old ?? []) tr.addMark(from, to, m);
    return false;
  });
}

function resolveParaFormat(tr: Transaction, d: WordDoc, c: Change, accept: boolean) {
  const para = tr.doc.nodeAt(c.from)!;
  const extra = para.attrs.pExtra as string[];
  if (accept) { tr.setNodeMarkup(c.from, undefined, { ...para.attrs, pExtra: extra.filter((x) => tagOf(x) !== 'pPrChange') }); return; }
  const old = d.paraFromChange(extra) ?? {};
  const keep = extra.filter((x) => tagOf(x) === 'sectPr');
  tr.setNodeMarkup(c.from, undefined, { ...para.attrs, ...old, pExtra: [...keep, ...((old.pExtra as string[]) ?? [])] });
}

export function applyChange(tr: Transaction, d: WordDoc, c: Change, accept: boolean) {
  switch (c.kind) {
    case 'ins':
      if (accept) tr.removeMark(c.from, c.to, schema.marks.tracked_ins); else tr.delete(c.from, c.to);
      break;
    case 'del':
      if (accept) tr.delete(c.from, c.to); else tr.removeMark(c.from, c.to, schema.marks.tracked_del);
      break;
    case 'mark-ins':
      if (accept) clearMark(tr, c.from, 'ins'); else joinNext(tr, c.from);
      break;
    case 'mark-del':
      if (accept) joinNext(tr, c.from); else clearMark(tr, c.from, 'del');
      break;
    case 'row-ins':
      if (accept) clearRow(tr, c.from); else deleteRow(tr, c.from);
      break;
    case 'row-del':
      if (accept) deleteRow(tr, c.from); else clearRow(tr, c.from);
      break;
    case 'format': resolveFormat(tr, d, c, accept); break;
    case 'para-format': resolveParaFormat(tr, d, c, accept); break;
  }
}

/** Later changes go first, so earlier positions stay valid. A row sorts before its own text. */
export function applyChanges(tr: Transaction, d: WordDoc, list: Change[], accept: boolean) {
  for (const c of [...list].sort((a, b) => b.from - a.from)) applyChange(tr, d, c, accept);
  return tr;
}

/** A copy with every change settled and comment anchors removed, for clean output. */
export function settledCopy(d: WordDoc, doc: PMNode, accept: boolean): PMNode {
  const state = EditorState.create({ doc, schema });
  const tr = applyChanges(state.tr, d, collectChanges(doc), accept);
  const anchors: [number, number][] = [];
  tr.doc.descendants((n, pos) => {
    if (n.type === schema.nodes.opaque_inline && (n.attrs.kind === 'comment' || (n.attrs.kind === 'range' && /commentRange/.test(n.attrs.xml)))) anchors.push([pos, pos + n.nodeSize]);
  });
  for (const [a, b] of anchors.reverse()) tr.delete(a, b);
  return tr.doc;
}
