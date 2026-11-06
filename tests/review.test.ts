import { describe, it, expect } from 'vitest';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { Node as PMNode } from 'prosemirror-model';
import { EditorState, TextSelection, Transaction } from 'prosemirror-state';
import { anchorNodes, createComment, setDone } from '../src/docx/comments';
import { WordDoc } from '../src/docx/document';
import { exportDocx } from '../src/docx/export';
import { blankDocument } from '../src/docx/template';
import { splitParagraph } from '../src/editor/commands';
import { schema } from '../src/editor/schema';
import { applyChanges, collectChanges, settledCopy } from '../src/review/changes';
import { trackTransaction } from '../src/review/track';

const dir = mkdtempSync(join(tmpdir(), 'folio-review-'));
const dan = { author: 'Dan', date: '2026-11-05T12:00:00Z' };
const sam = { author: 'Sam', date: '2026-11-05T12:00:00Z' };

async function start(...paras: string[]) {
  const d = await blankDocument();
  const body = schema.nodes.doc.create(d.body.attrs, paras.map((t) => schema.nodes.paragraph.create(null, t ? schema.text(t) : undefined)));
  const state = EditorState.create({ doc: body, schema });
  return { d, state };
}
const tracked = (state: EditorState, make: (s: EditorState) => Transaction, me = dan) => state.apply(trackTransaction(state, make(state), me));
const at = (doc: PMNode, needle: string) => {
  let pos = -1;
  doc.descendants((n, p) => { if (pos < 0 && n.isText && n.text!.includes(needle)) pos = p + n.text!.indexOf(needle); });
  return pos;
};
const paras = (doc: PMNode) => { const o: string[] = []; doc.forEach((n) => o.push(n.textContent)); return o; };
const settle = (d: WordDoc, doc: PMNode, accept: boolean) => paras(settledCopy(d, doc, accept));
const xmlOf = async (b: Uint8Array, path = 'word/document.xml') => (await JSZip.loadAsync(b)).file(path)!.async('string');
const py = (code: string, b: Uint8Array) => {
  const p = join(dir, `${Math.random().toString(36).slice(2)}.docx`);
  writeFileSync(p, b);
  return execFileSync('python3', ['-c', code, p]).toString().trim();
};

describe('tracking edits', () => {
  it('keeps deleted text and adds typed text after it', async () => {
    const { d, state } = await start('The quick fox');
    const from = at(state.doc, 'quick');
    const s = tracked(state, (st) => st.tr.insertText('slow', from, from + 5));
    expect(s.doc.textContent).toBe('The quickslow fox');
    const kinds = collectChanges(s.doc).map((c) => [c.kind, c.text]);
    expect(kinds).toEqual([['del', 'quick'], ['ins', 'slow']]);
    expect(settle(d, s.doc, true)).toEqual(['The slow fox']);
    expect(settle(d, s.doc, false)).toEqual(['The quick fox']);
  });

  it('removes your own tracked insertion outright when you delete it', async () => {
    const { state } = await start('Hello');
    let s = tracked(state, (st) => st.tr.insertText(' there', 6));
    const from = at(s.doc, ' there');
    s = tracked(s, (st) => st.tr.delete(from, from + 6));
    expect(s.doc.textContent).toBe('Hello');
    expect(collectChanges(s.doc)).toEqual([]);
  });

  it("marks someone else's insertion as deleted instead of removing it", async () => {
    const { d, state } = await start('Hello');
    let s = tracked(state, (st) => st.tr.insertText(' there', 6), sam);
    const from = at(s.doc, ' there');
    s = tracked(s, (st) => st.tr.delete(from, from + 6), dan);
    expect(s.doc.textContent).toBe('Hello there');
    expect(collectChanges(s.doc).map((c) => c.kind)).toEqual(['del']);
    expect(settle(d, s.doc, true)).toEqual(['Hello']);
  });

  it('tracks Enter as an inserted paragraph mark', async () => {
    const { d, state } = await start('OneTwo');
    const pos = at(state.doc, 'Two');
    const placed = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
    let captured: Transaction | null = null;
    splitParagraph(d)(placed, (tr) => { captured = tr; });
    const s = placed.apply(trackTransaction(placed, captured!, dan));
    expect(paras(s.doc)).toEqual(['One', 'Two']);
    expect(collectChanges(s.doc).map((c) => c.kind)).toEqual(['mark-ins']);
    expect(settle(d, s.doc, false)).toEqual(['OneTwo']);
    expect(settle(d, s.doc, true)).toEqual(['One', 'Two']);
  });

  it('tracks a deletion across paragraphs and a joined paragraph', async () => {
    const { d, state } = await start('Alpha', 'Beta', 'Gamma');
    const from = at(state.doc, 'pha'), to = at(state.doc, 'mma');
    const s = tracked(state, (st) => st.tr.delete(from, to));
    expect(paras(s.doc)).toEqual(['Alpha', 'Beta', 'Gamma']);
    expect(settle(d, s.doc, true)).toEqual(['Almma']);
    expect(settle(d, s.doc, false)).toEqual(['Alpha', 'Beta', 'Gamma']);
  });

  it('moves the caret the way Backspace and Delete should', async () => {
    const { state } = await start('abcd');
    const placed = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 3)));
    const back = tracked(placed, (st) => st.tr.delete(2, 3));
    expect(back.selection.from).toBe(2);
    const fwd = tracked(placed, (st) => st.tr.delete(3, 4));
    expect(fwd.selection.from).toBe(4);
  });

  it('handles replace-all as separate edits', async () => {
    const { d, state } = await start('cat and cat');
    const s = tracked(state, (st) => {
      const tr = st.tr;
      const second = 'cat and '.length + 1;
      tr.replaceWith(second, second + 3, schema.text('dog'));
      tr.replaceWith(1, 4, schema.text('dog'));
      return tr;
    });
    expect(settle(d, s.doc, true)).toEqual(['dog and dog']);
    expect(settle(d, s.doc, false)).toEqual(['cat and cat']);
    expect(collectChanges(s.doc).length).toBe(4);
  });

  it('accepts one change and leaves the rest', async () => {
    const { d, state } = await start('one two');
    const s = tracked(state, (st) => { const tr = st.tr; tr.insertText('X', 8); tr.insertText('Y', 1); return tr; });
    const list = collectChanges(s.doc);
    expect(list.length).toBe(2);
    const doc = s.apply(applyChanges(s.tr, d, [list[0]], true)).doc;
    expect(collectChanges(doc).map((c) => c.text)).toEqual(['X']);
  });
});

describe('persistence', () => {
  it('writes real revisions that survive a reopen and read correctly elsewhere', async () => {
    const { d, state } = await start('Keep this sentence', 'Second paragraph');
    const from = at(state.doc, 'this');
    let s = tracked(state, (st) => st.tr.insertText('that', from, from + 4));
    const end = at(s.doc, 'Second');
    s = tracked(s, (st) => st.tr.delete(end, end + 7));
    d.track = true;
    const out = await exportDocx(d, s.doc);
    const xml = await xmlOf(out);
    expect(xml).toMatch(/<w:del w:id="\d+" w:author="Dan"[^>]*><w:r><w:delText xml:space="preserve">this<\/w:delText>/);
    expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Dan"[^>]*><w:r><w:t xml:space="preserve">that<\/w:t>/);
    expect(await xmlOf(out, 'word/settings.xml')).toContain('<w:trackRevisions/>');
    const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);

    const back = await WordDoc.open(out);
    expect(back.track).toBe(true);
    expect(collectChanges(back.body).map((c) => [c.kind, c.author, c.text])).toEqual([['del', 'Dan', 'this'], ['ins', 'Dan', 'that'], ['del', 'Dan', 'Second ']]);
    expect(settle(back, back.body, true)).toEqual(['Keep that sentence', 'paragraph']);
    expect(settle(back, back.body, false)).toEqual(['Keep this sentence', 'Second paragraph']);
    const round = await exportDocx(back, back.body);
    expect(await xmlOf(round)).toBe(xml);
  });

  it('turns a rejected deletion back into normal text in the file', async () => {
    const { d, state } = await start('Plain words');
    const s = tracked(state, (st) => st.tr.delete(1, 7));
    const reopened = await WordDoc.open(await exportDocx(d, s.doc));
    const clean = settledCopy(reopened, reopened.body, false);
    const xml = await xmlOf(await exportDocx(reopened, clean));
    expect(xml).not.toMatch(/delText|<w:del /);
    expect(xml).toContain('Plain ');
  });

  it('accepts and rejects tracked formatting from Word', async () => {
    const { d } = await start('x');
    const z = await JSZip.loadAsync(await exportDocx(d, d.body));
    let xml = await z.file('word/document.xml')!.async('string');
    xml = xml.replace('<w:p/>', '<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="5" w:author="Sam" w:date="2026-01-01T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>word</w:t></w:r></w:p>');
    z.file('word/document.xml', xml);
    const w = await WordDoc.open(await z.generateAsync({ type: 'uint8array' }));
    const list = collectChanges(w.body);
    expect(list.map((c) => [c.kind, c.author])).toEqual([['format', 'Sam']]);
    const st = EditorState.create({ doc: w.body, schema });
    const rejected = st.apply(applyChanges(st.tr, w, list, false)).doc;
    const names = (doc: PMNode) => doc.firstChild!.firstChild!.marks.map((m) => m.type.name);
    expect(names(rejected)).toEqual(['italic']);
    const accepted = st.apply(applyChanges(st.tr, w, list, true)).doc;
    expect(names(accepted)).toEqual(['bold']);
  });

  it('stores comments, replies and resolved state in Word\'s parts', async () => {
    const { d, state } = await start('Review this sentence please');
    const c = createComment(d, 'Dan Wood', 'Is this right?', null);
    const reply = createComment(d, 'Sam', 'Yes\nSecond line', c.id);
    const a = anchorNodes(d, c.id), b = anchorNodes(d, reply.id);
    const from = at(state.doc, 'this'), to = from + 4;
    const tr = state.tr.insert(to, [a.end, b.end, a.ref, b.ref]).insert(from, [a.start, b.start]);
    setDone(d, c, true);
    const out = await exportDocx(d, tr.doc);
    const names = Object.keys((await JSZip.loadAsync(out)).files);
    expect(names).toContain('word/comments.xml');
    expect(names).toContain('word/commentsExtended.xml');
    expect(py(`
import sys
from docx import Document
print([(c.author, c.text) for c in Document(sys.argv[1]).comments])`, out)).toContain("('Dan Wood', 'Is this right?')");
    const back = await WordDoc.open(out);
    expect(back.comments.size).toBe(2);
    expect(back.comments.get(reply.id)!.parent).toBe(c.id);
    expect(back.comments.get(reply.id)!.text).toBe('Yes\nSecond line');
    expect(back.comments.get(c.id)!.done).toBe(true);
    // Removing the anchors removes the comment from the file.
    const st = EditorState.create({ doc: back.body, schema });
    const t = st.tr;
    const hits: number[] = [];
    back.body.descendants((n, p) => { if (n.type.name === 'opaque_inline' && /comment/i.test(n.attrs.xml)) hits.push(p); });
    hits.reverse().forEach((p) => t.delete(p, p + 1));
    const gone = await xmlOf(await exportDocx(back, t.doc), 'word/comments.xml');
    expect(gone).not.toContain('Is this right');
  });

  it('leaves existing comments parts untouched when nothing about them changed', async () => {
    const { d, state } = await start('Text');
    const c = createComment(d, 'Dan', 'Note', null);
    const a = anchorNodes(d, c.id);
    const out = await exportDocx(d, state.tr.insert(5, [a.end, a.ref]).insert(1, a.start).doc);
    const z1 = await JSZip.loadAsync(out);
    const before = await z1.file('word/comments.xml')!.async('string');
    const back = await WordDoc.open(out);
    const z2 = await JSZip.loadAsync(await exportDocx(back, back.body));
    expect(await z2.file('word/comments.xml')!.async('string')).toBe(before);
  });
});
