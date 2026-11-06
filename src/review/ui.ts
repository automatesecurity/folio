import { TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { anchorNodes, anchoredIds, Comment, createComment, editComment, setDone } from '../docx/comments';
import type { WordDoc } from '../docx/document';
import { authorColor, schema } from '../editor/schema';
import { applyChange, applyChanges, Change, ChangeKind, collectChanges } from './changes';
import { commentIdOf } from './plugin';

export type ViewMode = 'markup' | 'final' | 'original';

export interface ReviewHost {
  doc(): WordDoc;
  body(): EditorView;
  author(): string;
  ensureAuthor(): Promise<string | null>;
  ask(title: string, label: string, value: string, action: string): Promise<string | null>;
  confirm(title: string, body: string, action: string): Promise<boolean>;
  layout(): void;
  closeHf(): void;
  dirtyChanged(): void;
  say(m: string): void;
}

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const h = (tag: string, cls: string, text?: string) => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};
const button = (label: string, fn: () => void) => {
  const b = h('button', '', label) as HTMLButtonElement;
  b.type = 'button';
  b.addEventListener('click', (e) => { e.stopPropagation(); fn(); });
  return b;
};
const when = (date: string) => {
  const t = Date.parse(date);
  return Number.isNaN(t) ? '' : new Date(t).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
};
const VERB: Record<ChangeKind, string> = {
  ins: 'Inserted', del: 'Deleted', 'mark-ins': 'Inserted', 'mark-del': 'Deleted', 'row-ins': 'Inserted', 'row-del': 'Deleted',
  format: 'Formatted', 'para-format': 'Formatted',
};

interface Thread { root: Comment; replies: Comment[]; from: number; to: number }

export class Review {
  mode: ViewMode = 'markup';
  tracking = false;
  open = false;
  private paneOpen = true;
  private pending = 0;
  private lastMode: ViewMode = 'markup';

  constructor(private host: ReviewHost) {}

  get changes() { return collectChanges(this.host.body().state.doc); }

  threads(): Thread[] {
    const d = this.host.doc(), doc = this.host.body().state.doc;
    const ids = anchoredIds(doc);
    const spans = new Map<string, { from: number; to: number }>();
    doc.descendants((n, pos) => {
      if (n.type !== schema.nodes.opaque_inline || !['range', 'comment'].includes(n.attrs.kind) || !/comment/i.test(n.attrs.xml)) return;
      const id = commentIdOf(n.attrs.xml);
      const s = spans.get(id) ?? { from: pos, to: pos + 1 };
      s.from = Math.min(s.from, pos); s.to = Math.max(s.to, pos + 1);
      spans.set(id, s);
    });
    const list: Thread[] = [];
    for (const c of d.comments.values()) {
      if (c.parent && d.comments.has(c.parent) && ids.has(c.parent)) continue;
      if (!ids.has(c.id)) continue;
      const replies = [...d.comments.values()].filter((r) => r.parent === c.id && ids.has(r.id));
      const span = spans.get(c.id)!;
      list.push({ root: c, replies, from: span.from, to: span.to });
    }
    return list.sort((a, b) => a.from - b.from);
  }

  // Views -------------------------------------------------------------------

  setMode(mode: ViewMode) {
    this.mode = mode;
    if (mode !== 'final' || this.open) this.lastMode = mode;
    document.body.classList.remove('view-markup', 'view-final', 'view-original');
    document.body.classList.add(`view-${mode}`);
    $<HTMLSelectElement>('rv-view').value = mode;
    this.host.body().setProps({ editable: () => mode !== 'original' });
    this.host.layout();
  }

  /** Shows markup and the pane. */
  show() {
    this.host.closeHf();
    this.open = true;
    $('reviewbar').hidden = false;
    $('review-pane').hidden = !this.paneOpen;
    this.setMode(this.lastMode === 'final' ? 'markup' : this.lastMode);
    this.sync();
  }

  /** Back to normal editing: no bar, no pane, no markup on the page. */
  hide() {
    this.open = false;
    $('reviewbar').hidden = true;
    $('review-pane').hidden = true;
    const keep = this.lastMode;
    this.setMode('final');
    this.lastMode = keep;
    this.host.body().focus();
    this.sync();
  }
  toggle() { if (this.open) this.hide(); else this.show(); }

  togglePane() {
    this.paneOpen = !this.paneOpen;
    $('review-pane').hidden = !this.paneOpen;
    this.sync();
  }

  reset() {
    this.tracking = this.host.doc().track;
    this.open = false;
    this.lastMode = 'markup';
    $('reviewbar').hidden = true;
    $('review-pane').hidden = true;
    this.setMode('markup');
    this.sync();
  }

  async toggleTracking() {
    if (!this.tracking && !(await this.host.ensureAuthor())) return;
    this.tracking = !this.tracking;
    this.host.doc().track = this.tracking;
    this.host.say(this.tracking ? 'Track Changes on. Your edits are now recorded.' : 'Track Changes off.');
    this.host.dirtyChanged();
    this.sync();
    this.host.body().focus();
  }

  // Changes ------------------------------------------------------------------

  private commit(fn: (v: EditorView) => void) {
    fn(this.host.body());
    this.host.body().focus();
    this.sync();
  }

  private settle(c: Change, accept: boolean) {
    this.commit((v) => {
      const tr = v.state.tr.setMeta('noTrack', true);
      applyChange(tr, this.host.doc(), c, accept);
      v.dispatch(tr);
    });
  }

  /** Acts on changes inside the selection, or at the caret. */
  settleCurrent(accept: boolean) {
    const v = this.host.body();
    const { from, to } = v.state.selection;
    const all = this.changes;
    const hit = all.filter((c) => c.from <= to && c.to >= from);
    if (!hit.length) { this.go(1); return; }
    this.commit((view) => { view.dispatch(applyChanges(view.state.tr.setMeta('noTrack', true), this.host.doc(), hit, accept)); });
    this.go(1);
  }

  async settleAll(accept: boolean) {
    const all = this.changes;
    if (!all.length) { this.host.say('No tracked changes to review.'); return; }
    const verb = accept ? 'Accept' : 'Reject';
    if (!(await this.host.confirm(`${verb} all ${all.length} changes?`, 'You can undo this with Ctrl+Z.', `${verb} all`))) return;
    this.commit((v) => { v.dispatch(applyChanges(v.state.tr.setMeta('noTrack', true), this.host.doc(), all, accept)); });
    this.host.say(`${accept ? 'Accepted' : 'Rejected'} ${all.length} ${all.length === 1 ? 'change' : 'changes'}.`);
  }

  private select(from: number, to: number) {
    const v = this.host.body();
    const size = v.state.doc.content.size;
    let sel;
    try { sel = TextSelection.create(v.state.doc, Math.min(from, size), Math.min(to, size)); }
    catch { sel = TextSelection.near(v.state.doc.resolve(Math.min(from, size))); }
    v.dispatch(v.state.tr.setSelection(sel).scrollIntoView());
    v.focus();
  }

  go(dir: 1 | -1) {
    const v = this.host.body();
    const stops = [...this.changes.map((c) => ({ from: c.from, to: c.to })), ...this.threads().map((t) => ({ from: t.from, to: t.to }))].sort((a, b) => a.from - b.from);
    if (!stops.length) { this.host.say('No changes or comments to review.'); return; }
    const { from, to } = v.state.selection;
    const next = dir > 0 ? stops.find((s) => s.from > to) ?? stops[0] : [...stops].reverse().find((s) => s.to < from) ?? stops[stops.length - 1];
    this.select(next.from, next.to);
  }

  // Comments -------------------------------------------------------------------

  private wordAt() {
    const v = this.host.body();
    const sel = v.state.selection;
    if (!sel.empty) return { from: sel.from, to: sel.to };
    const $p = sel.$from;
    const text = $p.parent.textBetween(0, $p.parent.content.size, undefined, '￼');
    let a = $p.parentOffset, b = a;
    while (a > 0 && /\w/.test(text[a - 1])) a--;
    while (b < text.length && /\w/.test(text[b])) b++;
    return { from: $p.start() + a, to: $p.start() + b };
  }

  async newComment() {
    this.host.closeHf();
    const author = await this.host.ensureAuthor();
    if (!author) return;
    const range = this.wordAt();
    const text = await this.host.ask('New comment', 'Comment', '', 'Add comment');
    if (!text) return;
    const d = this.host.doc();
    const c = createComment(d, author, text, null);
    const a = anchorNodes(d, c.id);
    this.commit((v) => {
      const tr = v.state.tr.setMeta('noTrack', true);
      tr.insert(range.to, [a.end, a.ref]);
      tr.insert(range.from, a.start);
      v.dispatch(tr);
    });
    this.host.dirtyChanged();
    if (!this.open) this.show();
  }

  private anchorsOf(id: string) {
    const out: { pos: number; size: number; role: string }[] = [];
    this.host.body().state.doc.descendants((n, pos) => {
      if (n.type !== schema.nodes.opaque_inline) return;
      const x = n.attrs.xml as string;
      if (!/comment/i.test(x) || commentIdOf(x) !== id) return;
      out.push({ pos, size: n.nodeSize, role: /RangeStart/.test(x) ? 'start' : /RangeEnd/.test(x) ? 'end' : 'ref' });
    });
    return out;
  }

  private async reply(t: Thread, text: string) {
    const author = await this.host.ensureAuthor();
    if (!author) return;
    const d = this.host.doc();
    const c = createComment(d, author, text, t.root.id);
    const a = anchorNodes(d, c.id);
    const spots = this.anchorsOf(t.root.id);
    const place = (role: string) => spots.find((s) => s.role === role) ?? spots.find((s) => s.role === 'ref')!;
    const items = [
      { pos: place('ref').pos + 1, node: a.ref }, { pos: place('end').pos + 1, node: a.end }, { pos: place('start').pos + 1, node: a.start },
    ].sort((x, y) => y.pos - x.pos);
    this.commit((v) => {
      const tr = v.state.tr.setMeta('noTrack', true);
      for (const it of items) tr.insert(it.pos, it.node);
      v.dispatch(tr);
    });
    this.host.dirtyChanged();
  }

  private async remove(t: Thread) {
    const ok = await this.host.confirm('Delete this comment?', t.replies.length ? 'Its replies are deleted too. You can undo this with Ctrl+Z.' : 'You can undo this with Ctrl+Z.', 'Delete comment');
    if (!ok) return;
    const spots = [t.root, ...t.replies].flatMap((c) => this.anchorsOf(c.id)).sort((a, b) => b.pos - a.pos);
    this.commit((v) => {
      const tr = v.state.tr.setMeta('noTrack', true);
      for (const s of spots) tr.delete(s.pos, s.pos + s.size);
      v.dispatch(tr);
    });
  }

  // Pane ----------------------------------------------------------------------------

  sync() {
    if (this.pending) return;
    this.pending = requestAnimationFrame(() => { this.pending = 0; this.render(); });
  }

  private composer(host: HTMLElement, placeholder: string, value: string, action: string, done: (text: string) => void) {
    const box = h('div', 'rv-compose');
    const area = h('textarea', '') as HTMLTextAreaElement;
    area.placeholder = placeholder;
    area.value = value;
    area.rows = 2;
    area.setAttribute('aria-label', placeholder);
    const submit = () => { const t = area.value.trim(); if (t) { box.remove(); done(t); } };
    area.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); submit(); }
      if (e.key === 'Escape') { e.stopPropagation(); box.remove(); this.sync(); }
    });
    box.append(area, button(action, submit), button('Cancel', () => { box.remove(); this.sync(); }));
    host.append(box);
    area.focus();
  }

  private person(c: { author: string; date: string }) {
    const row = h('div', 'rv-person');
    const dot = h('i', `t-dot t-a${authorColor(c.author)}`);
    row.append(dot, h('strong', '', c.author || 'Unknown'), h('span', '', when(c.date)));
    return row;
  }

  private commentBlock(c: Comment) {
    const box = h('div', 'rv-comment');
    box.append(this.person(c), h('p', 'rv-text', c.text));
    return box;
  }

  private threadCard(t: Thread) {
    const d = this.host.doc();
    const card = h('article', `rv-card rv-thread${t.root.done ? ' is-done' : ''}`);
    card.dataset.pos = String(t.from);
    card.append(this.commentBlock(t.root));
    for (const r of t.replies) card.append(this.commentBlock(r));
    const acts = h('div', 'rv-acts');
    if (t.root.done) {
      card.prepend(h('div', 'rv-flag', 'Resolved'));
      acts.append(button('Reopen', () => { setDone(d, t.root, false); this.host.dirtyChanged(); this.sync(); }));
    } else {
      acts.append(
        button('Reply', () => this.composer(card, 'Write a reply', '', 'Reply', (text) => void this.reply(t, text))),
        button('Edit', () => this.composer(card, 'Edit comment', t.root.text, 'Save', (text) => { editComment(d, t.root, text); this.host.dirtyChanged(); this.sync(); })),
        button('Resolve', () => { setDone(d, t.root, true); this.host.dirtyChanged(); this.sync(); }),
      );
    }
    acts.append(button('Delete', () => void this.remove(t)));
    card.append(acts);
    card.addEventListener('click', () => this.select(t.from, t.to));
    return card;
  }

  private changeCard(c: Change) {
    const card = h('article', 'rv-card');
    card.dataset.pos = String(c.from);
    card.append(this.person(c));
    const line = h('p', 'rv-text');
    const text = c.text.trim();
    line.append(h('span', 'rv-verb', `${VERB[c.kind]} `), h('span', `rv-excerpt rv-${c.kind.endsWith('del') ? 'del' : c.kind.endsWith('ins') ? 'ins' : 'fmt'}`, text.length > 140 ? `${text.slice(0, 140)}...` : text));
    card.append(line);
    const acts = h('div', 'rv-acts');
    acts.append(button('Accept', () => this.settle(c, true)), button('Reject', () => this.settle(c, false)));
    card.append(acts);
    card.addEventListener('click', () => this.select(c.from, Math.max(c.to, c.from)));
    return card;
  }

  private render() {
    const changes = this.changes, threads = this.threads();
    const open = threads.filter((t) => !t.root.done).length;
    const parts = [
      changes.length ? `${changes.length} ${changes.length === 1 ? 'change' : 'changes'}` : '',
      threads.length ? `${threads.length} ${threads.length === 1 ? 'comment' : 'comments'}` : '',
    ].filter(Boolean);
    const pill = $('st-review');
    pill.hidden = !parts.length && !this.tracking;
    pill.textContent = parts.length ? parts.join(', ') : 'Review';
    pill.title = 'Open Review';
    const badge = $('st-track');
    badge.hidden = !this.tracking;
    const tb = $('rv-track');
    tb.setAttribute('aria-pressed', String(this.tracking));
    $('rv-pane').setAttribute('aria-pressed', String(this.paneOpen));
    $('rv-author').textContent = this.host.author() || 'Set your name';
    document.querySelector('[data-id=review]')?.setAttribute('aria-pressed', String(this.open));
    const pane = $('review-pane');
    if (pane.hidden || pane.contains(document.activeElement) && document.activeElement?.tagName === 'TEXTAREA') return;
    const list = $('rv-list');
    const scroll = list.scrollTop;
    const cards: { pos: number; el: HTMLElement }[] = [
      ...threads.map((t) => ({ pos: t.from, el: this.threadCard(t) })),
      ...changes.map((c) => ({ pos: c.from, el: this.changeCard(c) })),
    ].sort((a, b) => a.pos - b.pos);
    $('rv-count').textContent = `${open} open ${open === 1 ? 'comment' : 'comments'}, ${changes.length} ${changes.length === 1 ? 'change' : 'changes'}`;
    if (cards.length) list.replaceChildren(...cards.map((c) => c.el));
    else list.replaceChildren(h('p', 'rv-empty', 'Nothing to review. Turn on Track Changes to record edits, or select text and add a comment.'));
    list.scrollTop = scroll;
  }
}
