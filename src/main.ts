import { baseKeymap, chainCommands, deleteSelection, joinBackward, selectParentNode, toggleMark } from 'prosemirror-commands';
import { dropCursor } from 'prosemirror-dropcursor';
import { gapCursor } from 'prosemirror-gapcursor';
import { history, redo, redoDepth, undo, undoDepth } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { MarkType } from 'prosemirror-model';
import { DOMSerializer } from 'prosemirror-model';
import { Command, EditorState, Plugin, TextSelection, Transaction } from 'prosemirror-state';
import {
  addColumnAfter, addColumnBefore, addRowAfter, addRowBefore, deleteColumn, deleteRow, deleteTable, goToNextCell,
  isInTable, mergeCells, splitCell, tableEditing,
} from 'prosemirror-tables';
import { EditorView } from 'prosemirror-view';
import 'prosemirror-view/style/prosemirror.css';
import { HfKind, HfPart, Margins, WordDoc } from './docx/document';
import { exportDocx } from './docx/export';
import { DocxError } from './docx/package';
import { jcToCss, tw } from './docx/styles';
import { blankDocument } from './docx/template';
import {
  indent, insertHardBreak, insertPageBreak, insertTable, linkRange, selectedParagraphs, setAlign, setMark, setStyle,
  splitParagraph, toggleList, unlistAtStart,
} from './editor/commands';
import { findGo, findKey, findPlugin, findReplace, hiddenAnchorPlugin, Layout, listPlugin, paginatePlugin, requestLayout } from './editor/plugins';
import { clipboardSerializer, schema } from './editor/schema';
import { session } from './editor/session';
import { settledCopy } from './review/changes';
import { reviewPlugin } from './review/plugin';
import { trackTransaction } from './review/track';
import { Review } from './review/ui';
import * as platform from './platform';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

let doc: WordDoc;
/** The editor the toolbar acts on: the body, or a header or footer while one is open. */
let view: EditorView;
let bodyView: EditorView;
let hf: { view: EditorView; part: HfPart; kind: HfKind; page: number } | null = null;
const hfStates = new Map<HfPart, EditorState>();
let savedBody = schema.nodes.doc.create(null, schema.nodes.paragraph.create());
let filePath: string | null = null;
let fileName = 'Untitled.docx';
let isNew = true;
let layout: Layout | null = null;

let savedRev = 0;
let savedTrack = false;
const dirty = () => !bodyView.state.doc.eq(savedBody) || doc.hf.some((p) => !p.doc.eq(p.saved)) || doc.commentRev !== savedRev || doc.track !== savedTrack;

const AUTHOR_KEY = 'folio.author';
let authorName = '';
try { authorName = localStorage.getItem(AUTHOR_KEY) ?? ''; } catch { /* name is asked for again */ }

async function askAuthor(value: string) {
  let name = value;
  const r = await modal(`<h2>Your name</h2><p>Folio puts this on your comments and tracked changes.</p><label for="m-name">Name</label><input type="text" id="m-name" value="${esc(value)}"><div class="acts"><button data-r="cancel">Cancel</button><button class="primary" data-r="ok">Save name</button></div>`, (d) => {
    const i = d.querySelector<HTMLInputElement>('#m-name')!;
    i.focus(); i.select();
    i.addEventListener('input', () => { name = i.value; });
    d.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); d.close('ok'); } });
  });
  if (r !== 'ok' || !name.trim()) return null;
  authorName = name.trim();
  try { localStorage.setItem(AUTHOR_KEY, authorName); } catch { /* lasts for this session */ }
  review.sync();
  return authorName;
}

const review = new Review({
  doc: () => doc,
  body: () => bodyView,
  author: () => authorName,
  ensureAuthor: async () => authorName || askAuthor(''),
  ask: async (title, label, value, action) => {
    let text = value;
    const r = await modal(`<h2>${esc(title)}</h2><label for="m-text">${esc(label)}</label><textarea id="m-text" rows="4">${esc(value)}</textarea><div class="acts"><button data-r="cancel">Cancel</button><button class="primary" data-r="ok">${esc(action)}</button></div>`, (d) => {
      const t = d.querySelector<HTMLTextAreaElement>('#m-text')!;
      t.focus();
      t.addEventListener('input', () => { text = t.value; });
      t.addEventListener('keydown', (e) => { if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); d.close('ok'); } });
    });
    bodyView.focus();
    return r === 'ok' && text.trim() ? text.trim() : null;
  },
  confirm: async (title, body, action) => (await modal(`<h2>${esc(title)}</h2><p>${esc(body)}</p><div class="acts"><button data-r="cancel">Cancel</button><button class="primary" data-r="ok">${esc(action)}</button></div>`)) === 'ok',
  layout: () => requestLayout(),
  closeHf: () => closeHf(),
  dirtyChanged: () => syncTitle(),
  say: (m) => say(m),
});

function reviewer() {
  const date = new Date();
  date.setSeconds(0, 0);
  return { author: authorName || 'Folio user', date: date.toISOString().replace(/\.\d+Z$/, 'Z') };
}
const skipTracking = (tr: Transaction) => !!tr.getMeta('history$') || !!tr.getMeta('noTrack') || tr.getMeta('addToHistory') === false;

// Toolbar ---------------------------------------------------------------------

const icon = (d: string) => `<svg viewBox="0 0 20 20" aria-hidden="true">${d}</svg>`;
const I = {
  open: icon('<path d="M3 15V6h5l2 2h7v7z"/>'),
  save: icon('<path d="M4 4h10l2 2v10H4zM7 4v4h6V4M7 16v-5h6v5"/>'),
  pdf: icon('<path d="M5 3h7l3 3v11H5zM12 3v3h3M8 11l2 2 2-2M10 8v5"/>'),
  undo: icon('<path d="M8 5 4 9l4 4M4 9h8a4 4 0 0 1 0 8h-2"/>'),
  redo: icon('<path d="m12 5 4 4-4 4M16 9H8a4 4 0 0 0 0 8h2"/>'),
  left: icon('<path d="M3 5h14M3 9h9M3 13h14M3 17h9"/>'),
  center: icon('<path d="M3 5h14M6 9h8M3 13h14M6 17h8"/>'),
  right: icon('<path d="M3 5h14M8 9h9M3 13h14M8 17h9"/>'),
  justify: icon('<path d="M3 5h14M3 9h14M3 13h14M3 17h14"/>'),
  bullets: icon('<circle cx="4" cy="6" r="1"/><circle cx="4" cy="14" r="1"/><path d="M8 6h9M8 14h9"/>'),
  numbers: icon('<path d="M3 4h1.5v4M3 8h3M3 12.5c.6-.8 3-.8 3 .6 0 1-3 1.6-3 2.9h3M9 6h8M9 14h8"/>'),
  outdent: icon('<path d="M3 4h14M9 8h8M9 12h8M3 16h14M6 8l-3 2 3 2"/>'),
  indent: icon('<path d="M3 4h14M9 8h8M9 12h8M3 16h14M3 8l3 2-3 2"/>'),
  link: icon('<path d="M8.5 11.5a3 3 0 0 0 4.2 0l3-3a3 3 0 0 0-4.2-4.2l-1 1M11.5 8.5a3 3 0 0 0-4.2 0l-3 3a3 3 0 0 0 4.2 4.2l1-1"/>'),
  image: icon('<path d="M3 4h14v12H3zM3 13l4-4 3 3 3-2 4 4"/><circle cx="13" cy="7.5" r="1"/>'),
  table: icon('<path d="M3 4h14v12H3zM3 8h14M3 12h14M8 4v12M13 4v12"/>'),
  pagebreak: icon('<path d="M5 3v5h10V3M5 17v-4h10v4M3 10.5h2M7 10.5h2M11 10.5h2M15 10.5h2"/>'),
  review: icon('<path d="M3 4h14v10H9l-4 3v-3H3z"/><path d="M7 8h6M7 11h4"/>'),
  print: icon('<path d="M6 7V3h8v4M6 14H3V7h14v7h-3M6 11h8v6H6z"/>'),
  margins: icon('<path d="M4 3h12v14H4zM7 3v14M13 3v14M4 6h12M4 14h12"/>'),
  headfoot: icon('<path d="M4 3h12v14H4zM4 6.5h12M4 13.5h12"/><path d="M7 10h6" stroke-dasharray="1.5 2"/>'),
  find: icon('<circle cx="9" cy="9" r="5"/><path d="m13 13 4 4"/>'),
  pen: icon('<path d="M4 16l1-4 8-8 3 3-8 8z"/>'),
};

interface Btn { id: string; html: string; title: string; run: () => void; active?: () => boolean; cls?: string }
const buttons: Btn[] = [];
const toolbar = $('toolbar');
const run = (c: Command) => () => { c(view.state, view.dispatch, view); view.focus(); };

function btn(host: HTMLElement, b: Btn) {
  const el = document.createElement('button');
  el.type = 'button'; el.innerHTML = b.html; el.title = b.title; el.setAttribute('aria-label', b.title); el.dataset.id = b.id;
  if (b.cls) el.className = b.cls;
  if (b.active) el.setAttribute('aria-pressed', 'false');
  el.addEventListener('mousedown', (e) => e.preventDefault());
  el.addEventListener('click', b.run);
  host.append(el);
  buttons.push(b);
}
const sep = (host: HTMLElement) => { const s = document.createElement('span'); s.className = 'sep'; host.append(s); };
const markActive = (t: MarkType) => () => {
  const { from, to, empty, $from } = view.state.selection;
  if (empty) return !!t.isInSet(view.state.storedMarks ?? $from.marks());
  return view.state.doc.rangeHasMark(from, to, t);
};
function select(id: string, label: string, onChange: (v: string) => void) {
  const s = document.createElement('select');
  s.id = id; s.title = label; s.setAttribute('aria-label', label);
  s.addEventListener('change', () => { onChange(s.value); view.focus(); });
  toolbar.append(s);
  return s;
}
const FONTS = ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Courier New', 'Garamond', 'Tahoma'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
const alignCheck = (jc: string) => () => {
  const p = selectedParagraphs(view.state)[0]?.node;
  if (!p) return false;
  const eff = jcToCss(p.attrs.jc ?? doc.styles.resolvedPara(p.attrs.pStyle).jc);
  return eff === jcToCss(jc);
};

function buildToolbar() {
  const t = toolbar;
  btn(t, { id: 'open', html: I.open, title: 'Open (Ctrl+O)', run: () => void openDocument() });
  btn(t, { id: 'save', html: I.save, title: 'Save (Ctrl+S)', run: () => void save() });
  btn(t, { id: 'pdf', html: I.pdf, title: 'Export as PDF', run: () => void exportToPdf() });
  btn(t, { id: 'print', html: I.print, title: 'Print (Ctrl+P)', run: () => void printDocument() });
  sep(t);
  btn(t, { id: 'undo', html: I.undo, title: 'Undo (Ctrl+Z)', run: run(undo) });
  btn(t, { id: 'redo', html: I.redo, title: 'Redo (Ctrl+Y)', run: run(redo) });
  sep(t);
  styleSel = select('sel-style', 'Paragraph style', (v) => run(setStyle(v === '__default' ? null : v))());
  fontSel = select('sel-font', 'Font', (v) => run(setMark(schema.marks.font, { name: v }))());
  sizeSel = select('sel-size', 'Font size', (v) => run(setMark(schema.marks.size, { val: Number(v) * 2 }))());
  sep(t);
  const mark = (id: string, html: string, title: string, m: MarkType) => btn(t, { id, html, title, run: run(toggleMark(m)), active: markActive(m) });
  mark('b', '<b>B</b>', 'Bold (Ctrl+B)', schema.marks.bold);
  mark('i', '<i>I</i>', 'Italic (Ctrl+I)', schema.marks.italic);
  mark('u', '<u>U</u>', 'Underline (Ctrl+U)', schema.marks.underline);
  mark('s', '<s>S</s>', 'Strikethrough', schema.marks.strike);
  const vert = (val: string) => () => view.state.selection.$from.marks().some((m) => m.type === schema.marks.vert && m.attrs.val === val) || (view.state.storedMarks ?? []).some((m) => m.attrs.val === val);
  btn(t, { id: 'sup', html: 'x<sup style="font-size:9px">2</sup>', title: 'Superscript', run: () => run(setMark(schema.marks.vert, vert('superscript')() ? null : { val: 'superscript' }))(), active: vert('superscript') });
  btn(t, { id: 'sub', html: 'x<sub style="font-size:9px">2</sub>', title: 'Subscript', run: () => run(setMark(schema.marks.vert, vert('subscript')() ? null : { val: 'subscript' }))(), active: vert('subscript') });
  btn(t, { id: 'color', html: `${'<span>A</span>'}<i style="--c:#c00000"></i>`, cls: 'swatch', title: 'Text color', run: pickColor });
  btn(t, { id: 'hl', html: `${I.pen}<i style="--c:#ffff00"></i>`, cls: 'swatch', title: 'Highlight the selection', run: toggleHighlight, active: markActive(schema.marks.highlight) });
  sep(t);
  for (const [jc, label, ic] of [['left', 'Align left', I.left], ['center', 'Center', I.center], ['right', 'Align right', I.right], ['both', 'Justify', I.justify]] as const)
    btn(t, { id: `a-${jc}`, html: ic, title: label, run: () => run(setAlign(jc))(), active: alignCheck(jc) });
  sep(t);
  const listOn = (k: 'bullet' | 'ordered') => () => doc.numbering.kindOf(selectedParagraphs(view.state)[0]?.node.attrs.numId) === k;
  btn(t, { id: 'ul', html: I.bullets, title: 'Bulleted list', run: () => run(toggleList(doc, 'bullet'))(), active: listOn('bullet') });
  btn(t, { id: 'ol', html: I.numbers, title: 'Numbered list', run: () => run(toggleList(doc, 'ordered'))(), active: listOn('ordered') });
  btn(t, { id: 'out', html: I.outdent, title: 'Decrease indent', run: run(indent(-1)) });
  btn(t, { id: 'in', html: I.indent, title: 'Increase indent', run: run(indent(1)) });
  sep(t);
  btn(t, { id: 'link', html: I.link, title: 'Insert link (Ctrl+K)', run: () => void editLink() });
  btn(t, { id: 'img', html: I.image, title: 'Insert image', run: () => void chooseImage() });
  btn(t, { id: 'tbl', html: I.table, title: 'Insert table', run: () => void chooseTable() });
  btn(t, { id: 'pb', html: I.pagebreak, title: 'Insert page break (Ctrl+Enter)', run: run(insertPageBreak) });
  sep(t);
  btn(t, { id: 'hf', html: I.headfoot, title: 'Edit header and footer', run: () => openHf('header', currentPage()) });
  btn(t, { id: 'margins', html: I.margins, title: 'Page margins', run: () => void chooseMargins() });
  sep(t);
  btn(t, { id: 'review', html: I.review, title: 'Review: track changes and comments', run: () => review.toggle(), active: () => review.open });
  btn(t, { id: 'find', html: I.find, title: 'Find and replace (Ctrl+F)', run: openFind });

  const tb = $('tablebar');
  const tableBtn = (label: string, c: Command) => {
    const e = document.createElement('button');
    e.type = 'button'; e.textContent = label;
    e.addEventListener('mousedown', (ev) => ev.preventDefault());
    e.addEventListener('click', run(c));
    tb.append(e);
  };
  tableBtn('Row above', addRowBefore); tableBtn('Row below', addRowAfter); tableBtn('Delete row', deleteRow); sep(tb);
  tableBtn('Column left', addColumnBefore); tableBtn('Column right', addColumnAfter); tableBtn('Delete column', deleteColumn); sep(tb);
  tableBtn('Merge cells', mergeCells); tableBtn('Split cell', splitCell); sep(tb);
  tableBtn('Delete table', deleteTable);
}
let styleSel: HTMLSelectElement, fontSel: HTMLSelectElement, sizeSel: HTMLSelectElement;
let currentColor = 'C00000';

function pickColor() {
  const input = $<HTMLInputElement>('color-input');
  input.value = `#${currentColor}`;
  input.click();
}
function toggleHighlight() {
  const on = markActive(schema.marks.highlight)();
  run(setMark(schema.marks.highlight, on ? null : { val: 'yellow' }))();
}

function fillStyleSelect() {
  const has = doc.styles.paragraphChoices();
  styleSel.replaceChildren(...has.map((s) => new Option(s.name.replace(/^heading/, 'Heading'), s.id)));
  if (!doc.styles.defaultParagraph) styleSel.prepend(new Option('Normal', '__default'));
}
function setOptions(sel: HTMLSelectElement, values: string[], value: string) {
  if (![...sel.options].some((o) => o.value === value) || sel.options.length !== new Set([...values, value]).size) {
    sel.replaceChildren(...[...new Set([...values, value])].filter(Boolean).map((v) => new Option(v, v)));
  }
  sel.value = value;
}

function syncToolbar() {
  const s = view.state;
  for (const b of buttons) {
    if (!b.active) continue;
    toolbar.querySelector(`[data-id="${b.id}"]`)?.setAttribute('aria-pressed', String(b.active()));
  }
  (toolbar.querySelector('[data-id=undo]') as HTMLButtonElement).disabled = undoDepth(s) === 0;
  (toolbar.querySelector('[data-id=redo]') as HTMLButtonElement).disabled = redoDepth(s) === 0;
  const para = selectedParagraphs(s)[0]?.node;
  const sid = para?.attrs.pStyle ?? doc.styles.defaultParagraph ?? '__default';
  if ([...styleSel.options].some((o) => o.value === sid)) styleSel.value = sid;
  else {
    const def = doc.styles.styles.get(sid);
    if (def) { styleSel.append(new Option(def.name.charAt(0).toUpperCase() + def.name.slice(1), def.id)); styleSel.value = sid; }
  }
  const marks = s.storedMarks ?? s.selection.$from.marks();
  const base = doc.styles.resolvedRun(para?.attrs.pStyle);
  const font = (schema.marks.font.isInSet(marks)?.attrs.name as string | undefined) ?? base.font ?? 'Calibri';
  const size = (schema.marks.size.isInSet(marks)?.attrs.val as number | undefined) ?? base.size ?? 22;
  setOptions(fontSel, FONTS, font);
  setOptions(sizeSel, SIZES.map(String), String(size / 2));
  $('tablebar').hidden = !isInTable(s);
}

// Layout -----------------------------------------------------------------------

function applyPageGeometry() {
  const p = doc.page;
  $('zoomer').style.width = `${tw(p.width)}px`;
  $('sheet').style.padding = `${tw(doc.effTop)}px ${tw(p.right)}px 0 ${tw(p.left)}px`;
}

/** Keeps doc.page in step with the document, including after undo. */
function syncPageSettings() {
  const m = bodyView.state.doc.attrs.pgMar as Margins | null;
  if (!m) return;
  const p = doc.page;
  if ((Object.keys(m) as (keyof Margins)[]).every((k) => m[k] === p[k])) return;
  Object.assign(p, m);
  applyPageGeometry();
  requestLayout();
}

/** Page numbers are live fields, so each page needs its own copy of the header and footer. */
function hfContent(part: HfPart, page: number, pages: number) {
  const box = document.createElement('div');
  box.append(DOMSerializer.fromSchema(schema).serializeFragment(part.doc.content));
  box.querySelectorAll<HTMLElement>('[data-field=PAGE]').forEach((e) => { e.textContent = String(page); });
  box.querySelectorAll<HTMLElement>('[data-field=NUMPAGES]').forEach((e) => { e.textContent = String(pages); });
  return box;
}
function hfBand(kind: HfKind, index: number, pages: number) {
  const part = doc.findPart(kind, doc.slotFor(index, bodyView.state.doc.attrs.titlePg === true));
  if (!part || (hf?.part === part && hf.page === index)) return null;
  const box = hfContent(part, index + 1, pages);
  box.className = `hf pm-doc hf-${kind}`;
  return box;
}

function measureBands() {
  const p = doc.page;
  let top = 0, bottom = 0;
  const reach = (el: HTMLElement, header: boolean) => {
    const h = el.offsetHeight * 15;
    if (header) top = Math.max(top, p.header + h); else bottom = Math.max(bottom, p.footer + h);
  };
  $('pages').querySelectorAll<HTMLElement>('.hf').forEach((el) => reach(el, el.classList.contains('hf-header')));
  if (hf) reach($('hf-editor'), hf.kind === 'header');
  if (top === doc.hfTop && bottom === doc.hfBottom) return;
  doc.hfTop = top; doc.hfBottom = bottom;
  applyPageGeometry();
  requestLayout();
}

function renderPages(l: Layout) {
  layout = l;
  const host = $('pages');
  const p = doc.page;
  const n = l.pages.length;
  const total = l.pages[n - 1].top + l.pages[n - 1].height;
  $('sheet').style.minHeight = `${total}px`;
  host.style.height = `${total}px`;
  const frag = document.createDocumentFragment();
  l.pages.forEach((pg, i) => {
    const el = document.createElement('div');
    el.className = 'page';
    el.dataset.i = String(i);
    el.style.top = `${pg.top}px`; el.style.height = `${pg.height}px`;
    el.style.setProperty('--ml', `${tw(p.left)}px`); el.style.setProperty('--mr', `${tw(p.right)}px`);
    for (const kind of ['header', 'footer'] as const) {
      const box = hfBand(kind, i, n);
      if (!box) continue;
      box.style[kind === 'header' ? 'top' : 'bottom'] = `${tw(kind === 'header' ? p.header : p.footer)}px`;
      el.append(box);
    }
    frag.append(el);
  });
  host.replaceChildren(frag);
  placeHf();
  measureBands();
  syncStatus();
}

function currentPage() {
  if (!layout) return 0;
  let block = 0;
  try { block = bodyView.state.doc.resolve(bodyView.state.selection.from).index(0); } catch { /* selection outside blocks */ }
  return layout.blockPage[block] ?? 0;
}

function syncStatus() {
  const s = bodyView.state;
  let text = '';
  s.doc.descendants((n) => { if (n.isText) text += n.text; else if (n.isBlock) text += ' '; return true; });
  const words = text.trim() ? text.trim().split(/\s+/).length : 0;
  $('st-words').textContent = `${words.toLocaleString()} ${words === 1 ? 'word' : 'words'}`;
  if (layout) $('st-pages').textContent = `Page ${currentPage() + 1} of ${layout.pages.length}`;
}
const say = (m: string) => { $('st-message').textContent = m; };

/** What the file holds that Folio shows but cannot edit. */
function preservedSummary() {
  const c: Record<string, number> = {};
  const add = (k: string) => { c[k] = (c[k] ?? 0) + 1; };
  bodyView.state.doc.descendants((n) => {
    if (n.type === schema.nodes.opaque_inline) {
      const k = n.attrs.kind;
      if (k === 'ins' || k === 'del') add('Tracked moves and complex changes');
      else if (k === 'note') add('Footnotes and endnotes');
      else if (k === 'field') add('Fields');
      else if (k === 'math') add('Equations');
      else if (k === 'drawing' && !n.attrs.img) add('Shapes and graphics');
      else if (k === 'other' || k === 'sdt') add('Other inline objects');
    } else if (n.type === schema.nodes.opaque_block && n.attrs.kind !== 'bookmark') {
      add(n.attrs.kind === 'table' ? 'Complex tables' : n.attrs.kind === 'sdt' ? 'Content controls' : 'Other blocks');
    }
    return true;
  });
  return c;
}
function syncPreserved() {
  const c = preservedSummary();
  const total = Object.values(c).reduce((a, b) => a + b, 0);
  const b = $('st-preserved');
  b.hidden = !total;
  b.textContent = `${total} read-only ${total === 1 ? 'item' : 'items'}`;
}

// Dialogs ----------------------------------------------------------------------

function modal(html: string, wire?: (d: HTMLDialogElement) => void): Promise<string> {
  const d = $<HTMLDialogElement>('modal');
  d.innerHTML = html;
  return new Promise((resolve) => {
    d.querySelectorAll<HTMLButtonElement>('[data-r]').forEach((b) => b.addEventListener('click', () => { d.close(b.dataset.r); }));
    d.addEventListener('close', () => resolve(d.returnValue || 'cancel'), { once: true });
    d.returnValue = '';
    d.showModal();
    wire?.(d);
  });
}
const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]!));
async function showError(title: string, e: unknown) {
  await modal(`<h2>${esc(title)}</h2><p>${esc(e instanceof Error || e instanceof DocxError ? (e as Error).message : String(e))}</p><div class="acts"><button class="primary" data-r="ok">OK</button></div>`);
}
async function confirmDiscard(): Promise<boolean> {
  if (!dirty()) return true;
  const r = await modal(`<h2>Save changes to ${esc(fileName)}?</h2><p>Your changes will be lost if you don't save them.</p><div class="acts"><button data-r="cancel">Cancel</button><button data-r="discard">Don't save</button><button class="primary" data-r="save">Save</button></div>`);
  if (r === 'discard') return true;
  if (r === 'save') return save();
  return false;
}
async function editLink() {
  const r = linkRange(view.state);
  let value = '';
  const hasText = r.from !== r.to;
  const res = await modal(`<h2>${r.href ? 'Edit link' : 'Insert link'}</h2><label for="m-url">Web or email address</label><input type="text" id="m-url" value="${esc(r.href)}" placeholder="https://example.com">${hasText ? '' : '<label for="m-text">Text to show</label><input type="text" id="m-text">'}<div class="acts">${r.href ? '<button data-r="remove">Remove link</button>' : ''}<button data-r="cancel">Cancel</button><button class="primary" data-r="ok">Apply</button></div>`, (d) => {
    const u = d.querySelector<HTMLInputElement>('#m-url')!;
    u.focus(); u.addEventListener('input', () => { value = u.value; }); value = u.value;
    d.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); d.close('ok'); } });
    const t = d.querySelector<HTMLInputElement>('#m-text');
    t?.addEventListener('input', () => { linkText = t.value; });
  });
  const text = linkText; linkText = '';
  view.focus();
  if (res === 'remove') { view.dispatch(view.state.tr.removeMark(r.from, r.to, schema.marks.link)); return; }
  if (res !== 'ok' || !value.trim()) return;
  let href = value.trim();
  if (!/^[a-z][a-z0-9+.-]*:/i.test(href)) href = href.includes('@') ? `mailto:${href}` : `https://${href}`;
  if (!/^(https?:|mailto:)/i.test(href)) { await showError('Link not added', new Error('Links must start with http, https or mailto.')); return; }
  const mark = schema.marks.link.create({ href });
  const tr = view.state.tr;
  if (hasText) tr.removeMark(r.from, r.to, schema.marks.link).addMark(r.from, r.to, mark);
  else tr.replaceSelectionWith(schema.text(text || href, [mark]), false);
  view.dispatch(tr);
}
let linkText = '';
async function chooseTable() {
  let rows = 3, cols = 3;
  const r = await modal(`<h2>Insert table</h2><div class="row"><div><label for="m-c">Columns</label><input type="number" id="m-c" min="1" max="20" value="3"></div><div><label for="m-r">Rows</label><input type="number" id="m-r" min="1" max="200" value="3"></div></div><div class="acts"><button data-r="cancel">Cancel</button><button class="primary" data-r="ok">Insert table</button></div>`, (d) => {
    const c = d.querySelector<HTMLInputElement>('#m-c')!, w = d.querySelector<HTMLInputElement>('#m-r')!;
    c.addEventListener('input', () => { cols = Number(c.value); }); w.addEventListener('input', () => { rows = Number(w.value); });
  });
  view.focus();
  if (r === 'ok' && rows >= 1 && cols >= 1) run(insertTable(doc, Math.min(rows, 200), Math.min(cols, 20)))();
}
const EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/bmp': 'bmp' };
function sniffExt(b: Uint8Array) {
  if (b[0] === 0x89 && b[1] === 0x50) return 'png';
  if (b[0] === 0xff && b[1] === 0xd8) return 'jpeg';
  if (b[0] === 0x47 && b[1] === 0x49) return 'gif';
  if (b[0] === 0x42 && b[1] === 0x4d) return 'bmp';
  if (b[8] === 0x57 && b[9] === 0x45) return 'webp';
  return null;
}
async function insertImage(bytes: Uint8Array, name: string, pos?: number) {
  const ext = sniffExt(bytes);
  if (!ext) { await showError('Image not added', new Error('Folio can insert PNG, JPEG, GIF, BMP and WebP images.')); return; }
  const part = hf?.part;
  const { rId, key } = await doc.addImage(bytes, ext, part);
  const img = new Image();
  img.src = session.assets.get(key)!;
  await img.decode().catch(() => undefined);
  const max = tw(doc.page.width - doc.page.left - doc.page.right);
  const scale = Math.min(1, max / (img.naturalWidth || max));
  const w = Math.round((img.naturalWidth || 300) * scale), h = Math.round((img.naturalHeight || 200) * scale);
  const node = schema.nodes.opaque_inline.create({ kind: 'drawing', xml: doc.drawingXml(rId, w, h, Date.now() % 1000000, name), img: key, w, h, label: name, session: session.id });
  const tr = view.state.tr;
  if (pos !== undefined) tr.insert(pos, node); else tr.replaceSelectionWith(node, false);
  view.dispatch(tr.scrollIntoView());
}
const PRESETS: Record<string, Partial<Margins>> = {
  Normal: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
  Narrow: { top: 720, bottom: 720, left: 720, right: 720 },
  Moderate: { top: 1440, bottom: 1440, left: 1080, right: 1080 },
  Wide: { top: 1440, bottom: 1440, left: 2880, right: 2880 },
};
async function chooseMargins() {
  const m = bodyView.state.doc.attrs.pgMar as Margins;
  const inches = (v: number) => String(Math.round((v / 1440) * 1000) / 1000);
  const field = (k: keyof Margins, label: string) => `<div><label for="m-${k}">${label}</label><input type="number" id="m-${k}" min="0" step="0.05" value="${inches(m[k])}"></div>`;
  const preset = Object.keys(PRESETS).find((k) => (Object.keys(PRESETS[k]) as (keyof Margins)[]).every((x) => PRESETS[k][x] === m[x])) ?? 'Custom';
  const ok = await modal(`<h2>Margins</h2><label for="m-preset">Preset</label><select id="m-preset">${[...Object.keys(PRESETS), 'Custom'].map((k) => `<option${k === preset ? ' selected' : ''}>${k}</option>`).join('')}</select><div class="row">${field('top', 'Top (in)')}${field('bottom', 'Bottom (in)')}</div><div class="row">${field('left', 'Left (in)')}${field('right', 'Right (in)')}</div><div class="row">${field('header', 'Header from edge (in)')}${field('footer', 'Footer from edge (in)')}</div><p class="hint">Applies to the whole document.</p><div class="acts"><button data-r="cancel">Cancel</button><button class="primary" data-r="ok">Apply margins</button></div>`, (d) => {
    const sel = d.querySelector<HTMLSelectElement>('#m-preset')!;
    const input = (k: string) => d.querySelector<HTMLInputElement>(`#m-${k}`)!;
    sel.addEventListener('change', () => {
      for (const [k, v] of Object.entries(PRESETS[sel.value] ?? {})) input(k).value = inches(v as number);
    });
    for (const k of ['top', 'bottom', 'left', 'right']) input(k).addEventListener('input', () => { sel.value = 'Custom'; });
    d.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); d.close('ok'); } });
  });
  const d = $<HTMLDialogElement>('modal');
  const next = Object.fromEntries((Object.keys(m) as (keyof Margins)[]).map((k) => [k, Math.round(Number(d.querySelector<HTMLInputElement>(`#m-${k}`)!.value) * 1440)])) as unknown as Margins;
  if (ok !== 'ok') { view.focus(); return; }
  const p = doc.page;
  const bad = Object.values(next).some((v) => !Number.isFinite(v) || v < 0);
  if (bad || p.width - next.left - next.right < 720 || p.height - next.top - next.bottom < 720) {
    await showError('Margins not changed', new Error('Margins must be positive and leave at least half an inch of text area each way.'));
    return;
  }
  bodyView.dispatch(bodyView.state.tr.setDocAttribute('pgMar', next));
  view.focus();
}

async function chooseImage() {
  try {
    const f = await platform.pickImage();
    if (f) await insertImage(f.bytes, f.name);
  } catch (e) { await showError('Image not added', e); }
  view.focus();
}

// Find bar ---------------------------------------------------------------------

function openFind() {
  closeHf();
  $('findbar').hidden = false;
  const i = $<HTMLInputElement>('find-input');
  const { from, to, empty } = bodyView.state.selection;
  if (!empty) { i.value = bodyView.state.doc.textBetween(from, to, ' '); runFind(); }
  i.focus(); i.select();
}
function closeFind() {
  $('findbar').hidden = true;
  bodyView.dispatch(bodyView.state.tr.setMeta(findKey, { query: '' }));
  bodyView.focus();
}
function runFind() {
  bodyView.dispatch(bodyView.state.tr.setMeta(findKey, { query: $<HTMLInputElement>('find-input').value, matchCase: $<HTMLInputElement>('find-case').checked, index: 0 }));
  syncFindCount();
  const f = findKey.getState(bodyView.state)!;
  if (f.matches.length) { findGo(bodyView, 1); findGo(bodyView, -1); }
}
function syncFindCount() {
  const f = findKey.getState(bodyView.state)!;
  $('find-count').textContent = f.query ? (f.matches.length ? `${f.index + 1} of ${f.matches.length}` : 'No matches') : '';
}
function wireFind() {
  $('find-input').addEventListener('input', runFind);
  $('find-case').addEventListener('change', runFind);
  $('findbar').addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeFind();
    if (e.key === 'Enter' && e.target === $('find-input')) { e.preventDefault(); findGo(bodyView, e.shiftKey ? -1 : 1); syncFindCount(); }
  });
  $('findbar').addEventListener('click', (e) => {
    const act = (e.target as HTMLElement).dataset.find;
    if (!act) return;
    const rep = $<HTMLInputElement>('replace-input').value;
    if (act === 'close') closeFind();
    if (act === 'next') findGo(bodyView, 1);
    if (act === 'prev') findGo(bodyView, -1);
    if (act === 'replace') { findReplace(bodyView, rep, false); }
    if (act === 'all') { const n = findReplace(bodyView, rep, true); say(`Replaced ${n} ${n === 1 ? 'match' : 'matches'}`); }
    syncFindCount();
  });
}

// Document lifecycle --------------------------------------------------------------

function appliedStyles() {
  let el = document.getElementById('doc-styles');
  if (!el) { el = document.createElement('style'); el.id = 'doc-styles'; document.head.append(el); }
  el.textContent = doc.styles.css();
}

function plugins(isBody: boolean) {
  const enter = () => splitParagraph(doc);
  return [
    hiddenAnchorPlugin(),
    history(),
    keymap({
      'Mod-z': undo, 'Mod-y': redo, 'Mod-Shift-z': redo,
      'Mod-b': toggleMark(schema.marks.bold), 'Mod-i': toggleMark(schema.marks.italic), 'Mod-u': toggleMark(schema.marks.underline),
      Enter: (s, d, v) => enter()(s, d, v),
      'Shift-Enter': insertHardBreak, 'Mod-Enter': insertPageBreak,
      Backspace: chainCommands(deleteSelection, unlistAtStart, joinBackward),
      Tab: (s, d, v) => (isInTable(s) ? goToNextCell(1)(s, d) : selectedParagraphs(s).some((p) => p.node.attrs.numId) ? indent(1)(s, d, v) : d ? (d(s.tr.insertText('\t').scrollIntoView()), true) : true),
      'Shift-Tab': (s, d, v) => (isInTable(s) ? goToNextCell(-1)(s, d) : indent(-1)(s, d, v)),
      Escape: selectParentNode,
    }),
    keymap(baseKeymap),
    dropCursor({ color: '#a67c2d' }), gapCursor(), tableEditing(),
    listPlugin(() => doc),
    reviewPlugin(),
    ...(isBody ? [paginatePlugin(renderPages, () => doc), findPlugin()] : []),
    new Plugin({
      props: {
        handleClick(v, pos, e) {
          if (!(e.ctrlKey || e.metaKey)) return false;
          const a = (e.target as HTMLElement).closest('a');
          const href = a?.getAttribute('data-href');
          if (href && /^(https?:|mailto:)/i.test(href)) { platform.openExternal(href); return true; }
          void pos; void v;
          return false;
        },
        handlePaste(v, e) {
          const file = [...(e.clipboardData?.files ?? [])].find((f) => EXT[f.type]);
          if (!file) return false;
          void file.arrayBuffer().then((b) => insertImage(new Uint8Array(b), file.name || 'Pasted image'));
          void v;
          return true;
        },
        handleDrop(v, e) {
          const files = [...((e as DragEvent).dataTransfer?.files ?? [])];
          const docx = files.find((f) => /\.(docx|docm|dotx)$/i.test(f.name));
          if (docx) {
            void docx.arrayBuffer().then((b) => openBytes({ path: null, name: docx.name, bytes: new Uint8Array(b) }));
            return true;
          }
          const img = files.find((f) => EXT[f.type]);
          if (!img) return false;
          const at = v.posAtCoords({ left: (e as DragEvent).clientX, top: (e as DragEvent).clientY })?.pos;
          void img.arrayBuffer().then((b) => insertImage(new Uint8Array(b), img.name, at));
          return true;
        },
      },
    }),
  ];
}

function makeView(host: HTMLElement, state: EditorState, after: (v: EditorView, tr: Transaction) => void, transform?: (v: EditorView, tr: Transaction) => Transaction) {
  const v: EditorView = new EditorView(host, {
    state,
    clipboardSerializer: clipboardSerializer(),
    attributes: { class: 'pm-doc', spellcheck: 'true', 'aria-multiline': 'true', role: 'textbox' },
    dispatchTransaction(incoming) {
      const tr = transform ? transform(v, incoming) : incoming;
      v.updateState(v.state.apply(tr));
      after(v, tr);
      syncToolbar();
      if (tr.docChanged) syncTitle();
    },
  });
  return v;
}

function mount(d: WordDoc) {
  closeHf();
  doc = d;
  hfStates.clear();
  appliedStyles();
  applyPageGeometry();
  fillStyleSelect();
  savedBody = d.body;
  savedRev = d.commentRev;
  savedTrack = d.track;
  layout = null;
  if (bodyView) bodyView.destroy();
  bodyView = makeView($('editor'), EditorState.create({ doc: d.body, plugins: plugins(true) }), (_, tr) => {
    if (tr.docChanged) { syncPreserved(); syncPageSettings(); review.sync(); }
    if (tr.docChanged || tr.selectionSet) syncStatus();
    syncFindCount();
  }, (v, tr) => (review.tracking && !skipTracking(tr) ? trackTransaction(v.state, tr, reviewer()) : tr));
  bodyView.dom.setAttribute('aria-label', 'Document');
  view = bodyView;
  review.reset();
  syncToolbar(); syncPreserved(); syncTitle(); syncStatus();
  view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(0))));
  view.focus();
  $('desk').scrollTop = 0;
}

// Headers and footers ------------------------------------------------------------

const SLOT_LABEL = { default: '', first: 'First page ', even: 'Even page ' };

function placeHf() {
  if (!hf || !layout) return;
  const el = $('hf-editor'), p = doc.page;
  const pg = layout.pages[Math.min(hf.page, layout.pages.length - 1)];
  el.style.left = `${tw(p.left)}px`;
  el.style.width = `${tw(p.width - p.left - p.right)}px`;
  el.style.top = `${hf.kind === 'header' ? pg.top + tw(p.header) : pg.top + pg.height - tw(p.footer) - el.offsetHeight}px`;
}

function openHf(kind: HfKind, pageIndex: number) {
  closeHf();
  const titlePg = bodyView.state.doc.attrs.titlePg === true;
  const part = doc.hfPart(kind, doc.slotFor(pageIndex, titlePg));
  const state = hfStates.get(part) ?? EditorState.create({ doc: part.doc, plugins: plugins(false) });
  const v = makeView($('hf-editor'), state, (ed, tr) => {
    hfStates.set(part, ed.state);
    part.doc = ed.state.doc;
    if (tr.docChanged && layout) renderPages(layout);
  });
  v.dom.setAttribute('aria-label', `${kind === 'header' ? 'Header' : 'Footer'} text`);
  hf = { view: v, part, kind, page: pageIndex };
  view = v;
  $('hfbar').hidden = false; $('hf-editor').hidden = false;
  document.body.classList.add('hf-editing');
  $('hf-label').textContent = `${SLOT_LABEL[part.slot]}${SLOT_LABEL[part.slot] ? kind : kind === 'header' ? 'Header' : 'Footer'}`;
  $<HTMLInputElement>('hf-first').checked = titlePg;
  $('hf-switch').textContent = kind === 'header' ? 'Go to footer' : 'Go to header';
  if (layout) renderPages(layout);
  v.focus();
  v.dispatch(v.state.tr.setSelection(TextSelection.near(v.state.doc.resolve(v.state.doc.content.size), -1)));
  syncToolbar();
  $('hf-editor').scrollIntoView({ block: 'nearest' });
}

function closeHf() {
  if (!hf) return;
  hf.view.destroy();
  hf = null;
  view = bodyView;
  $('hfbar').hidden = true; $('hf-editor').hidden = true;
  document.body.classList.remove('hf-editing');
  if (layout) renderPages(layout);
  view.focus();
  syncToolbar();
}

function insertField(instr: 'PAGE' | 'NUMPAGES') {
  if (!hf) return;
  const w = hf.part.w;
  const label = instr === 'PAGE' ? '1' : String(layout?.pages.length ?? 1);
  const xml = `<${w}fldSimple ${w}instr=" ${instr} "><${w}r><${w}t>${label}</${w}t></${w}r></${w}fldSimple>`;
  const node = schema.nodes.opaque_inline.create({ kind: 'field', xml, label, instr, session: session.id });
  view.dispatch(view.state.tr.replaceSelectionWith(node, false).scrollIntoView());
  view.focus();
}

function wireHf() {
  $('hf-page').addEventListener('click', () => insertField('PAGE'));
  $('hf-pages').addEventListener('click', () => insertField('NUMPAGES'));
  $('hf-close').addEventListener('click', closeHf);
  $('hf-switch').addEventListener('click', () => { if (hf) openHf(hf.kind === 'header' ? 'footer' : 'header', hf.page); });
  $<HTMLInputElement>('hf-first').addEventListener('change', (e) => {
    const on = (e.target as HTMLInputElement).checked;
    const at = hf;
    bodyView.dispatch(bodyView.state.tr.setDocAttribute('titlePg', on));
    if (at) openHf(at.kind, at.page);
  });
  // Double-clicking the top or bottom margin of a page opens that header or footer.
  $('desk').addEventListener('dblclick', (e) => {
    const pageEl = document.elementsFromPoint(e.clientX, e.clientY).find((x) => x.classList.contains('page')) as HTMLElement | undefined;
    if (!pageEl || (e.target as HTMLElement).closest('#hf-editor')) return;
    const r = pageEl.getBoundingClientRect();
    const y = ((e.clientY - r.top) / r.height) * pageEl.offsetHeight;
    const index = Number(pageEl.dataset.i);
    if (y < tw(doc.effTop)) { e.preventDefault(); openHf('header', index); }
    else if (y > pageEl.offsetHeight - tw(doc.effBottom)) { e.preventDefault(); openHf('footer', index); }
    else closeHf();
  });
}

// Printing --------------------------------------------------------------------------

/** Lays the on-screen pages out as real sheets, so print matches what's on screen. */
function buildPrintPages() {
  if (!layout) return;
  const p = doc.page, root = $('print-root');
  const n = layout.pages.length;
  const sheets = layout.pages.map((pg, i) => {
    const sheet = document.createElement('div');
    sheet.className = 'print-page';
    sheet.style.minHeight = `${tw(p.height) - 1}px`;
    sheet.style.width = `${tw(p.width)}px`;
    sheet.style.padding = `${tw(doc.effTop)}px ${tw(p.right)}px ${pg.height > tw(p.height) ? tw(doc.effBottom) : 0}px ${tw(p.left)}px`;
    const body = document.createElement('div');
    body.className = 'pm-doc';
    sheet.append(body);
    for (const kind of ['header', 'footer'] as const) {
      const part = doc.findPart(kind, doc.slotFor(i, bodyView.state.doc.attrs.titlePg === true));
      if (!part) continue;
      const box = hfContent(part, i + 1, n);
      box.className = `hf pm-doc hf-${kind}`;
      box.style[kind === 'header' ? 'top' : 'bottom'] = `${tw(kind === 'header' ? p.header : p.footer)}px`;
      box.style.left = `${tw(p.left)}px`; box.style.right = `${tw(p.right)}px`;
      sheet.append(box);
    }
    return { sheet, body };
  });
  const kids = bodyView.dom.children;
  layout.blockPage.forEach((page, i) => {
    const el = kids[i]?.cloneNode(true) as HTMLElement | undefined;
    if (!el) return;
    el.style.marginTop = '';
    sheets[page].body.append(el);
  });
  root.replaceChildren(...sheets.map((x) => x.sheet));
  let rule = document.getElementById('print-rule');
  if (!rule) { rule = document.createElement('style'); rule.id = 'print-rule'; document.head.append(rule); }
  rule.textContent = `@page { size: ${p.width / 1440}in ${p.height / 1440}in; margin: 0; }`;
}

async function printDocument() {
  closeHf();
  try {
    buildPrintPages();
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    await platform.printPage();
  } catch (e) { await showError('Couldn\'t print the document', e); }
}

function syncTitle() {
  platform.setTitle(`${dirty() ? '* ' : ''}${fileName} - Folio`);
  (window as unknown as { __dirty?: boolean }).__dirty = dirty();
}

async function newDocument() {
  if (!(await confirmDiscard())) return;
  filePath = null; fileName = 'Untitled.docx'; isNew = true;
  mount(await blankDocument());
  say('');
}
async function openBytes(f: platform.OpenedFile) {
  try {
    const d = await WordDoc.open(f.bytes);
    filePath = f.path; fileName = f.name; isNew = false;
    mount(d);
    say(`Opened ${f.name}`);
  } catch (e) { await showError(`Couldn't open ${f.name}`, e); }
}
async function openDocument() {
  if (!(await confirmDiscard())) return;
  try {
    const f = await platform.pickDocument();
    if (f) await openBytes(f);
  } catch (e) { await showError('Couldn\'t open the file', e); }
}
async function save(forceAs = false): Promise<boolean> {
  try {
    let path = filePath;
    if (forceAs || !path || isNew) {
      path = await platform.pickSavePath(fileName, 'docx');
      if (!path) return false;
    }
    say('Saving');
    const body = bodyView.state.doc;
    const bytes = await exportDocx(doc, body);
    await platform.writeFile(path, bytes, 'docx');
    if (platform.isDesktop) { filePath = path; fileName = path.split(/[\\/]/).pop()!; }
    else fileName = path;
    isNew = false;
    savedBody = body;
    for (const part of doc.hf) part.saved = part.doc;
    savedRev = doc.commentRev;
    savedTrack = doc.track;
    syncTitle();
    say(`Saved ${fileName}`);
    return true;
  } catch (e) { say(''); await showError('Couldn\'t save the file', e); return false; }
}
async function exportToPdf() {
  try {
    const name = fileName.replace(/\.(docx|docm|dotx)$/i, '') + '.pdf';
    const path = await platform.pickSavePath(name, 'pdf');
    if (!path) return;
    say('Exporting PDF');
    // The PDF follows the review view, the same way printing does.
    const current = bodyView.state.doc;
    const body = review.mode === 'markup' ? current : settledCopy(doc, current, review.mode === 'final');
    let pdf: Uint8Array | null = null, engine = 'LibreOffice';
    try { pdf = await platform.libreOfficePdf(await exportDocx(doc, body, true)); }
    catch (e) { say('LibreOffice failed, using the built-in exporter'); console.warn(e); }
    if (!pdf) {
      const { exportPdf } = await import('./pdf/export');
      pdf = await exportPdf(doc, body);
      engine = 'built-in exporter';
    }
    await platform.writeFile(path, pdf, 'pdf');
    say(`Exported ${path.split(/[\\/]/).pop()} with the ${engine}`);
  } catch (e) { say(''); await showError('Couldn\'t export the PDF', e); }
}
async function showPreserved() {
  const c = preservedSummary();
  await modal(`<h2>Read-only content</h2><p>Folio shows these items but can't edit them. They stay in your document when you save.</p><ul>${Object.entries(c).map(([k, v]) => `<li>${esc(k)}: ${v}</li>`).join('')}</ul><div class="acts"><button class="primary" data-r="ok">OK</button></div>`);
  view.focus();
}

// A shortcut can arrive from both the native menu and the webview. Windows routes
// them differently from Linux and macOS, so both paths land here and duplicates drop.
const lastAction = new Map<string, number>();
async function showAbout() {
  await modal(`<div class="about"><h2>Folio</h2><p class="about-version">Version ${__APP_VERSION__}</p><p>Developed by Daniel Wood</p><p class="about-links"><a href="https://www.linkedin.com/in/danielewood">LinkedIn</a><a href="https://github.com/automatesecurity">GitHub</a></p><p class="about-statement">Bloat-free, open source word processing for the masses.</p></div><div class="acts"><button class="primary" data-r="ok">Close</button></div>`, (d) => {
    d.querySelector<HTMLButtonElement>('[data-r=ok]')!.focus();
    d.querySelectorAll<HTMLAnchorElement>('a').forEach((a) => a.addEventListener('click', (e) => { e.preventDefault(); platform.openExternal(a.href); }));
  });
  view.focus();
}

function handleMenu(id: string) {
  const now = performance.now();
  if (now - (lastAction.get(id) ?? -1e9) < 400) return;
  lastAction.set(id, now);
  if (($('modal') as HTMLDialogElement).open) return;
  switch (id) {
    case 'new': void newDocument(); break;
    case 'open': void openDocument(); break;
    case 'save': void save(); break;
    case 'save_as': void save(true); break;
    case 'export_pdf': void exportToPdf(); break;
    case 'print': void printDocument(); break;
    case 'about': void showAbout(); break;
    case 'review': review.toggle(); break;
    case 'track': void review.toggleTracking(); break;
    case 'new_comment': void review.newComment(); break;
    case 'accept_all': void review.settleAll(true); break;
    case 'reject_all': void review.settleAll(false); break;
    case 'margins': void chooseMargins(); break;
    case 'edit_header': openHf('header', currentPage()); break;
    case 'edit_footer': openHf('footer', currentPage()); break;
    case 'find': openFind(); break;
    case 'zoom_in': setZoom(zoom + 10); break;
    case 'zoom_out': setZoom(zoom - 10); break;
    case 'zoom_reset': setZoom(100); break;
  }
}
let zoom = 100;
function setZoom(z: number) {
  zoom = Math.min(200, Math.max(50, z));
  ($('zoomer').style as CSSStyleDeclaration & { zoom: string }).zoom = String(zoom / 100);
  $<HTMLInputElement>('zoom').value = String(zoom);
  $('zoom-val').textContent = `${zoom}%`;
}

function wireReview() {
  $('rv-track').addEventListener('click', () => void review.toggleTracking());
  $('st-track').addEventListener('click', () => void review.toggleTracking());
  $('st-review').addEventListener('click', () => review.show());
  $<HTMLSelectElement>('rv-view').addEventListener('change', (e) => review.setMode((e.target as HTMLSelectElement).value as 'markup'));
  $('rv-comment').addEventListener('click', () => void review.newComment());
  $('rv-prev').addEventListener('click', () => review.go(-1));
  $('rv-next').addEventListener('click', () => review.go(1));
  $('rv-accept').addEventListener('click', () => review.settleCurrent(true));
  $('rv-reject').addEventListener('click', () => review.settleCurrent(false));
  $('rv-accept-all').addEventListener('click', () => void review.settleAll(true));
  $('rv-reject-all').addEventListener('click', () => void review.settleAll(false));
  $('rv-pane').addEventListener('click', () => review.togglePane());
  $('rv-close').addEventListener('click', () => review.hide());
  $('rv-author').addEventListener('click', () => void askAuthor(authorName));
  document.querySelectorAll('#reviewbar button').forEach((b) => b.addEventListener('mousedown', (e) => e.preventDefault()));
}

function wireGlobals() {
  $('zoom').addEventListener('input', (e) => setZoom(Number((e.target as HTMLInputElement).value)));
  $('st-preserved').addEventListener('click', () => void showPreserved());
  $<HTMLInputElement>('color-input').addEventListener('input', (e) => {
    currentColor = (e.target as HTMLInputElement).value.slice(1).toUpperCase();
    (toolbar.querySelector('[data-id=color] i') as HTMLElement).style.setProperty('--c', `#${currentColor}`);
    run(setMark(schema.marks.color, { val: currentColor }))();
  });
  $('sheet').addEventListener('mousedown', (e) => {
    if (e.target !== $('sheet')) return;
    e.preventDefault();
    const y = e.clientY;
    const r = view.dom.getBoundingClientRect();
    const end = view.state.doc.content.size;
    view.dispatch(view.state.tr.setSelection(y < r.top ? TextSelection.near(view.state.doc.resolve(0)) : TextSelection.near(view.state.doc.resolve(end), -1)));
    view.focus();
  });
  document.addEventListener('keydown', (e) => {
    // A reload would discard the open document.
    if (e.key === 'F5' || ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'r')) { e.preventDefault(); return; }
    if (e.key === 'Escape' && hf && !($('modal') as HTMLDialogElement).open) { closeHf(); return; }
    if (!(e.ctrlKey || e.metaKey)) return;
    const k = e.key.toLowerCase();
    const id = e.shiftKey
      ? { s: 'save_as', e: 'export_pdf' }[k]
      : { k: 'link', f: 'find', s: 'save', o: 'open', n: 'new', p: 'print', '=': 'zoom_in', '-': 'zoom_out', '0': 'zoom_reset' }[k];
    if (e.altKey && k === 'm') { e.preventDefault(); handleMenu('new_comment'); return; }
    if (!id) return;
    e.preventDefault();
    if (id === 'link') void editLink(); else handleMenu(id);
  });
  platform.onMenu(handleMenu);
  platform.onCloseRequest(confirmDiscard);
}

async function start() {
  buildToolbar();
  wireFind();
  wireHf();
  wireReview();
  wireGlobals();
  mount(await blankDocument());
  const launch = await platform.launchDocument().catch(() => null);
  if (launch) await openBytes(launch);
}
void start();

// UI test hook. Not present in the desktop app.
if (!platform.isDesktop) (window as unknown as Record<string, unknown>).__folio = {
  openBytes: (b: number[], name: string) => openBytes({ path: null, name, bytes: new Uint8Array(b) }),
  save: async () => Array.from(await exportDocx(doc, bodyView.state.doc)),
  pdf: async () => { const { exportPdf } = await import('./pdf/export'); return Array.from(await exportPdf(doc, bodyView.state.doc)); },
  view: () => view,
  preparePrint: buildPrintPages,
  menu: handleMenu,
  review: () => review,
  body: () => bodyView,
};
