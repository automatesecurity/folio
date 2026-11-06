import { Fragment, Mark, Node as PMNode } from 'prosemirror-model';
import { schema } from '../editor/schema';
import { session } from '../editor/session';
import { DocxError, Package, Rel, resolveTarget } from './package';
import { Comment, readComments } from './comments';
import { Numbering, parseTheme, StyleSheet } from './styles';
import { attr, child, elements, escAttr, local, parseXml, rawAttrs, textOf, walk, XEl } from './xml';

export interface Source { xml: string; node: PMNode }
export interface PageSetup { width: number; height: number; top: number; bottom: number; left: number; right: number; header: number; footer: number; landscape: boolean }
export type HfKind = 'header' | 'footer';
export type HfSlot = 'default' | 'first' | 'even';
export interface Margins { top: number; bottom: number; left: number; right: number; header: number; footer: number }

/** One header or footer. Each is its own small document with its own part in the package. */
export interface HfPart {
  kind: HfKind;
  slot: HfSlot;
  /** Null until the part has to exist in the package. */
  path: string | null;
  head: string;
  tail: string;
  w: string;
  rPrefix: string | null;
  doc: PMNode;
  /** What the app last saved, for the unsaved-changes check. */
  saved: PMNode;
  /** What was last written into the package. */
  written: PMNode | null;
}

const MAIN_NS = /(wordprocessingml\/2006\/main|ooxml\/wordprocessingml\/main)$/;
const REL_NS = /(officeDocument\/2006\/relationships|ooxml\/officeDocument\/relationships)$/;
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RENDERABLE = new Set(['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg']);
const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };

export class WordDoc {
  pkg!: Package;
  mainPath = 'word/document.xml';
  /** Namespace prefix for WordprocessingML including the colon ('w:'), or '' if it is the default namespace. */
  w = 'w:';
  rPrefix: string | null = 'r';
  head = '';
  tail = '';
  sources = new Map<number, Source>();
  images = new Map<string, { bytes: Uint8Array; ext: string }>();
  rels: Rel[] = [];
  styles!: StyleSheet;
  numbering!: Numbering;
  page: PageSetup = { width: 12240, height: 15840, top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720, landscape: false };
  hf: HfPart[] = [];
  comments = new Map<string, Comment>();
  /** Bumped by every comment edit, for the unsaved-changes check. */
  commentRev = 0;
  commentsPath: string | null = null;
  commentsHead = '';
  commentsTail = '';
  extendedPath: string | null = null;
  extendedHead = '';
  extendedTail = '';
  track = false;
  settingsPath: string | null = null;
  private revCounter = 100000;
  evenOdd = false;
  originalSect: { margins: Margins; titlePg: boolean } | null = null;
  addedRefs: { tag: string; slot: HfSlot; rId: string }[] = [];
  /** How far headers and footers reach from the page edge, in twips. Tall ones push the body in. */
  hfTop = 0;
  hfBottom = 0;
  body!: PMNode;
  notes: string[] = [];
  linkRels = new Map<string, string>();
  private srcCounter = 0;
  private src = '';
  private imgPrefix = '';

  get effTop() { return Math.max(this.page.top, this.hfTop); }
  get effBottom() { return Math.max(this.page.bottom, this.hfBottom); }
  findPart(kind: HfKind, slot: HfSlot) { return this.hf.find((p) => p.kind === kind && p.slot === slot); }
  slotFor(pageIndex: number, titlePg: boolean): HfSlot {
    if (pageIndex === 0 && titlePg) return 'first';
    return this.evenOdd && pageIndex % 2 === 1 ? 'even' : 'default';
  }

  /** Finds the part, or starts an empty one. Nothing reaches the package until it has content. */
  hfPart(kind: HfKind, slot: HfSlot): HfPart {
    const found = this.findPart(kind, slot);
    if (found) return found;
    const styleId = kind === 'header' ? 'Header' : 'Footer';
    const pStyle = this.styles.styles.has(styleId) ? styleId : null;
    const body = schema.nodes.doc.create(null, schema.nodes.paragraph.create({ pStyle }));
    const w = this.w, tag = kind === 'header' ? 'hdr' : 'ftr';
    const nsW = w ? `xmlns:${w.slice(0, -1)}` : 'xmlns';
    const part: HfPart = {
      kind, slot, path: null, w, rPrefix: 'r', doc: body, saved: body, written: null, tail: `</${w}${tag}>`,
      head: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${w}${tag} ${nsW}="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${REL_BASE}">`,
    };
    this.hf.push(part);
    return part;
  }

  async allocatePart(part: HfPart) {
    if (part.path) return;
    const dir = this.mainPath.slice(0, this.mainPath.lastIndexOf('/') + 1);
    let n = 1;
    while (this.pkg.has(`${dir}${part.kind}${n}.xml`) || this.hf.some((h) => h.path === `${dir}${part.kind}${n}.xml`)) n++;
    const path = `${dir}${part.kind}${n}.xml`;
    this.pkg.set(path, part.head + part.tail);
    await this.pkg.ensureContentOverride(`/${path}`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${part.kind}+xml`);
    const rId = await this.pkg.addRel(this.mainPath, `${REL_BASE}/${part.kind}`, `${part.kind}${n}.xml`);
    this.addedRefs.push({ tag: `${part.kind}Reference`, slot: part.slot, rId });
    part.path = path;
  }

  static async open(bytes: Uint8Array): Promise<WordDoc> {
    const d = new WordDoc();
    d.pkg = await Package.load(bytes);
    await d.load();
    return d;
  }

  private async load() {
    this.mainPath = await this.pkg.mainPart();
    this.src = (await this.pkg.text(this.mainPath))!;
    let root: XEl;
    try { root = parseXml(this.src); }
    catch (e) { throw new DocxError(`The document's content is damaged: ${(e as Error).message}`); }
    if (local(root.name) !== 'document') throw new DocxError('This package does not contain a Word document body. It may be a different Office file.');

    this.w = '';
    this.rPrefix = null;
    for (const k in root.attrs) {
      if (MAIN_NS.test(root.attrs[k])) this.w = k === 'xmlns' ? '' : `${k.slice(6)}:`;
      if (REL_NS.test(root.attrs[k]) && k.startsWith('xmlns:')) this.rPrefix = k.slice(6);
    }
    const body = child(root, 'body');
    if (!body) throw new DocxError('This document has no body.');

    this.rels = await this.pkg.rels(this.mainPath);
    const relPath = (suffix: string) => { const r = this.rels.find((x) => x.type.endsWith(suffix) && !x.external); return r ? resolveTarget(this.mainPath, r.target) : null; };
    const stylesPath = relPath('/styles'), numPath = relPath('/numbering'), themePath = relPath('/theme');
    this.styles = new StyleSheet(stylesPath ? await this.pkg.text(stylesPath) : null, parseTheme(themePath ? await this.pkg.text(themePath) : null));
    this.numbering = new Numbering(numPath ? await this.pkg.text(numPath) : null);

    session.reset();
    session.styles = this.styles;
    session.numbering = this.numbering;
    session.newList = (kind) => this.numbering.newList(kind, this.w, true);

    await this.loadImages(this.mainPath, '');
    this.settingsPath = relPath('/settings');
    const settingsXml = (this.settingsPath && await this.pkg.text(this.settingsPath)) ?? '';
    this.evenOdd = /<[\w.:-]*evenAndOddHeaders\b(?![^>]*val="(0|false)")/.test(settingsXml);
    this.track = /<[\w.:-]*trackRevisions\b(?![^>]*val="(0|false)")/.test(settingsXml);
    for (const m of this.src.matchAll(/\bid="(\d+)"/g)) this.revCounter = Math.max(this.revCounter, Number(m[1]) + 1);
    await readComments(this);

    const kids = elements(body);
    const lastSect = kids.length && local(kids[kids.length - 1].name) === 'sectPr' ? kids[kids.length - 1] : undefined;
    const blocks = lastSect ? kids.slice(0, -1) : kids;
    let cursor = body.innerStart;
    this.head = this.src.slice(0, cursor);
    const nodes: PMNode[] = [];
    for (const el of blocks) {
      nodes.push(this.block(el, this.src.slice(cursor, el.end)));
      cursor = el.end;
    }
    this.tail = this.src.slice(cursor);
    if (!nodes.length) nodes.push(schema.nodes.paragraph.create());
    const titlePg = await this.loadPage(root, lastSect);
    const { top, bottom, left, right, header, footer } = this.page;
    const margins = { top, bottom, left, right, header, footer };
    this.originalSect = { margins, titlePg };
    this.body = schema.nodes.doc.create({ pgMar: margins, titlePg }, nodes);
  }

  private async loadImages(owner: string, prefix: string) {
    for (const r of await this.pkg.rels(owner)) {
      if (!r.type.endsWith('/image') || r.external) continue;
      const path = resolveTarget(owner, r.target);
      const ext = path.split('.').pop()!.toLowerCase();
      const bytes = await this.pkg.bytes(path);
      if (!bytes || !RENDERABLE.has(ext)) continue;
      const key = prefix + r.id;
      this.images.set(key, { bytes, ext });
      session.assets.set(key, URL.createObjectURL(new Blob([bytes as BlobPart], { type: MIME[ext] })));
    }
  }

  private async loadPage(root: XEl, sect?: XEl) {
    const p = this.page;
    if (sect) {
      const sz = child(sect, 'pgSz'), m = child(sect, 'pgMar');
      const num = (e: XEl | undefined, k: string, d: number) => { const v = Number(attr(e, k)); return Number.isFinite(v) && attr(e, k) !== undefined ? Math.abs(v) : d; };
      p.width = num(sz, 'w', p.width); p.height = num(sz, 'h', p.height);
      p.top = num(m, 'top', p.top); p.bottom = num(m, 'bottom', p.bottom);
      p.left = num(m, 'left', p.left); p.right = num(m, 'right', p.right);
      p.header = num(m, 'header', p.header); p.footer = num(m, 'footer', p.footer);
      p.landscape = attr(sz, 'orient') === 'landscape';
    }
    const allSects = [...(sect ? [sect] : []), ...[...walk(root)].filter((e) => local(e.name) === 'sectPr')];
    for (const sec of allSects) {
      for (const ref of elements(sec)) {
        const l = local(ref.name);
        if (l !== 'headerReference' && l !== 'footerReference') continue;
        const kind: HfKind = l === 'headerReference' ? 'header' : 'footer';
        const slot = (attr(ref, 'type') ?? 'default') as HfSlot;
        if (!['default', 'first', 'even'].includes(slot) || this.findPart(kind, slot)) continue;
        const rel = this.rels.find((r) => r.id === attr(ref, 'id'));
        if (!rel || rel.external) continue;
        const part = await this.readPart(kind, slot, resolveTarget(this.mainPath, rel.target));
        if (part) this.hf.push(part);
      }
    }
    return allSects.some((x) => !!child(x, 'titlePg'));
  }

  private async readPart(kind: HfKind, slot: HfSlot, path: string): Promise<HfPart | null> {
    const xml = await this.pkg.text(path);
    if (xml === null) return null;
    const root = parseXml(xml);
    const saved = { src: this.src, rels: this.rels, prefix: this.imgPrefix };
    this.src = xml;
    this.rels = await this.pkg.rels(path);
    this.imgPrefix = `${path}#`;
    await this.loadImages(path, this.imgPrefix);
    let w = '', rPrefix: string | null = null;
    for (const k in root.attrs) {
      if (MAIN_NS.test(root.attrs[k])) w = k === 'xmlns' ? '' : `${k.slice(6)}:`;
      if (REL_NS.test(root.attrs[k]) && k.startsWith('xmlns:')) rPrefix = k.slice(6);
    }
    const selfClosed = root.innerStart === root.end;
    let cursor = root.innerStart;
    const head = selfClosed ? `${xml.slice(0, root.end - 2)}>` : xml.slice(0, cursor);
    const nodes: PMNode[] = [];
    for (const el of elements(root)) {
      nodes.push(this.block(el, xml.slice(cursor, el.end)));
      cursor = el.end;
    }
    const tail = selfClosed ? `</${root.name}>${xml.slice(root.end)}` : xml.slice(cursor);
    if (!nodes.length) nodes.push(schema.nodes.paragraph.create());
    const doc = schema.nodes.doc.create(null, nodes);
    this.src = saved.src; this.rels = saved.rels; this.imgPrefix = saved.prefix;
    return { kind, slot, path, head, tail, w, rPrefix, doc, saved: doc, written: doc };
  }

  // Blocks -----------------------------------------------------------------

  private register(node: PMNode, xml: string): PMNode {
    this.sources.set(node.attrs.srcId, { xml, node });
    return node;
  }
  bodySource() { return this.src; }
  commentsWritten: string | null = null;
  private nextId() { return ++this.srcCounter; }
  nextRevId() { return ++this.revCounter; }
  private raw(el: XEl) { return this.src.slice(el.start, el.end); }

  private block(el: XEl, xml: string): PMNode {
    const l = local(el.name);
    if (l === 'p') return this.paragraph(el, xml);
    if (l === 'tbl') return this.table(el, xml) ?? this.opaqueBlock(el, xml, 'table');
    return this.opaqueBlock(el, xml, /^(bookmark|comment|perm|move.*Range)/.test(l) ? 'bookmark' : l === 'sdt' ? 'sdt' : 'other');
  }

  private opaqueBlock(el: XEl, xml: string, kind: string): PMNode {
    const preview: { text: string; cls: string | null }[] = [];
    const visit = (e: XEl) => {
      if (preview.length > 500) return;
      const l = local(e.name);
      if (l === 'p') {
        preview.push({ text: textOf(e).replace(/\t/g, '    '), cls: attr(child(child(e, 'pPr') ?? e, 'pStyle'), 'val') ?? null });
      } else if (l === 'tr') {
        preview.push({ text: elements(e).filter((c) => local(c.name) === 'tc').map((c) => textOf(c)).join('  |  '), cls: null });
      } else elements(e).forEach(visit);
    };
    visit(el);
    return this.register(schema.nodes.opaque_block.create({ srcId: this.nextId(), kind, xml, preview, session: session.id }), xml);
  }

  private pPrAttrs(pPr: XEl): Record<string, unknown> {
    const attrs: Record<string, unknown> = { pExtra: [] };
    const extra: string[] = [];
    for (const c of elements(pPr)) {
      switch (local(c.name)) {
        case 'pStyle': attrs.pStyle = attr(c, 'val') ?? null; break;
        case 'jc': attrs.jc = attr(c, 'val') ?? null; break;
        case 'ind': {
          const m: Record<string, string> = {};
          for (const k in c.attrs) m[local(k)] = c.attrs[k];
          attrs.ind = m;
          break;
        }
        case 'numPr': {
          const parts = elements(c).map((x) => local(x.name));
          if (parts.every((x) => x === 'numId' || x === 'ilvl')) {
            attrs.numId = attr(child(c, 'numId'), 'val') ?? null;
            attrs.ilvl = Number(attr(child(c, 'ilvl'), 'val') ?? 0);
          } else extra.push(this.raw(c));
          break;
        }
        case 'rPr': {
          // A tracked paragraph mark lives here. Everything else stays raw.
          const marks = elements(c).filter((x) => ['ins', 'del'].includes(local(x.name)));
          let raw = this.raw(c);
          const pmark: Record<string, Revision> = {};
          for (const m of [...marks].reverse()) {
            pmark[local(m.name)] = revisionOf(m);
            raw = raw.slice(0, m.start - c.start) + raw.slice(m.end - c.start);
          }
          if (marks.length) attrs.pmark = pmark;
          if (elements(c).length > marks.length) extra.push(raw);
          break;
        }
        default: extra.push(this.raw(c));
      }
    }
    attrs.pExtra = extra;
    return attrs;
  }

  private paragraph(el: XEl, xml: string): PMNode {
    const pPr = child(el, 'pPr');
    const attrs = { srcId: this.nextId(), rawAttrs: rawAttrs(el), ...(pPr ? this.pPrAttrs(pPr) : {}) };
    const inline = this.inlines(elements(el).filter((c) => local(c.name) !== 'pPr'), []);
    return this.register(schema.nodes.paragraph.create(attrs, Fragment.fromArray(inline)), xml);
  }

  /** The formatting recorded inside a tracked rPrChange, for rejecting it. */
  marksFromChange(fragments: string[]): readonly Mark[] | null {
    const xml = fragments.find((x) => /^<[\w.-]*:?rPrChange\b/.test(x));
    if (!xml) return null;
    const keep = this.src;
    this.src = xml;
    const old = this.marks(child(parseXml(xml), 'rPr'), []);
    this.src = keep;
    return old;
  }
  /** The paragraph properties recorded inside a tracked pPrChange, for rejecting it. */
  paraFromChange(fragments: string[]): Record<string, unknown> | null {
    const xml = fragments.find((x) => /^<[\w.-]*:?pPrChange\b/.test(x));
    if (!xml) return null;
    const keep = this.src;
    this.src = xml;
    const old = child(parseXml(xml), 'pPr');
    const attrs = old ? this.pPrAttrs(old) : { pExtra: [] };
    this.src = keep;
    return { pStyle: null, jc: null, numId: null, ilvl: null, ind: null, ...attrs };
  }

  private table(el: XEl, xml: string): PMNode | null {
    interface Cell { el: XEl; col: number; span: number; rows: number; row: number }
    const grid: number[] = [];
    let tblPr = '', gridRaw = '';
    const rowEls: XEl[] = [];
    for (const c of elements(el)) {
      const l = local(c.name);
      if (l === 'tblPr') tblPr = this.raw(c);
      else if (l === 'tblGrid') {
        gridRaw = this.raw(c);
        for (const g of elements(c)) if (local(g.name) === 'gridCol') grid.push(Number(attr(g, 'w')) || 0);
      } else if (l === 'tr') rowEls.push(c);
      else return null; // row-level content controls, bookmarks and so on
    }
    if (!rowEls.length) return null;
    const rows: Cell[][] = [];
    let width = 0;
    for (let r = 0; r < rowEls.length; r++) {
      const trPr = child(rowEls[r], 'trPr');
      if (trPr && elements(trPr).some((c) => ['gridBefore', 'gridAfter', 'wBefore', 'wAfter'].includes(local(c.name)))) return null;
      const out: Cell[] = [];
      let col = 0;
      for (const c of elements(rowEls[r])) {
        const l = local(c.name);
        if (l === 'trPr' || l === 'tblPrEx') continue;
        if (l !== 'tc') return null;
        const tcPr = child(c, 'tcPr');
        if (child(tcPr ?? c, 'hMerge') && tcPr) return null;
        const span = Number(attr(child(tcPr ?? c, 'gridSpan'), 'val') ?? 1) || 1;
        const vm = tcPr ? child(tcPr, 'vMerge') : undefined;
        if (vm && attr(vm, 'val') !== 'restart') {
          const above = r > 0 ? findCovering(rows, r - 1, col) : undefined;
          if (!above || above.span !== span) return null;
          if (textOf(c).trim() || elements(c).filter((x) => local(x.name) !== 'tcPr').length > 1) return null;
          above.rows++;
        } else out.push({ el: c, col, span, rows: 1, row: r });
        col += span;
      }
      if (!out.length) return null;
      if (r === 0) width = col;
      else if (col !== width) return null;
      rows.push(out);
    }
    if (grid.length && grid.length !== width) return null;

    const nodes = rows.map((cells, r) => {
      const rowEl = rowEls[r];
      const content = cells.map((c) => this.cell(c.el, c.span, c.rows, grid, c.col));
      const trPr = child(rowEl, 'trPr');
      const tblEx = child(rowEl, 'tblPrEx');
      let trRaw = trPr ? this.raw(trPr) : '';
      let rowChange: { kind: string; rev: Revision } | null = null;
      const tracked = trPr && elements(trPr).find((x) => ['ins', 'del'].includes(local(x.name)));
      if (trPr && tracked) {
        rowChange = { kind: local(tracked.name), rev: revisionOf(tracked) };
        trRaw = trRaw.slice(0, tracked.start - trPr.start) + trRaw.slice(tracked.end - trPr.start);
        if (elements(trPr).length === 1) trRaw = '';
      }
      const node = schema.nodes.table_row.create(
        { srcId: this.nextId(), trPr: (tblEx ? this.raw(tblEx) : '') + trRaw, rawAttrs: rawAttrs(rowEl), rowChange }, content);
      return this.register(node, this.raw(rowEl));
    });
    return this.register(schema.nodes.table.create({ srcId: this.nextId(), tblPr, gridRaw, rawAttrs: rawAttrs(el) }, nodes), xml);
  }

  private cell(el: XEl, span: number, rows: number, grid: number[], col: number): PMNode {
    const tcPr = child(el, 'tcPr');
    const extra = tcPr ? elements(tcPr).filter((c) => !['tcW', 'gridSpan', 'vMerge'].includes(local(c.name))).map((c) => this.raw(c)) : [];
    const blocks = elements(el).filter((c) => local(c.name) !== 'tcPr').map((c) => this.block(c, this.raw(c)));
    if (!blocks.length) blocks.push(schema.nodes.paragraph.create());
    const colwidth = grid.length ? grid.slice(col, col + span) : null;
    const node = schema.nodes.table_cell.create({ srcId: this.nextId(), colspan: span, rowspan: rows, colwidth, tcPr: extra, rawAttrs: rawAttrs(el) }, blocks);
    return this.register(node, this.raw(el));
  }

  // Inline -----------------------------------------------------------------

  private opaque(kind: string, xml: string, extra: Record<string, unknown> = {}) {
    return schema.nodes.opaque_inline.create({ kind, xml, session: session.id, ...extra });
  }

  private inlines(els: XEl[], base: readonly Mark[]): PMNode[] {
    const tracked = base.filter((m) => m.type === schema.marks.tracked_ins || m.type === schema.marks.tracked_del);
    const out = new TrackedList(tracked);
    for (let i = 0; i < els.length; i++) {
      const el = els[i];
      const l = local(el.name);
      if (l === 'r') {
        const end = fieldEnd(els, i);
        if (end >= i) {
          const body = els.slice(i, end + 1);
          const code = fieldCode(body);
          out.push(this.opaque('field', this.src.slice(el.start, els[end].end), { label: fieldResult(body) || code, instr: code.split(/\s+/)[0]?.toUpperCase() || null }));
          i = end;
        } else out.push(...this.run(el, base));
        continue;
      }
      switch (l) {
        case 'proofErr': break; // spell-check residue: kept while the paragraph is unedited, stale afterwards
        case 'hyperlink': {
          const simple = elements(el).every((c) => ['r', 'proofErr', 'bookmarkStart', 'bookmarkEnd'].includes(local(c.name)));
          if (!simple) { out.push(this.opaque('other', this.raw(el), { label: textOf(el) })); break; }
          const rId = attr(el, 'id') ?? null;
          const rel = this.rels.find((r) => r.id === rId);
          const mark = schema.marks.link.create({ href: rel?.target ?? null, rId, anchor: attr(el, 'anchor') ?? null, rawAttrs: rawAttrs(el) });
          out.push(...this.inlines(elements(el), mark.addToSet(base as Mark[])));
          break;
        }
        case 'bookmarkStart': case 'bookmarkEnd': out.push(this.opaque('bookmark', this.raw(el))); break;
        case 'ins': case 'del': {
          if (!elements(el).every((c) => TRACKABLE.has(local(c.name)))) { out.push(this.opaque(l, this.raw(el), { label: textOf(el) })); break; }
          const rev = revisionOf(el);
          const mark = (l === 'ins' ? schema.marks.tracked_ins : schema.marks.tracked_del).create({ author: rev.author, date: rev.date, raw: rev.raw });
          out.push(...this.inlines(elements(el), mark.addToSet(base as Mark[])));
          break;
        }
        case 'moveTo': out.push(this.opaque('ins', this.raw(el), { label: textOf(el) })); break;
        case 'moveFrom': out.push(this.opaque('del', this.raw(el), { label: textOf(el) })); break;
        case 'fldSimple': out.push(this.opaque('field', this.raw(el), { label: textOf(el) || attr(el, 'instr') || '', instr: attr(el, 'instr')?.trim().split(/\s+/)[0]?.toUpperCase() ?? null })); break;
        case 'oMath': case 'oMathPara': out.push(this.opaque('math', this.raw(el), { label: textOf(el, new Set(['t'])) })); break;
        default:
          if (/Range(Start|End)$|^perm(Start|End)$/.test(l)) out.push(this.opaque('range', this.raw(el)));
          else out.push(this.opaque('other', this.raw(el), { label: textOf(el) }));
      }
    }
    return out.nodes;
  }

  private run(el: XEl, base: readonly Mark[]): PMNode[] {
    const rPr = child(el, 'rPr');
    const marks = this.marks(rPr, base);
    const rPrRaw = rPr ? this.raw(rPr) : '';
    const out: PMNode[] = [];
    let text = '';
    const flush = () => { if (text) { out.push(schema.text(text, marks)); text = ''; } };
    for (const c of elements(el)) {
      const l = local(c.name);
      switch (l) {
        case 'rPr': case 'lastRenderedPageBreak': break;
        case 't': case 'delText': {
          let s = c.children.map((x) => (x.kind === 'text' ? x.text : '')).join('');
          if (c.attrs['xml:space'] !== 'preserve') s = s.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '');
          text += s.replace(/[\r\n]+/g, ' ');
          break;
        }
        case 'tab': text += '\t'; break;
        case 'noBreakHyphen': text += '‑'; break;
        case 'softHyphen': text += '­'; break;
        case 'cr': flush(); out.push(schema.nodes.hard_break.create({ xml: this.raw(c) }, null, marks)); break;
        case 'br':
          flush();
          if (attr(c, 'type') === 'page') out.push(schema.nodes.page_break.create({ xml: this.raw(c) }, null, marks));
          else out.push(schema.nodes.hard_break.create({ xml: this.raw(c) }, null, marks));
          break;
        default: {
          flush();
          const xml = wrapRun(el, this.w, rPrRaw, this.raw(c), this.src);
          out.push(this.opaqueRunChild(l, c, xml));
        }
      }
    }
    flush();
    return out;
  }

  private opaqueRunChild(l: string, c: XEl, xml: string): PMNode {
    switch (l) {
      case 'drawing': case 'pict': case 'object': case 'AlternateContent': return this.opaque('drawing', xml, drawingInfo(c, this.images, this.imgPrefix));
      case 'footnoteReference': case 'endnoteReference': this.notes.push(l); return this.opaque('note', xml, { label: `[${attr(c, 'id') ?? ''}]` });
      case 'commentReference': return this.opaque('comment', xml, { label: '\u{1F4AC}' });
      case 'fldChar': case 'instrText': case 'delInstrText': return this.opaque('fieldPart', xml);
      case 'sym': return this.opaque('sym', xml, { label: '□' });
      default: return this.opaque('other', xml, { label: textOf(c) });
    }
  }

  private marks(rPr: XEl | undefined, base: readonly Mark[]): readonly Mark[] {
    let set: readonly Mark[] = base;
    if (!rPr) return set;
    const m = schema.marks;
    const add = (mk: Mark) => { set = mk.addToSet(set as Mark[]); };
    const extra: string[] = [];
    for (const c of elements(rPr)) {
      const l = local(c.name);
      const v = attr(c, 'val');
      const off = v === '0' || v === 'false' || v === 'off';
      switch (l) {
        case 'rStyle': if (v) add(m.rstyle.create({ id: v })); else extra.push(this.raw(c)); break;
        case 'b': if (off) extra.push(this.raw(c)); else add(m.bold.create()); break;
        case 'i': if (off) extra.push(this.raw(c)); else add(m.italic.create()); break;
        case 'strike': if (off) extra.push(this.raw(c)); else add(m.strike.create()); break;
        case 'u': if (v === 'none') extra.push(this.raw(c)); else add(m.underline.create({ val: v ?? 'single', raw: this.raw(c) })); break;
        case 'vertAlign': if (v === 'superscript' || v === 'subscript') add(m.vert.create({ val: v })); else extra.push(this.raw(c)); break;
        case 'color': if (v) add(m.color.create({ val: v, raw: this.raw(c) })); else extra.push(this.raw(c)); break;
        case 'highlight': if (v && v !== 'none') add(m.highlight.create({ val: v })); else extra.push(this.raw(c)); break;
        case 'sz': if (v) add(m.size.create({ val: Number(v) })); else extra.push(this.raw(c)); break;
        case 'rFonts': {
          const th = attr(c, 'asciiTheme');
          const name = th ? (th.startsWith('major') ? this.styles.theme.major : this.styles.theme.minor) : attr(c, 'ascii');
          if (name) add(m.font.create({ name, raw: this.raw(c) })); else extra.push(this.raw(c));
          break;
        }
        default: extra.push(this.raw(c));
      }
    }
    if (extra.length) add(m.rextra.create({ xml: extra }));
    return set;
  }

  // Additions from the editor -------------------------------------------------

  async addImage(bytes: Uint8Array, ext: string, part?: HfPart): Promise<{ rId: string; key: string }> {
    if (part) await this.allocatePart(part);
    const owner = part?.path ?? this.mainPath;
    let n = 1;
    while (this.pkg.has(`word/media/folio${n}.${ext}`)) n++;
    const path = `word/media/folio${n}.${ext}`;
    this.pkg.set(path, bytes);
    await this.pkg.ensureContentDefault(ext, MIME[ext]);
    const rId = await this.pkg.addRel(owner, `${REL_BASE}/image`, `media/folio${n}.${ext}`);
    const key = part ? `${owner}#${rId}` : rId;
    this.images.set(key, { bytes, ext });
    session.assets.set(key, URL.createObjectURL(new Blob([bytes as BlobPart], { type: MIME[ext] })));
    return { rId, key };
  }

  drawingXml(rId: string, wPx: number, hPx: number, id: number, name: string): string {
    const cx = Math.round(wPx * 9525), cy = Math.round(hPx * 9525), w = this.w;
    const safe = escAttr(name);
    return `<${w}r><${w}drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="${safe}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="${safe}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="${REL_BASE}" r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></${w}drawing></${w}r>`;
  }
}

function findCovering(rows: { col: number; span: number; rows: number; row: number }[][], r: number, col: number) {
  for (let i = r; i >= 0; i--) {
    const c = rows[i].find((x) => x.col === col);
    if (c && c.row + c.rows - 1 === r) return c;
    if (c) return undefined;
  }
  return undefined;
}

function wrapRun(run: XEl, w: string, rPr: string, inner: string, src: string) {
  const open = src.slice(run.start, run.innerStart);
  return `${open}${rPr}${inner}</${run.name}>`;
}

const TRACKABLE = new Set(['r', 'ins', 'del', 'proofErr', 'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd']);

export interface Revision { author: string; date: string; raw: string | null }
const revisionOf = (el: XEl): Revision => ({ author: attr(el, 'author') ?? '', date: attr(el, 'date') ?? '', raw: rawAttrs(el) });

/** Carries a tracked insertion or deletion across every node inside it, atoms included. */
class TrackedList {
  nodes: PMNode[] = [];
  constructor(private marks: readonly Mark[]) {}
  push(...items: PMNode[]) {
    for (const n of items) this.nodes.push(this.marks.reduce((x, m) => x.mark(m.addToSet(x.marks as Mark[])), n));
  }
}

const fldTypes = (run: XEl) => elements(run).filter((c) => local(c.name) === 'fldChar').map((c) => attr(c, 'fldCharType'));

/** If a complex field opens at index i and closes in this paragraph, returns the closing index. */
function fieldEnd(els: XEl[], i: number): number {
  if (fldTypes(els[i])[0] !== 'begin') return -1;
  let depth = 0;
  for (let j = i; j < els.length; j++) {
    if (local(els[j].name) !== 'r') continue;
    for (const t of fldTypes(els[j])) {
      if (t === 'begin') depth++;
      else if (t === 'end' && --depth === 0) return j;
    }
  }
  return -1;
}
function fieldResult(runs: XEl[]) {
  let seen = false, out = '';
  for (const r of runs) {
    const types = fldTypes(r);
    if (types.includes('separate')) seen = true;
    if (!types.length && seen) out += textOf(r);
  }
  return out;
}
const fieldCode = (runs: XEl[]) => runs.map((r) => textOf(r, new Set(['instrText']))).join('').trim();

function drawingInfo(el: XEl, images: Map<string, unknown>, prefix: string) {
  let img: string | null = null, w = 0, h = 0, label = '';
  let hasText = '';
  for (const e of walk(el)) {
    const l = local(e.name);
    if (l === 'blip' && !img) { const id = attr(e, 'embed'); if (id && images.has(prefix + id)) img = prefix + id; }
    if (l === 'imagedata' && !img) { const id = attr(e, 'id'); if (id && images.has(prefix + id)) img = prefix + id; }
    if (l === 'extent' && !w) { w = Number(attr(e, 'cx') ?? 0) / 9525; h = Number(attr(e, 'cy') ?? 0) / 9525; }
    if (l === 'shape' && !w) {
      const s = attr(e, 'style') ?? '';
      const mw = /width:([\d.]+)pt/.exec(s), mh = /height:([\d.]+)pt/.exec(s);
      if (mw) w = (+mw[1] * 96) / 72;
      if (mh) h = (+mh[1] * 96) / 72;
    }
    if (l === 'docPr') label = attr(e, 'descr') || attr(e, 'name') || label;
    if (l === 'chart') label = 'Chart';
    if (l === 'txbxContent') hasText = textOf(e);
  }
  if (hasText) label = hasText;
  return { img, w, h, label };
}

/** Stand-ins for page fields in built-in PDF text, swapped per page. */
export const PAGE_FIELD = String.fromCharCode(0xe000);
export const PAGES_FIELD = String.fromCharCode(0xe001);
export const fillPageFields = (s: string, page: number, pages: number) => s.split(PAGE_FIELD).join(String(page)).split(PAGES_FIELD).join(String(pages));
