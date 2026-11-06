import { Mark, Node as PMNode } from 'prosemirror-model';
import { TableMap } from 'prosemirror-tables';
import { schema } from '../editor/schema';
import { resolveTarget } from './package';
import type { HfPart, Margins, Revision, WordDoc } from './document';
import { writeComments } from './comments';
import { elements, escAttr, escText, local, parseXml } from './xml';

const RPR_ORDER = 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange'.split(' ');
const PPR_ORDER = 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split(' ');
const TCPR_ORDER = 'cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark headers cellIns cellDel cellMerge tcPrChange'.split(' ');

const tagName = (x: string) => /^<(?:[\w.-]+:)?([\w.-]+)/.exec(x)?.[1] ?? '';
function ordered(items: string[], order: string[]) {
  const rank = (x: string) => { const i = order.indexOf(tagName(x)); return i < 0 ? order.length : i; };
  return items.map((x, i) => [x, i] as const).sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1]).map((x) => x[0]).join('');
}

/**
 * Bookmarks, comment ranges and split fields span paragraphs. A partial delete can
 * orphan half a pair, and a paste can duplicate an id Word requires to be unique.
 */
function orphanAtoms(doc: PMNode) {
  const drop = new Set<PMNode>();
  const starts = new Map<string, PMNode>();
  const ends = new Map<string, PMNode>();
  const refs = new Set<string>();
  const stack: PMNode[] = [];
  doc.descendants((n) => {
    if (n.type !== schema.nodes.opaque_inline) return;
    const x = n.attrs.xml as string;
    const kind = n.attrs.kind;
    const id = /\bid="([^"]+)"/.exec(x)?.[1] ?? '';
    if (kind === 'bookmark' || kind === 'range') {
      const tag = tagName(x);
      const key = `${tag.replace(/(Start|End)$/, '')}:${id}`;
      if (tag.endsWith('Start')) { if (starts.has(key)) drop.add(n); else starts.set(key, n); }
      else if (!starts.has(key) || ends.has(key)) drop.add(n);
      else ends.set(key, n);
    } else if (kind === 'comment') {
      if (refs.has(id)) drop.add(n); else refs.add(id);
    } else if (kind === 'fieldPart') {
      const t = /fldCharType="(\w+)"/.exec(x)?.[1];
      if (t === 'begin') stack.push(n);
      else if (t === 'end') { if (stack.length) stack.pop(); else drop.add(n); }
      else if (!stack.length) drop.add(n);
    }
  });
  for (const [key, n] of starts) if (!ends.has(key)) drop.add(n);
  for (const n of stack) drop.add(n);
  return drop;
}

class Writer {
  used = new Set<number>();
  revIds = new Set<string>();
  paraIds = new Set<string>();
  w: string;
  constructor(private d: WordDoc, private drop: Set<PMNode>, w: string, private rPrefix: string | null, private owner: string) { this.w = w; }

  private verbatim(node: PMNode): string | null {
    const id = node.attrs.srcId as number | null;
    if (id == null || this.used.has(id)) return null;
    const s = this.d.sources.get(id);
    if (!s || !s.node.eq(node) || this.containsDropped(node)) return null;
    this.used.add(id);
    const m = /\w+:paraId="([^"]+)"/.exec(s.xml);
    if (m) this.paraIds.add(m[1]);
    return s.xml;
  }
  private containsDropped(node: PMNode) {
    if (!this.drop.size) return false;
    let hit = false;
    node.descendants((n) => { if (this.drop.has(n)) hit = true; return !hit; });
    return hit;
  }

  /** Revision ids must be unique, but a split paragraph or run can repeat one. */
  private revAttrs(rev: Pick<Revision, 'author' | 'date' | 'raw'>): string {
    const w = this.w;
    if (!rev.raw) {
      const id = String(this.d.nextRevId());
      this.revIds.add(id);
      return ` ${w}id="${id}" ${w}author="${escAttr(rev.author)}" ${w}date="${escAttr(rev.date)}"`;
    }
    const id = /\bid="([^"]+)"/.exec(rev.raw)?.[1] ?? '';
    if (id && !this.revIds.has(id)) { this.revIds.add(id); return rev.raw; }
    const fresh = String(this.d.nextRevId());
    this.revIds.add(fresh);
    return rev.raw.replace(/(\bid=")[^"]+"/, `$1${fresh}"`);
  }

  blocks(parent: PMNode): string {
    let out = '';
    parent.forEach((n) => { out += this.block(n); });
    return out;
  }
  block(n: PMNode): string {
    const v = this.verbatim(n);
    if (v !== null) return v;
    if (n.type === schema.nodes.paragraph) return this.paragraph(n);
    if (n.type === schema.nodes.table) return this.table(n);
    return n.attrs.xml;
  }

  private cleanAttrs(raw: string) {
    return raw.replace(/\s\w+:(paraId|textId)="([^"]+)"/g, (m, kind, v) => {
      if (kind === 'textId') return m;
      if (this.paraIds.has(v)) return '';
      this.paraIds.add(v);
      return m;
    });
  }

  paragraph(n: PMNode): string {
    const w = this.w, a = n.attrs;
    const items: string[] = [...(a.pExtra as string[])];
    if (a.pStyle) items.push(`<${w}pStyle ${w}val="${escAttr(a.pStyle)}"/>`);
    if (a.numId) items.push(`<${w}numPr><${w}ilvl ${w}val="${a.ilvl ?? 0}"/><${w}numId ${w}val="${escAttr(a.numId)}"/></${w}numPr>`);
    if (a.ind) items.push(`<${w}ind${Object.entries(a.ind as Record<string, string>).map(([k, v]) => ` ${w}${k}="${escAttr(v)}"`).join('')}/>`);
    if (a.jc) items.push(`<${w}jc ${w}val="${escAttr(a.jc)}"/>`);
    const pm = a.pmark as { ins?: Revision; del?: Revision } | null;
    if (pm) {
      const inner = (pm.ins ? `<${w}ins${this.revAttrs(pm.ins)}/>` : '') + (pm.del ? `<${w}del${this.revAttrs(pm.del)}/>` : '');
      const at = items.findIndex((x) => tagName(x) === 'rPr');
      if (at < 0) items.push(`<${w}rPr>${inner}</${w}rPr>`);
      else items[at] = items[at].replace(/\/?>/, (m) => (m === '/>' ? `>${inner}</${w}rPr>` : `>${inner}`));
    }
    const pPr = items.length ? `<${w}pPr>${ordered(items, PPR_ORDER)}</${w}pPr>` : '';
    return `<${w}p${this.cleanAttrs(a.rawAttrs)}>${pPr}${this.inline(n)}</${w}p>`;
  }

  private rPr(marks: readonly Mark[]): string {
    const w = this.w, items: string[] = [];
    let extras: string[] = [];
    for (const m of marks) {
      switch (m.type.name) {
        case 'bold': items.push(`<${w}b/>`); break;
        case 'italic': items.push(`<${w}i/>`); break;
        case 'strike': items.push(`<${w}strike/>`); break;
        case 'underline': items.push(m.attrs.raw ?? `<${w}u ${w}val="${escAttr(m.attrs.val)}"/>`); break;
        case 'vert': items.push(`<${w}vertAlign ${w}val="${escAttr(m.attrs.val)}"/>`); break;
        case 'color': items.push(m.attrs.raw ?? `<${w}color ${w}val="${escAttr(m.attrs.val)}"/>`); break;
        case 'highlight': items.push(`<${w}highlight ${w}val="${escAttr(m.attrs.val)}"/>`); break;
        case 'size': items.push(`<${w}sz ${w}val="${Number(m.attrs.val)}"/>`); break;
        case 'font': {
          const f = escAttr(m.attrs.name);
          items.push(m.attrs.raw ?? `<${w}rFonts ${w}ascii="${f}" ${w}hAnsi="${f}" ${w}eastAsia="${f}" ${w}cs="${f}"/>`);
          break;
        }
        case 'rstyle': items.push(`<${w}rStyle ${w}val="${escAttr(m.attrs.id)}"/>`); break;
        case 'rextra': extras = m.attrs.xml; break;
      }
    }
    const have = new Set(items.map(tagName));
    items.push(...extras.filter((x) => !have.has(tagName(x))));
    return items.length ? `<${w}rPr>${ordered(items, RPR_ORDER)}</${w}rPr>` : '';
  }

  private textRun(text: string, marks: readonly Mark[], deleted: boolean): string {
    const w = this.w;
    let inner = '';
    for (const part of text.split(/(\t|‑|­)/)) {
      if (!part) continue;
      if (part === '\t') inner += `<${w}tab/>`;
      else if (part === '‑') inner += `<${w}noBreakHyphen/>`;
      else if (part === '­') inner += `<${w}softHyphen/>`;
      else {
        const tag = deleted ? 'delText' : 't';
        inner += `<${w}${tag} xml:space="preserve">${escText(part)}</${w}${tag}>`;
      }
    }
    return `<${w}r>${this.rPr(marks)}${inner}</${w}r>`;
  }

  private atom(n: PMNode, deleted: boolean): string {
    const w = this.w;
    if (n.isText) return this.textRun(n.text!, n.marks, deleted);
    if (this.drop.has(n)) return '';
    if (n.type === schema.nodes.opaque_inline) {
      const x = n.attrs.xml as string;
      if (['bookmark', 'range'].includes(n.attrs.kind)) return x;
      // Text inside a deletion uses different element names, in both directions.
      return deleted
        ? x.replace(/(<\/?(?:[\w.-]+:)?)t(?=[\s>])/g, '$1delText').replace(/(<\/?(?:[\w.-]+:)?)instrText\b/g, '$1delInstrText')
        : x.replace(/(<\/?(?:[\w.-]+:)?)delText\b/g, '$1t').replace(/(<\/?(?:[\w.-]+:)?)delInstrText\b/g, '$1instrText');
    }
    if (n.type === schema.nodes.page_break) return `<${w}r>${this.rPr(n.marks)}${n.attrs.xml ?? `<${w}br ${w}type="page"/>`}</${w}r>`;
    return `<${w}r>${this.rPr(n.marks)}${n.attrs.xml ?? `<${w}br/>`}</${w}r>`;
  }

  /** Word nests these as hyperlink, then insertion, then deletion. */
  private group(nodes: PMNode[], level: number): string {
    const types = [schema.marks.link, schema.marks.tracked_ins, schema.marks.tracked_del];
    if (level === types.length) return nodes.map((n) => this.atom(n, !!schema.marks.tracked_del.isInSet(n.marks))).join('');
    const type = types[level];
    let out = '';
    for (let i = 0; i < nodes.length;) {
      const mark = type.isInSet(nodes[i].marks) ?? null;
      let j = i + 1;
      while (j < nodes.length) {
        const next = type.isInSet(nodes[j].marks) ?? null;
        if (mark ? !next || !mark.eq(next) : !!next) break;
        j++;
      }
      const inner = this.group(nodes.slice(i, j), level + 1);
      if (!mark) out += inner;
      else if (level === 0) out += this.hyperlink(mark, inner);
      else if (inner) {
        const tag = level === 1 ? 'ins' : 'del';
        out += `<${this.w}${tag}${this.revAttrs(mark.attrs as Revision)}>${inner}</${this.w}${tag}>`;
      }
      i = j;
    }
    return out;
  }

  private inline(p: PMNode): string {
    const nodes: PMNode[] = [];
    p.forEach((n) => nodes.push(n));
    return this.group(nodes, 0);
  }

  private hyperlink(m: Mark, inner: string) {
    const w = this.w;
    if (m.attrs.rawAttrs) return `<${w}hyperlink${m.attrs.rawAttrs}>${inner}</${w}hyperlink>`;
    const rId = this.d.linkRels.get(`${this.owner}\n${m.attrs.href}`);
    const p = this.rPrefix ?? 'r';
    const ns = this.rPrefix ? '' : ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    return `<${w}hyperlink${ns} ${p}:id="${rId}" ${w}history="1">${inner}</${w}hyperlink>`;
  }

  table(t: PMNode): string {
    const w = this.w, map = TableMap.get(t);
    const widths = columnWidths(t, map);
    const original = (t.attrs.gridRaw as string).match(/w:w="(\d+)"|\bw="(\d+)"/g)?.map((x) => +x.replace(/\D/g, ''));
    const sameGrid = original && original.length === widths.length && original.every((v, i) => v === widths[i]);
    const grid = sameGrid ? t.attrs.gridRaw : `<${w}tblGrid>${widths.map((x) => `<${w}gridCol ${w}w="${x}"/>`).join('')}</${w}tblGrid>`;
    let rows = '';
    let rowStart = 0;
    for (let r = 0; r < map.height; r++) {
      const row = t.child(r);
      const v = sameGrid ? this.verbatim(row) : null;
      if (v !== null) { rows += v; rowStart += row.nodeSize; continue; }
      let cells = '';
      for (let c = 0; c < map.width;) {
        const pos = map.map[r * map.width + c];
        const rect = map.findCell(pos);
        const cell = t.nodeAt(pos)!;
        const span = rect.right - rect.left;
        const width = widths.slice(rect.left, rect.right).reduce((a, b) => a + b, 0);
        if (rect.top === r) cells += this.cell(cell, width, sameGrid === true);
        else cells += this.mergedContinuation(cell, span, width);
        c = rect.right;
      }
      let trPr = row.attrs.trPr as string;
      const change = row.attrs.rowChange as { kind: string; rev: Revision } | null;
      if (change) {
        const el = `<${w}${change.kind}${this.revAttrs(change.rev)}/>`;
        const close = trPr.lastIndexOf(`</${w}trPr>`);
        trPr = close >= 0 ? trPr.slice(0, close) + el + trPr.slice(close) : `${trPr}<${w}trPr>${el}</${w}trPr>`;
      }
      rows += `<${w}tr${this.cleanAttrs(row.attrs.rawAttrs)}>${trPr}${cells}</${w}tr>`;
      rowStart += row.nodeSize;
    }
    return `<${w}tbl${t.attrs.rawAttrs}>${t.attrs.tblPr}${grid}${rows}</${w}tbl>`;
  }

  private tcPr(cell: PMNode, width: number, extra: string[]) {
    const w = this.w;
    const items = [`<${w}tcW ${w}w="${width}" ${w}type="dxa"/>`, ...cell.attrs.tcPr as string[], ...extra];
    if (cell.attrs.colspan > 1) items.push(`<${w}gridSpan ${w}val="${cell.attrs.colspan}"/>`);
    return `<${w}tcPr>${ordered(items, TCPR_ORDER)}</${w}tcPr>`;
  }
  private cell(cell: PMNode, width: number, allowVerbatim: boolean): string {
    const w = this.w;
    if (allowVerbatim) { const v = this.verbatim(cell); if (v !== null) return v; }
    const extra = cell.attrs.rowspan > 1 ? [`<${w}vMerge ${w}val="restart"/>`] : [];
    let inner = this.blocks(cell);
    if (cell.lastChild?.type !== schema.nodes.paragraph) inner += `<${w}p/>`;
    return `<${w}tc${cell.attrs.rawAttrs}>${this.tcPr(cell, width, extra)}${inner}</${w}tc>`;
  }
  private mergedContinuation(cell: PMNode, _span: number, width: number) {
    const w = this.w;
    return `<${w}tc>${this.tcPr(cell, width, [`<${w}vMerge/>`])}<${w}p/></${w}tc>`;
  }
}

function columnWidths(t: PMNode, map: TableMap): number[] {
  const out = new Array<number>(map.width).fill(0);
  for (let i = 0; i < map.map.length; i++) {
    const col = i % map.width;
    const cell = t.nodeAt(map.map[i])!;
    const rect = map.findCell(map.map[i]);
    const w = (cell.attrs.colwidth as number[] | null)?.[col - rect.left];
    if (w && !out[col]) out[col] = w;
  }
  const known = out.filter(Boolean);
  const fill = known.length ? Math.round(known.reduce((a, b) => a + b, 0) / known.length) : 2400;
  return out.map((x) => x || fill);
}

async function addNumbering(d: WordDoc) {
  const n = d.numbering;
  if (!n.addedAbstract.length && !n.addedNum.length) return;
  const rels = await d.pkg.rels(d.mainPath);
  const rel = rels.find((r) => r.type.endsWith('/numbering'));
  const w = d.w;
  const ns = w ? `xmlns:${w.slice(0, -1)}` : 'xmlns';
  let path: string;
  let xml: string;
  if (rel) {
    path = resolveTarget(d.mainPath, rel.target);
    xml = (await d.pkg.text(path))!;
  } else {
    path = 'word/numbering.xml';
    xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${w}numbering ${ns}="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></${w}numbering>`;
    await d.pkg.addRel(d.mainPath, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', 'numbering.xml');
    await d.pkg.ensureContentOverride('/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
  }
  // abstractNum must precede num, and num must precede numIdMacAtCleanup.
  const first = new RegExp(`<(?:[\\w.-]+:)?num\\s`).exec(xml);
  const close = xml.lastIndexOf('</');
  const anchorA = first ? first.index : (xml.search(/<(?:[\w.-]+:)?numIdMacAtCleanup/) >= 0 ? xml.search(/<(?:[\w.-]+:)?numIdMacAtCleanup/) : close);
  xml = xml.slice(0, anchorA) + n.addedAbstract.join('') + xml.slice(anchorA);
  const mac = xml.search(/<(?:[\w.-]+:)?numIdMacAtCleanup/);
  const anchorN = mac >= 0 ? mac : xml.lastIndexOf('</');
  xml = xml.slice(0, anchorN) + n.addedNum.join('') + xml.slice(anchorN);
  d.pkg.set(path, xml);
  n.addedAbstract = [];
  n.addedNum = [];
}

function esc(s: string) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/** Section properties live in raw XML, so settings changed in the app are applied to the final text. */
function finalizeSections(d: WordDoc, xml: string, doc: PMNode): string {
  const w = esc(d.w), W = d.w;
  const orig = d.originalSect;
  const margins = doc.attrs.pgMar as Margins | null;
  const marginsChanged = !!(margins && orig && JSON.stringify(margins) !== JSON.stringify(orig.margins));
  const titleChanged = !!(orig && doc.attrs.titlePg !== orig.titlePg);
  if (!marginsChanged && !titleChanged && !d.addedRefs.length) return xml;

  if (!new RegExp(`<${w}sectPr\\b`).test(xml)) {
    const p = d.page;
    const sect = `<${W}sectPr><${W}pgSz ${W}w="${p.width}" ${W}h="${p.height}"/><${W}pgMar ${W}top="${p.top}" ${W}right="${p.right}" ${W}bottom="${p.bottom}" ${W}left="${p.left}" ${W}header="${p.header}" ${W}footer="${p.footer}" ${W}gutter="0"/></${W}sectPr>`;
    xml = xml.replace(new RegExp(`</${w}body>`), `${sect}</${W}body>`);
  }
  if (marginsChanged) {
    xml = xml.replace(new RegExp(`<${w}pgMar\\b[^>]*/>`, 'g'), (tag) => {
      for (const [k, v] of Object.entries(margins!)) {
        const re = new RegExp(`(\\s${w}${k}=")[^"]*"`);
        tag = re.test(tag) ? tag.replace(re, `$1${v}"`) : tag.replace(/\s*\/>$/, ` ${W}${k}="${v}"/>`);
      }
      return tag;
    });
  }
  const section = new RegExp(`<${w}sectPr\\b[^>]*?/>|<${w}sectPr\\b[^>]*>[\\s\\S]*?</${w}sectPr>`, 'g');
  const expand = (span: string) => (span.endsWith('/>') && !span.includes(`</${W}sectPr>`) ? `${span.slice(0, -2)}></${W}sectPr>` : span);
  if (d.addedRefs.length) {
    xml = xml.replace(section, (raw) => {
      let span = expand(raw);
      for (const ref of d.addedRefs) {
        if (new RegExp(`<${w}${ref.tag}\\b[^>]*type="${ref.slot}"`).test(span)) continue;
        const p = d.rPrefix ?? 'r';
        const ns = d.rPrefix ? '' : ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
        const el = `<${W}${ref.tag}${ns} ${W}type="${ref.slot}" ${p}:id="${ref.rId}"/>`;
        span = span.replace('>', `>${el}`);
      }
      return span;
    });
  }
  if (titleChanged) {
    if (doc.attrs.titlePg) {
      let done = false;
      xml = xml.replace(section, (raw) => {
        if (done) return raw;
        done = true;
        const span = expand(raw);
        const after = new RegExp(`<${w}(textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\\b`).exec(span);
        const at = after ? after.index : span.lastIndexOf('</');
        return `${span.slice(0, at)}<${W}titlePg/>${span.slice(at)}`;
      });
    } else xml = xml.replace(new RegExp(`<${w}titlePg\\b[^>]*/>`, 'g'), '');
  }
  return xml;
}

export function documentXml(d: WordDoc, doc: PMNode): string {
  const writer = new Writer(d, orphanAtoms(doc), d.w, d.rPrefix, d.mainPath);
  return finalizeSections(d, d.head + writer.blocks(doc) + d.tail, doc);
}

const BEFORE_TRACK = new Set('writeProtection view zoom removePersonalInformation removeDateAndTime doNotDisplayPageBoundaries displayBackgroundShape printPostScriptOverText printFractionalCharacterWidth printFormsData embedTrueTypeFonts embedSystemFonts saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop hideSpellingErrors hideGrammaticalErrors activeWritingStyle proofState formsDesign attachedTemplate linkStyles stylePaneFormatFilter stylePaneSortMethod documentType mailMerge revisionView'.split(' '));

async function writeTrackSetting(d: WordDoc) {
  if (!d.settingsPath) return;
  const xml = await d.pkg.text(d.settingsPath);
  if (xml === null) return;
  const on = /<[\w.:-]*trackRevisions\b(?![^>]*val="(0|false)")/.test(xml);
  if (on === d.track) return;
  if (!d.track) { d.pkg.set(d.settingsPath, xml.replace(/<[\w.:-]*trackRevisions\b[^>]*\/>/, '')); return; }
  const root = parseXml(xml);
  const after = elements(root).find((e) => !BEFORE_TRACK.has(local(e.name)));
  const at = after ? after.start : root.innerEnd;
  const prefix = root.name.includes(':') ? `${root.name.split(':')[0]}:` : '';
  d.pkg.set(d.settingsPath, `${xml.slice(0, at)}<${prefix}trackRevisions/>${xml.slice(at)}`);
}

async function ensureLinks(d: WordDoc, owner: string, doc: PMNode) {
  const hrefs = new Set<string>();
  doc.descendants((n) => {
    for (const m of n.marks) if (m.type === schema.marks.link && !m.attrs.rawAttrs && m.attrs.href) hrefs.add(m.attrs.href);
  });
  const type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
  for (const h of hrefs) {
    const key = `${owner}\n${h}`;
    if (!d.linkRels.has(key)) d.linkRels.set(key, await d.pkg.addRel(owner, type, h, true));
  }
}

const isBlank = (doc: PMNode) => doc.childCount === 1 && doc.firstChild!.isTextblock && doc.firstChild!.content.size === 0;

async function writePart(d: WordDoc, part: HfPart) {
  if (part.path ? part.written?.eq(part.doc) : isBlank(part.doc)) return;
  await d.allocatePart(part);
  await ensureLinks(d, part.path!, part.doc);
  const writer = new Writer(d, orphanAtoms(part.doc), part.w, part.rPrefix, part.path!);
  d.pkg.set(part.path!, part.head + writer.blocks(part.doc) + part.tail);
  part.written = part.doc;
}

/** forOutput skips comment bookkeeping for throwaway copies made for PDF export. */
export async function exportDocx(d: WordDoc, doc: PMNode, forOutput = false): Promise<Uint8Array> {
  await ensureLinks(d, d.mainPath, doc);
  for (const part of d.hf) await writePart(d, part);
  await addNumbering(d);
  if (!forOutput) {
    await writeComments(d, doc);
    await writeTrackSetting(d);
  }
  d.pkg.set(d.mainPath, documentXml(d, doc));
  return d.pkg.generate();
}
