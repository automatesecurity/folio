import { Mark, Node as PMNode } from 'prosemirror-model';
import { fillPageFields, HfPart, PAGE_FIELD, PAGES_FIELD, WordDoc } from '../docx/document';
import { jcToCss, ListCounter, parseParaFragment, RunProps } from '../docx/styles';
import { effectiveNum } from '../editor/plugins';
import { HIGHLIGHTS, schema } from '../editor/schema';

/* eslint-disable @typescript-eslint/no-explicit-any */
const pt = (twips: number) => twips / 20;

const FAMILIES: Record<string, string> = { Carlito: 'Carlito', Caladea: 'Caladea', Sans: 'LiberationSans', Serif: 'LiberationSerif', Mono: 'LiberationMono' };
function familyFor(name?: string): keyof typeof FAMILIES {
  const n = (name ?? 'calibri').toLowerCase();
  if (n.includes('calibri') || n.includes('carlito') || n.includes('aptos')) return 'Carlito';
  if (n.includes('cambria') || n.includes('caladea')) return 'Caladea';
  if (/mono|courier|consol|code/.test(n)) return 'Mono';
  if (/times|georgia|garamond|serif|book|palatino|minion/.test(n) && !n.includes('sans')) return 'Serif';
  return 'Sans';
}

class Builder {
  used = new Set<string>();
  counter: ListCounter;
  images = new Map<string, string>();
  constructor(private d: WordDoc) { this.counter = new ListCounter(d.numbering); }

  private run(text: string, base: RunProps, marks: readonly Mark[]) {
    const r: RunProps = { ...base };
    const o: any = {};
    let vert: string | null = null, highlight: string | undefined, href: string | undefined;
    for (const m of marks) {
      switch (m.type.name) {
        case 'rstyle': Object.assign(r, this.d.styles.resolvedChar(m.attrs.id)); break;
      }
    }
    for (const m of marks) {
      switch (m.type.name) {
        case 'bold': r.bold = true; break;
        case 'italic': r.italic = true; break;
        case 'underline': r.underline = true; break;
        case 'strike': r.strike = true; break;
        case 'vert': vert = m.attrs.val; break;
        case 'color': if (/^[0-9a-f]{6}$/i.test(m.attrs.val)) r.color = m.attrs.val; break;
        case 'highlight': highlight = HIGHLIGHTS[m.attrs.val]; break;
        case 'font': r.font = m.attrs.name; break;
        case 'size': r.size = m.attrs.val; break;
        case 'link': href = m.attrs.href ?? undefined; break;
        case 'tracked_ins': r.underline = true; r.color = '1F6FB2'; break;
        case 'tracked_del': r.strike = true; r.color = 'B3261E'; break;
        case 'rextra': {
          const x = (m.attrs.xml as string[]).join('');
          if (/:caps\b(?![^>]*val="(0|false)")/.test(x)) r.caps = true;
          if (/:b\b[^>]*val="(0|false)"/.test(x)) r.bold = false;
          if (/:i\b[^>]*val="(0|false)"/.test(x)) r.italic = false;
          if (/:vanish\b(?![^>]*val="(0|false)")/.test(x)) r.hidden = true;
          break;
        }
      }
    }
    if (r.hidden) return null;
    const fam = familyFor(r.font);
    this.used.add(fam);
    o.text = (r.caps ? text.toUpperCase() : text).replace(/\t/g, '    ');
    o.font = FAMILIES[fam];
    o.fontSize = (r.size ?? 22) / 2;
    o.bold = !!r.bold;
    o.italics = !!r.italic;
    if (r.color && /^[0-9a-f]{6}$/i.test(r.color)) o.color = `#${r.color}`;
    const deco = [r.underline && 'underline', r.strike && 'lineThrough'].filter(Boolean);
    if (deco.length) o.decoration = deco;
    if (highlight) o.background = highlight;
    if (vert === 'superscript') o.sup = true;
    if (vert === 'subscript') o.sub = true;
    if (href && /^(https?:|mailto:)/i.test(href)) o.link = href;
    return o;
  }

  private paragraph(n: PMNode, out: any[], width: number, state: { breakNext: boolean }) {
    const d = this.d;
    const own = (n.attrs.pExtra as string[]).length ? parseParaFragment((n.attrs.pExtra as string[]).join('')) : {};
    const style = { ...d.styles.resolvedPara(n.attrs.pStyle), ...own };
    const base = d.styles.resolvedRun(n.attrs.pStyle);
    const num = effectiveNum(d, n);
    const marker = num ? this.counter.next(num.numId, num.ilvl) : null;
    const ind = n.attrs.ind as Record<string, string> | null;
    const left = Number(ind?.left ?? ind?.start ?? marker?.level.left ?? style.left ?? 0);
    const right = Number(ind?.right ?? ind?.end ?? style.right ?? 0);
    const hanging = ind?.hanging !== undefined ? +ind.hanging : marker?.level.first !== undefined ? -marker.level.first : style.first !== undefined && style.first < 0 ? -style.first : 0;
    const first = ind?.firstLine !== undefined ? +ind.firstLine : (style.first ?? 0) > 0 ? style.first! : 0;
    const alignment = jcToCss(n.attrs.jc ?? style.jc);

    const segments: any[][] = [[]];
    const blocks: any[] = [];
    let breakAfter = false;
    n.forEach((c) => {
      if (c.isText) { const r = this.run(c.text!, base, c.marks); if (r) segments[segments.length - 1].push(r); return; }
      if (c.type === schema.nodes.hard_break) { const r = this.run('\n', base, c.marks); if (r) segments[segments.length - 1].push(r); return; }
      if (c.type === schema.nodes.page_break) { breakAfter = true; return; }
      const a = c.attrs;
      if (a.kind === 'drawing') {
        const data = a.img ? this.images.get(a.img) : undefined;
        if (data) {
          const scale = Math.min(1, width / ((a.w || 100) * 0.75));
          blocks.push({ at: segments.length, image: data, width: (a.w || 100) * 0.75 * scale, height: (a.h || 100) * 0.75 * scale });
          segments.push([]);
        } else if (a.label) { const r = this.run(`[${a.label}]`, base, []); if (r) segments[segments.length - 1].push(r); }
        return;
      }
      if (a.kind === 'del') return;
      if (a.kind === 'note') { const r = this.run(a.label, base, []); if (r) segments[segments.length - 1].push({ ...r, sup: true }); return; }
      const label = a.instr === 'PAGE' ? PAGE_FIELD : a.instr === 'NUMPAGES' ? PAGES_FIELD : a.label;
      if (label && a.kind !== 'comment') { const r = this.run(label, base, []); if (r) segments[segments.length - 1].push(r); }
    });

    const lineHeight = style.line && (style.lineRule === 'auto' || !style.lineRule) ? style.line / 240 : 1;
    const before = pt(style.before ?? 0), after = pt(style.after ?? 0);
    let isFirst = true;
    segments.forEach((runs, i) => {
      const image = blocks.find((b) => b.at === i);
      if (runs.length || (i === 0 && !blocks.length)) {
        const fallback = this.run(' ', base, []);
        let node: any = { text: runs.length ? runs : [fallback ?? { text: ' ' }], alignment, lineHeight, leadingIndent: pt(first) };
        if (marker && isFirst) {
          const r = this.run(marker.label || ' ', base, []);
          node = { columns: [{ width: pt(hanging), text: [r], lineHeight }, { width: '*', stack: [node] }], columnGap: 0 };
          node.margin = [pt(left - hanging), isFirst ? before : 0, pt(right), 0];
        } else node.margin = [pt(left), isFirst ? before : 0, pt(right), 0];
        if (isFirst && state.breakNext) { node.pageBreak = 'before'; state.breakNext = false; }
        if (isFirst && style.pageBreakBefore) node.pageBreak = 'before';
        out.push(node);
        isFirst = false;
      }
      if (image) {
        const node: any = { image: image.image, width: image.width, height: image.height, alignment, margin: [pt(left), isFirst ? before : 0, pt(right), 0] };
        if (isFirst && state.breakNext) { node.pageBreak = 'before'; state.breakNext = false; }
        out.push(node);
        isFirst = false;
      }
    });
    if (after) out.push({ text: '', margin: [0, 0, 0, after], fontSize: 0.1, lineHeight: 0 });
    if (breakAfter) state.breakNext = true;
  }

  private table(t: PMNode, out: any[], state: { breakNext: boolean }) {
    let cols = 0;
    t.firstChild!.forEach((c) => { cols += c.attrs.colspan; });
    const widths = new Array<number | undefined>(cols).fill(undefined);
    t.forEach((row) => {
      let col = 0;
      row.forEach((cell) => {
        ((cell.attrs.colwidth as number[] | null) ?? []).forEach((w, i) => { if (w) widths[col + i] ??= pt(w); });
        col += cell.attrs.colspan;
      });
    });
    const body: any[][] = Array.from({ length: t.childCount }, () => new Array(cols).fill(null));
    t.forEach((row, _o, r) => {
      let col = 0;
      row.forEach((cell) => {
        while (body[r][col] !== null) col++;
        const stack: any[] = [];
        const avail = (cell.attrs.colwidth as number[] | null)?.reduce((a, b) => a + pt(b), 0) || pt(this.d.page.width - this.d.page.left - this.d.page.right) / cols;
        this.blocks(cell, stack, avail - 12, { breakNext: false });
        const c: any = { stack, colSpan: cell.attrs.colspan, rowSpan: cell.attrs.rowspan };
        const fill = /fill="([0-9a-fA-F]{6})"/.exec((cell.attrs.tcPr as string[]).join(''));
        if (fill) c.fillColor = `#${fill[1]}`;
        for (let i = 0; i < cell.attrs.rowspan; i++) for (let j = 0; j < cell.attrs.colspan; j++) if (r + i < body.length) body[r + i][col + j] = {};
        body[r][col] = c;
        col += cell.attrs.colspan;
      });
    });
    const grid = /:tblBorders\b/.test(t.attrs.tblPr) || /:tblStyle\b[^>]*val="[^"]*Grid/.test(t.attrs.tblPr);
    const node: any = {
      table: { widths: widths.map((w) => w ?? '*'), body },
      layout: {
        hLineWidth: () => (grid ? 0.5 : 0), vLineWidth: () => (grid ? 0.5 : 0),
        paddingLeft: () => 5.4, paddingRight: () => 5.4, paddingTop: () => 0, paddingBottom: () => 0,
      },
      margin: [0, 0, 0, 6],
    };
    if (state.breakNext) { node.pageBreak = 'before'; state.breakNext = false; }
    out.push(node);
  }

  blocks(parent: PMNode, out: any[], width: number, state: { breakNext: boolean }) {
    parent.forEach((n) => {
      if (n.type === schema.nodes.paragraph) this.paragraph(n, out, width, state);
      else if (n.type === schema.nodes.table) this.table(n, out, state);
      else if (n.type === schema.nodes.opaque_block && n.attrs.kind !== 'bookmark') {
        for (const p of n.attrs.preview as { text: string; cls: string | null }[]) {
          const props = this.d.styles.resolvedPara(p.cls);
          const r = this.run(p.text || ' ', this.d.styles.resolvedRun(p.cls), []);
          if (r) out.push({ text: [r], margin: [pt(props.left ?? 0), pt(props.before ?? 0), 0, pt(props.after ?? 0)] });
        }
      }
    });
  }
}

async function toDataUrl(bytes: Uint8Array, ext: string): Promise<string | null> {
  const b64 = (x: Uint8Array) => { let s = ''; for (let i = 0; i < x.length; i += 8192) s += String.fromCharCode(...x.subarray(i, i + 8192)); return btoa(s); };
  if (ext === 'png') return `data:image/png;base64,${b64(bytes)}`;
  if (ext === 'jpg' || ext === 'jpeg') return `data:image/jpeg;base64,${b64(bytes)}`;
  try {
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: ext === 'svg' ? 'image/svg+xml' : `image/${ext}` }));
    const img = new Image();
    img.src = url;
    await img.decode();
    const c = document.createElement('canvas');
    c.width = img.naturalWidth || 300; c.height = img.naturalHeight || 150;
    c.getContext('2d')!.drawImage(img, 0, 0);
    URL.revokeObjectURL(url);
    return c.toDataURL('image/png');
  } catch { return null; }
}

async function fontBase64(file: string) {
  const res = await fetch(`/fonts/${file}`);
  if (!res.ok) throw new Error(`Bundled font ${file} is missing.`);
  const bytes = new Uint8Array(await res.arrayBuffer());
  let s = '';
  for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return btoa(s);
}

const fillClone = (n: any, page: number, pages: number): any => {
  if (typeof n === 'string') return fillPageFields(n, page, pages);
  if (Array.isArray(n)) return n.map((x) => fillClone(x, page, pages));
  if (n && typeof n === 'object') return Object.fromEntries(Object.entries(n).map(([k, v]) => [k, k === 'image' ? v : fillClone(v, page, pages)]));
  return n;
};

export async function exportPdf(d: WordDoc, doc: PMNode): Promise<Uint8Array> {
  const b = new Builder(d);
  const images = new Set<string>();
  const collect = (root: PMNode) => root.descendants((n) => { if (n.type === schema.nodes.opaque_inline && n.attrs.img) images.add(n.attrs.img); });
  collect(doc);
  d.hf.forEach((part) => collect(part.doc));
  for (const id of images) {
    const im = d.images.get(id);
    const url = im && await toDataUrl(im.bytes, im.ext);
    if (url) b.images.set(id, url);
  }
  const p = d.page;
  const width = pt(p.width - p.left - p.right);
  const content: any[] = [];
  b.blocks(doc, content, width, { breakNext: false });
  const bands = new Map<HfPart, any[]>();
  for (const part of d.hf) {
    const hb = new Builder(d);
    hb.images = b.images;
    const out: any[] = [];
    hb.blocks(part.doc, out, width, { breakNext: false });
    hb.used.forEach((u) => b.used.add(u));
    bands.set(part, out);
  }
  const titlePg = doc.attrs.titlePg === true;
  const band = (kind: HfPart['kind']) => (page: number, pages: number) => {
    const part = d.findPart(kind, d.slotFor(page - 1, titlePg));
    if (!part || !part.doc.textContent.trim() && !bands.get(part)?.some((x) => x.image)) return null;
    const stack = fillClone(bands.get(part), page, pages);
    if (kind === 'header') return { stack, margin: [pt(p.left), pt(p.header), pt(p.right), 0] };
    const height = d.hfBottom > p.footer ? pt(d.hfBottom - p.footer) : 14;
    return { stack, margin: [pt(p.left), Math.max(0, pt(d.effBottom - p.footer) - height), pt(p.right), 0] };
  };
  b.used.add(familyFor(d.styles.resolvedRun(null).font));
  const def: any = {
    pageSize: { width: pt(p.width), height: pt(p.height) },
    pageMargins: [pt(p.left), pt(d.effTop), pt(p.right), pt(d.effBottom)],
    defaultStyle: { font: FAMILIES[familyFor(d.styles.resolvedRun(null).font)] },
    content, header: band('header'), footer: band('footer'),
  };
  const mod: any = await import('pdfmake/build/pdfmake.js');
  const pdfMake = mod.default ?? mod;
  const vfs: Record<string, { data: string; encoding: string }> = {};
  const fonts: Record<string, any> = {};
  await Promise.all([...b.used].map(async (key) => {
    const fam = FAMILIES[key];
    const f: Record<string, string> = {};
    await Promise.all(([['normal', 'Regular'], ['bold', 'Bold'], ['italics', 'Italic'], ['bolditalics', 'BoldItalic']] as const).map(async ([slot, name]) => {
      const file = `${fam}-${name}.ttf`;
      vfs[file] = { data: await fontBase64(file), encoding: 'base64' };
      f[slot] = file;
    }));
    fonts[fam] = f;
  }));
  pdfMake.addVirtualFileSystem(vfs);
  pdfMake.addFonts(fonts);
  const buffer = await pdfMake.createPdf(def).getBuffer();
  return new Uint8Array(buffer);
}
