import { attr, child, elements, local, parseXml, walk, XEl } from './xml';

export interface RunProps {
  font?: string; size?: number; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean;
  color?: string; caps?: boolean; smallCaps?: boolean; shade?: string; hidden?: boolean;
}
export interface ParaProps {
  before?: number; after?: number; line?: number; lineRule?: string; left?: number; right?: number; first?: number;
  jc?: string; keepNext?: boolean; pageBreakBefore?: boolean; shade?: string; numId?: string; ilvl?: number;
  borderBottom?: string; borderTop?: string; contextual?: boolean;
}
export interface StyleDef {
  id: string; name: string; type: string; basedOn?: string; next?: string; qFormat: boolean; hidden: boolean;
  priority: number; isDefault: boolean; run: RunProps; para: ParaProps;
}

const safeColor = (v?: string) => (v && /^[0-9a-fA-F]{6}$/.test(v) ? `#${v}` : undefined);
export const cssIdent = (s: string) => s.replace(/[^a-zA-Z0-9_-]/g, (c) => `_${c.charCodeAt(0).toString(16)}`);
const on = (el: XEl) => { const v = attr(el, 'val'); return v === undefined || !['0', 'false', 'off'].includes(v); };

const FONT_FALLBACK: Record<string, string> = {
  calibri: 'Carlito', 'calibri light': 'Carlito', cambria: 'Caladea', 'times new roman': 'Liberation Serif', times: 'Liberation Serif',
  arial: 'Liberation Sans', helvetica: 'Liberation Sans', 'courier new': 'Liberation Mono', courier: 'Liberation Mono',
};
export function fontStack(name?: string): string {
  const clean = (name ?? '').replace(/["'\;{}<>()]/g, '').trim();
  if (!clean) return 'Carlito, sans-serif';
  const sub = FONT_FALLBACK[clean.toLowerCase()];
  const generic = /mono|courier|consol/i.test(clean) ? 'monospace' : /times|serif|georgia|garamond|cambria|book/i.test(clean) && !/sans/i.test(clean) ? 'serif' : 'sans-serif';
  return `"${clean}"${sub ? `, "${sub}"` : ''}, ${generic}`;
}

export interface Theme { major?: string; minor?: string; colors: Record<string, string> }

export function parseTheme(xml: string | null): Theme {
  const t: Theme = { colors: {} };
  if (!xml) return t;
  const root = parseXml(xml);
  for (const e of walk(root)) {
    const l = local(e.name);
    if (l === 'majorFont') t.major = attr(child(e, 'latin'), 'typeface');
    if (l === 'minorFont') t.minor = attr(child(e, 'latin'), 'typeface');
    if (l === 'clrScheme') for (const c of elements(e)) {
      const v = elements(c)[0];
      const val = v?.attrs.lastClr ?? v?.attrs.val;
      if (val) t.colors[local(c.name)] = val;
    }
  }
  return t;
}

function readRun(rPr: XEl | undefined, theme: Theme, into: RunProps = {}): RunProps {
  if (!rPr) return into;
  for (const e of elements(rPr)) {
    switch (local(e.name)) {
      case 'rFonts': {
        const th = attr(e, 'asciiTheme');
        into.font = th ? (th.startsWith('major') ? theme.major : theme.minor) ?? attr(e, 'ascii') : attr(e, 'ascii') ?? into.font;
        break;
      }
      case 'sz': { const v = Number(attr(e, 'val')); if (v > 0) into.size = v; break; }
      case 'b': into.bold = on(e); break;
      case 'i': into.italic = on(e); break;
      case 'strike': into.strike = on(e); break;
      case 'caps': into.caps = on(e); break;
      case 'smallCaps': into.smallCaps = on(e); break;
      case 'vanish': into.hidden = on(e); break;
      case 'u': into.underline = attr(e, 'val') !== 'none'; break;
      case 'color': {
        const th = attr(e, 'themeColor');
        const v = attr(e, 'val');
        into.color = (v && v !== 'auto' ? v : undefined) ?? (th ? theme.colors[themeKey(th)] : undefined) ?? into.color;
        break;
      }
      case 'shd': { const f = attr(e, 'fill'); if (f && f !== 'auto') into.shade = f; break; }
    }
  }
  return into;
}
const themeKey = (t: string) => ({ text1: 'dk1', text2: 'dk2', background1: 'lt1', background2: 'lt2' } as Record<string, string>)[t] ?? t;

function readPara(pPr: XEl | undefined, into: ParaProps = {}): ParaProps {
  if (!pPr) return into;
  for (const e of elements(pPr)) {
    switch (local(e.name)) {
      case 'spacing': {
        const b = attr(e, 'before'), a = attr(e, 'after'), l = attr(e, 'line');
        if (b !== undefined) into.before = +b;
        if (a !== undefined) into.after = +a;
        if (l !== undefined) { into.line = +l; into.lineRule = attr(e, 'lineRule') ?? 'auto'; }
        break;
      }
      case 'ind': {
        const l = attr(e, 'left') ?? attr(e, 'start'), r = attr(e, 'right') ?? attr(e, 'end'), f = attr(e, 'firstLine'), h = attr(e, 'hanging');
        if (l !== undefined) into.left = +l;
        if (r !== undefined) into.right = +r;
        if (f !== undefined) into.first = +f;
        if (h !== undefined) into.first = -h;
        break;
      }
      case 'jc': into.jc = attr(e, 'val'); break;
      case 'keepNext': into.keepNext = on(e); break;
      case 'pageBreakBefore': into.pageBreakBefore = on(e); break;
      case 'contextualSpacing': into.contextual = on(e); break;
      case 'shd': { const f = attr(e, 'fill'); if (f && f !== 'auto') into.shade = f; break; }
      case 'numPr': {
        const n = attr(child(e, 'numId'), 'val'), il = attr(child(e, 'ilvl'), 'val');
        if (n) into.numId = n;
        if (il) into.ilvl = +il;
        break;
      }
      case 'pBdr': {
        const fmt = (b?: XEl) => b && attr(b, 'val') !== 'none' && attr(b, 'val') !== 'nil'
          ? `${Math.max(1, (Number(attr(b, 'sz')) || 4) / 6)}px solid ${safeColor(attr(b, 'color')) ?? '#000'}` : undefined;
        into.borderBottom = fmt(child(e, 'bottom'));
        into.borderTop = fmt(child(e, 'top'));
        break;
      }
    }
  }
  return into;
}
export const parseParaFragment = (raw: string): ParaProps => readPara(parseXml(`<x xmlns:w="w">${raw}</x>`));

export const jcToCss = (jc?: string) =>
  jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : jc === 'both' || jc === 'distribute' ? 'justify' : 'left';
export const tw = (n: number) => n / 15; // twips to CSS px

export function paraCss(p: ParaProps): string {
  const s: string[] = [];
  if (p.before !== undefined) s.push(`padding-top:${tw(p.before)}px`);
  if (p.after !== undefined) s.push(`padding-bottom:${tw(p.after)}px`);
  if (p.line !== undefined) {
    if (p.lineRule === 'auto' || !p.lineRule) s.push(`line-height:${((p.line / 240) * 1.2).toFixed(3)}`);
    else s.push(`line-height:${tw(p.line)}px`);
  }
  if (p.left !== undefined) s.push(`margin-left:${tw(p.left)}px`);
  if (p.right !== undefined) s.push(`margin-right:${tw(p.right)}px`);
  if (p.first !== undefined) s.push(`text-indent:${tw(p.first)}px`);
  if (p.jc) s.push(`text-align:${jcToCss(p.jc)}`);
  const shade = safeColor(p.shade);
  if (shade) s.push(`background:${shade}`);
  if (p.borderBottom) s.push(`border-bottom:${p.borderBottom}`);
  if (p.borderTop) s.push(`border-top:${p.borderTop}`);
  return s.join(';');
}
export function runCss(r: RunProps): string {
  const s: string[] = [];
  if (r.font) s.push(`font-family:${fontStack(r.font)}`);
  if (r.size) s.push(`font-size:${r.size / 2}pt`);
  if (r.bold !== undefined) s.push(`font-weight:${r.bold ? 700 : 400}`);
  if (r.italic !== undefined) s.push(`font-style:${r.italic ? 'italic' : 'normal'}`);
  const color = safeColor(r.color);
  if (color) s.push(`color:${color}`);
  const deco = [r.underline && 'underline', r.strike && 'line-through'].filter(Boolean).join(' ');
  if (deco) s.push(`text-decoration:${deco}`);
  if (r.caps) s.push('text-transform:uppercase');
  if (r.smallCaps) s.push('font-variant:small-caps');
  const shade = safeColor(r.shade);
  if (shade) s.push(`background:${shade}`);
  return s.join(';');
}

export class StyleSheet {
  styles = new Map<string, StyleDef>();
  defaults: { run: RunProps; para: ParaProps } = { run: {}, para: {} };
  defaultParagraph?: string;
  constructor(xml: string | null, public theme: Theme) {
    if (!xml) return;
    const root = parseXml(xml);
    const dd = child(root, 'docDefaults');
    if (dd) {
      readRun(child(child(dd, 'rPrDefault') ?? dd, 'rPr'), theme, this.defaults.run);
      readPara(child(child(dd, 'pPrDefault') ?? dd, 'pPr'), this.defaults.para);
    }
    for (const e of elements(root)) {
      if (local(e.name) !== 'style') continue;
      const id = attr(e, 'styleId');
      if (!id) continue;
      const isDefault = ['1', 'true', 'on'].includes(attr(e, 'default') ?? '');
      const type = attr(e, 'type') ?? 'paragraph';
      const def: StyleDef = {
        id, type, isDefault, name: attr(child(e, 'name'), 'val') ?? id, basedOn: attr(child(e, 'basedOn'), 'val'),
        next: attr(child(e, 'next'), 'val'), qFormat: !!child(e, 'qFormat'),
        hidden: !!child(e, 'semiHidden') || !!child(e, 'hidden'), priority: Number(attr(child(e, 'uiPriority'), 'val') ?? 99),
        run: readRun(child(e, 'rPr'), theme), para: readPara(child(e, 'pPr')),
      };
      this.styles.set(id, def);
      if (isDefault && type === 'paragraph') this.defaultParagraph = id;
    }
  }
  private chain(id: string | null | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let cur = id ? this.styles.get(id) : undefined;
    while (cur && !seen.has(cur.id)) { out.unshift(cur); seen.add(cur.id); cur = cur.basedOn ? this.styles.get(cur.basedOn) : undefined; }
    return out;
  }
  resolvedPara(id?: string | null): ParaProps {
    const sid = id ?? this.defaultParagraph;
    return Object.assign({}, this.defaults.para, ...this.chain(sid).map((s) => s.para));
  }
  resolvedRun(id?: string | null): RunProps {
    const sid = id ?? this.defaultParagraph;
    return Object.assign({}, this.defaults.run, ...this.chain(sid).map((s) => s.run));
  }
  resolvedChar(id?: string | null): RunProps {
    return Object.assign({}, ...this.chain(id).map((s) => s.run));
  }
  paragraphChoices(): StyleDef[] {
    return [...this.styles.values()]
      .filter((s) => s.type === 'paragraph' && !s.hidden && (s.qFormat || s.isDefault))
      .sort((a, b) => a.priority - b.priority);
  }
  css(): string {
    const lines: string[] = [];
    const base = { ...this.defaults.run, ...this.resolvedRun(null) };
    lines.push(`.pm-doc{${runCss({ font: 'Calibri', size: 22, ...base })};${paraCss(this.resolvedPara(null))}}`.replace(/padding-[a-z]+:[^;]+;?/g, '').replace(/margin-[a-z]+:[^;]+;?/g, '').replace(/text-indent:[^;]+;?/g, ''));
    for (const s of this.styles.values()) {
      if (s.type === 'paragraph') {
        lines.push(`.pm-doc .ps-${cssIdent(s.id)}{${paraCss(this.resolvedPara(s.id))};${runCss(this.resolvedRun(s.id))}}`);
      } else if (s.type === 'character') {
        lines.push(`.pm-doc .cs-${cssIdent(s.id)}{${runCss(this.resolvedChar(s.id))}}`);
      }
    }
    lines.push(`.pm-doc p:not([class*="ps-"]){${paraCss(this.resolvedPara(null))}}`);
    return lines.join('\n');
  }
}

// Numbering ---------------------------------------------------------------

export interface Level { start: number; fmt: string; text: string; left?: number; first?: number; font?: string }
export interface AbstractNum { id: string; levels: Map<number, Level> }

export class Numbering {
  abstracts = new Map<string, AbstractNum>();
  nums = new Map<string, { abstractId: string; overrides: Map<number, number> }>();
  addedAbstract: string[] = [];
  addedNum: string[] = [];
  constructor(xml: string | null) {
    if (!xml) return;
    this.read(parseXml(xml));
  }
  private read(root: XEl) {
    for (const e of elements(root)) {
      const l = local(e.name);
      if (l === 'abstractNum') {
        const id = attr(e, 'abstractNumId')!;
        const levels = new Map<number, Level>();
        for (const lv of elements(e)) {
          if (local(lv.name) !== 'lvl') continue;
          const ind = child(child(lv, 'pPr') ?? lv, 'ind');
          const left = attr(ind, 'left') ?? attr(ind, 'start'), hanging = attr(ind, 'hanging'), first = attr(ind, 'firstLine');
          levels.set(+(attr(lv, 'ilvl') ?? 0), {
            start: +(attr(child(lv, 'start'), 'val') ?? 1), fmt: attr(child(lv, 'numFmt'), 'val') ?? 'decimal',
            text: attr(child(lv, 'lvlText'), 'val') ?? '', left: left !== undefined ? +left : undefined,
            first: hanging !== undefined ? -hanging : first !== undefined ? +first : undefined,
            font: attr(child(child(lv, 'rPr') ?? lv, 'rFonts'), 'ascii'),
          });
        }
        this.abstracts.set(id, { id, levels });
      } else if (l === 'num') {
        const overrides = new Map<number, number>();
        for (const o of elements(e)) {
          if (local(o.name) !== 'lvlOverride') continue;
          const s = attr(child(o, 'startOverride'), 'val');
          if (s !== undefined) overrides.set(+(attr(o, 'ilvl') ?? 0), +s);
        }
        this.nums.set(attr(e, 'numId')!, { abstractId: attr(child(e, 'abstractNumId'), 'val') ?? '', overrides });
      }
    }
  }
  level(numId: string, ilvl: number): Level | undefined {
    const n = this.nums.get(numId);
    return n ? this.abstracts.get(n.abstractId)?.levels.get(ilvl) : undefined;
  }
  kindOf(numId: string | null | undefined): 'bullet' | 'ordered' | null {
    if (!numId) return null;
    const l = this.level(numId, 0);
    return l ? (l.fmt === 'bullet' ? 'bullet' : l.fmt === 'none' ? null : 'ordered') : null;
  }
  private nextIds() {
    const a = Math.max(-1, ...[...this.abstracts.keys()].map(Number)) + 1;
    const n = Math.max(0, ...[...this.nums.keys()].map(Number)) + 1;
    return { a, n };
  }
  /** Reuses a bullet definition, gives each new numbered list its own counter. */
  newList(kind: 'bullet' | 'ordered', w: string, ns: boolean): string {
    if (kind === 'bullet') {
      for (const [id] of this.nums) if (this.kindOf(id) === 'bullet') return id;
    }
    const { a, n } = this.nextIds();
    let abstractId: string | undefined;
    if (kind === 'ordered') for (const [id, v] of this.nums) if (this.kindOf(id) === 'ordered' && v.overrides.size === 0) abstractId = v.abstractId;
    if (abstractId === undefined) {
      abstractId = String(a);
      const levels = new Map<number, Level>();
      const xml = abstractXml(abstractId, kind, w, levels);
      this.abstracts.set(abstractId, { id: abstractId, levels });
      this.addedAbstract.push(xml);
    }
    const overrides = new Map<number, number>();
    let inner = `<${w}abstractNumId ${w}val="${abstractId}"/>`;
    if (kind === 'ordered') {
      overrides.set(0, 1);
      inner += `<${w}lvlOverride ${w}ilvl="0"><${w}startOverride ${w}val="1"/></${w}lvlOverride>`;
    }
    this.nums.set(String(n), { abstractId, overrides });
    this.addedNum.push(`<${w}num ${w}numId="${n}">${inner}</${w}num>`);
    return String(n);
  }
}

function abstractXml(id: string, kind: 'bullet' | 'ordered', w: string, levels: Map<number, Level>) {
  const fmts = ['decimal', 'lowerLetter', 'lowerRoman'];
  const bullets = ['\u2022', '\u25E6', '\u25AA'];
  let out = `<${w}abstractNum ${w}abstractNumId="${id}"><${w}multiLevelType ${w}val="hybridMultilevel"/>`;
  for (let i = 0; i < 9; i++) {
    const fmt = kind === 'bullet' ? 'bullet' : fmts[i % 3];
    const text = kind === 'bullet' ? bullets[i % 3] : `%${i + 1}.`;
    const left = 720 * (i + 1);
    levels.set(i, { start: 1, fmt, text, left, first: -360 });
    out += `<${w}lvl ${w}ilvl="${i}"><${w}start ${w}val="1"/><${w}numFmt ${w}val="${fmt}"/><${w}lvlText ${w}val="${text}"/><${w}lvlJc ${w}val="left"/><${w}pPr><${w}ind ${w}left="${left}" ${w}hanging="360"/></${w}pPr></${w}lvl>`;
  }
  return out + `</${w}abstractNum>`;
}

const roman = (n: number) => {
  const map: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let o = '';
  for (const [v, s] of map) while (n >= v) { o += s; n -= v; }
  return o;
};
const letters = (n: number) => { let s = ''; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };
export function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case 'lowerLetter': return letters(n);
    case 'upperLetter': return letters(n).toUpperCase();
    case 'lowerRoman': return roman(n);
    case 'upperRoman': return roman(n).toUpperCase();
    case 'decimalZero': return String(n).padStart(2, '0');
    default: return String(n);
  }
}
export function bulletChar(l: Level): string {
  const c = l.text.codePointAt(0) ?? 0x2022;
  if (c >= 0xf000 && c <= 0xf0ff) {
    const low = c & 0xff;
    if (/symbol/i.test(l.font ?? '')) return low === 0xb7 ? '\u2022' : '\u2022';
    if (/wingdings/i.test(l.font ?? '')) return ({ 0xa7: '\u25AA', 0xd8: '\u27A2', 0xfc: '\u2713', 0x71: '\u2751', 0x76: '\u2756' } as Record<number, string>)[low] ?? '\u25AA';
    return '\u2022';
  }
  return l.text === 'o' && /courier/i.test(l.font ?? '') ? '\u25E6' : l.text || '\u2022';
}

/** Walks paragraphs in order and hands out the marker text Word would show. */
export class ListCounter {
  private counters = new Map<string, number[]>();
  private started = new Set<string>();
  constructor(private numbering: Numbering) {}
  next(numId: string, ilvl: number): { label: string; level: Level } | null {
    const lvl = this.numbering.level(numId, ilvl);
    const num = this.numbering.nums.get(numId);
    if (!lvl || !num || lvl.fmt === 'none') return lvl ? { label: '', level: lvl } : null;
    const key = num.abstractId;
    const c = this.counters.get(key) ?? [];
    this.counters.set(key, c);
    const ov = num.overrides.get(ilvl);
    const sKey = `${numId}:${ilvl}`;
    if (ov !== undefined && !this.started.has(sKey)) { c[ilvl] = ov; this.started.add(sKey); }
    else c[ilvl] = c[ilvl] === undefined ? lvl.start : c[ilvl] + 1;
    c.length = ilvl + 1;
    if (lvl.fmt === 'bullet') return { label: bulletChar(lvl), level: lvl };
    for (let i = 0; i < ilvl; i++) if (c[i] === undefined) c[i] = this.numbering.level(numId, i)?.start ?? 1;
    const label = lvl.text.replace(/%(\d)/g, (_, d) => {
      const i = +d - 1;
      return formatNumber(c[i] ?? 1, this.numbering.level(numId, i)?.fmt ?? 'decimal');
    });
    return { label, level: lvl };
  }
}
