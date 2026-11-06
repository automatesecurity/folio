import { Node as PMNode } from 'prosemirror-model';
import { schema } from '../editor/schema';
import type { WordDoc } from './document';
import { resolveTarget } from './package';
import { attr, elements, escAttr, escText, local, parseXml, textOf, walk, XEl } from './xml';

export interface Comment {
  id: string;
  author: string;
  initials: string;
  date: string;
  text: string;
  /** Original XML, kept until the comment is edited. */
  xml: string | null;
  paraId: string | null;
  parent: string | null;
  done: boolean;
  /** True if the body held this comment's anchors when it was opened or created. */
  anchored: boolean;
  dirty: boolean;
}

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export async function readComments(d: WordDoc) {
  const rels = await d.pkg.rels(d.mainPath);
  const rel = rels.find((r) => r.type.endsWith('/comments') && !r.external);
  if (!rel) return;
  d.commentsPath = resolveTarget(d.mainPath, rel.target);
  const xml = await d.pkg.text(d.commentsPath);
  if (!xml) return;
  const root = parseXml(xml);
  const bodyXml = d.bodySource();
  const list = elements(root).filter((e) => local(e.name) === 'comment');
  d.commentsHead = list.length ? xml.slice(0, list[0].start) : xml.slice(0, root.innerStart);
  d.commentsTail = list.length ? xml.slice(list[list.length - 1].end) : xml.slice(root.innerEnd);
  for (const c of list) {
    const id = attr(c, 'id') ?? '';
    const paras = elements(c).filter((e) => local(e.name) === 'p');
    const last = paras[paras.length - 1];
    d.comments.set(id, {
      id, author: attr(c, 'author') ?? '', initials: attr(c, 'initials') ?? '', date: attr(c, 'date') ?? '',
      text: paras.map((p) => textOf(p, new Set(['t']))).join('\n'), xml: xml.slice(c.start, c.end),
      paraId: last ? attr(last, 'paraId') ?? null : null, parent: null, done: false,
      anchored: new RegExp(`commentReference\\b[^>]*\\bid="${id}"`).test(bodyXml), dirty: false,
    });
  }
  const ext = rels.find((r) => r.type.toLowerCase().endsWith('/commentsextended') && !r.external);
  if (!ext) return;
  d.extendedPath = resolveTarget(d.mainPath, ext.target);
  const extXml = await d.pkg.text(d.extendedPath);
  if (!extXml) return;
  const extRoot = parseXml(extXml);
  d.extendedHead = extRoot.innerStart === extRoot.end ? `${extXml.slice(0, extRoot.end - 2)}>` : extXml.slice(0, extRoot.innerStart);
  d.extendedTail = extRoot.innerStart === extRoot.end ? `</${extRoot.name}>` : extXml.slice(extRoot.innerEnd);
  const byPara = new Map([...d.comments.values()].filter((c) => c.paraId).map((c) => [c.paraId!, c]));
  for (const e of [...walk(extRoot)].filter((x: XEl) => local(x.name) === 'commentEx')) {
    const c = byPara.get(attr(e, 'paraId') ?? '');
    if (!c) continue;
    c.done = attr(e, 'done') === '1';
    const parent = byPara.get(attr(e, 'paraIdParent') ?? '');
    c.parent = parent ? parent.id : null;
  }
}

const newParaId = () => {
  const n = 1 + Math.floor(Math.random() * 0x7ffffffe);
  return n.toString(16).toUpperCase().padStart(8, '0');
};

export const initialsOf = (name: string) => name.split(/\s+/).filter(Boolean).map((x) => x[0]).join('').slice(0, 3).toUpperCase();

export function createComment(d: WordDoc, author: string, text: string, parent: string | null): Comment {
  let id = 0;
  while (d.comments.has(String(id))) id++;
  const c: Comment = {
    id: String(id), author, initials: initialsOf(author), date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), text, xml: null,
    paraId: newParaId(), parent, done: false, anchored: true, dirty: true,
  };
  d.comments.set(c.id, c);
  d.commentRev++;
  return c;
}

export function editComment(d: WordDoc, c: Comment, text: string) {
  c.text = text;
  c.xml = null;
  c.dirty = true;
  d.commentRev++;
}
export function setDone(d: WordDoc, c: Comment, done: boolean) {
  c.done = done;
  c.dirty = true;
  d.commentRev++;
}

/** The three anchors Word expects: start and end of the range, and the reference run. */
export function anchorNodes(d: WordDoc, id: string) {
  const w = d.w;
  const make = (kind: string, xml: string) => schema.nodes.opaque_inline.create({ kind, xml });
  return {
    start: make('range', `<${w}commentRangeStart ${w}id="${id}"/>`),
    end: make('range', `<${w}commentRangeEnd ${w}id="${id}"/>`),
    ref: make('comment', `<${w}r><${w}commentReference ${w}id="${id}"/></${w}r>`),
  };
}

export function anchoredIds(doc: PMNode) {
  const ids = new Set<string>();
  doc.descendants((n) => {
    if (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'comment') {
      const m = /\bid="([^"]+)"/.exec(n.attrs.xml);
      if (m) ids.add(m[1]);
    }
  });
  return ids;
}

function commentXml(d: WordDoc, c: Comment, w: string) {
  if (!c.paraId) c.paraId = newParaId();
  if (c.xml) {
    if (/\w+:paraId="/.test(c.xml)) return c.xml;
    const at = Math.max(c.xml.lastIndexOf(`<${w}p>`), c.xml.lastIndexOf(`<${w}p `));
    if (at < 0) return c.xml;
    const nameEnd = at + w.length + 2;
    return `${c.xml.slice(0, nameEnd)} w14:paraId="${c.paraId}"${c.xml.slice(nameEnd)}`;
  }
  const lines = c.text.split('\n');
  const paras = lines.map((line, i) => {
    const id = i === lines.length - 1 ? ` w14:paraId="${c.paraId}" w14:textId="77777777"` : '';
    const ref = i === 0 ? `<${w}r><${w}annotationRef/></${w}r>` : '';
    return `<${w}p${id}>${ref}<${w}r><${w}t xml:space="preserve">${escText(line)}</${w}t></${w}r></${w}p>`;
  }).join('');
  return `<${w}comment ${w}id="${c.id}" ${w}author="${escAttr(c.author)}" ${w}date="${c.date}" ${w}initials="${escAttr(c.initials)}">${paras}</${w}comment>`;
}

/** paraId attributes need w14 declared, and declared ignorable for older readers. */
function withNamespaces(head: string, w: string) {
  const tag = /<[\w.-]*:?comments\b[^>]*>/.exec(head);
  if (!tag || /xmlns:w14=/.test(tag[0])) return head;
  let open = tag[0].replace(/\/?>$/, ` xmlns:w14="${W14}">`);
  if (!/xmlns:mc=/.test(open)) open = open.replace(/>$/, ` xmlns:mc="${MC}">`);
  open = /mc:Ignorable="/.test(open) ? open.replace(/mc:Ignorable="/, 'mc:Ignorable="w14 ') : open.replace(/>$/, ' mc:Ignorable="w14">');
  void w;
  return head.replace(tag[0], open);
}

async function pruneIds(d: WordDoc, removed: Comment[]) {
  if (!removed.length) return;
  const rels = await d.pkg.rels(d.mainPath);
  const part = async (suffix: string) => {
    const r = rels.find((x) => x.type.toLowerCase().endsWith(suffix));
    return r ? resolveTarget(d.mainPath, r.target) : null;
  };
  const idsPath = await part('/commentsids');
  const extensiblePath = await part('/commentsextensible');
  const durable: string[] = [];
  if (idsPath) {
    let xml = (await d.pkg.text(idsPath)) ?? '';
    for (const c of removed) {
      if (!c.paraId) continue;
      xml = xml.replace(new RegExp(`<[\\w.-]*:?commentId\\b[^>]*paraId="${c.paraId}"[^>]*/>`), (m) => {
        const id = /durableId="([^"]+)"/.exec(m)?.[1];
        if (id) durable.push(id);
        return '';
      });
    }
    d.pkg.set(idsPath, xml);
  }
  if (extensiblePath && durable.length) {
    let xml = (await d.pkg.text(extensiblePath)) ?? '';
    for (const id of durable) xml = xml.replace(new RegExp(`<[\\w.-]*:?commentExtensible\\b[^>]*durableId="${id}"[^>]*?(/>|>[\\s\\S]*?</[\\w.-]*:?commentExtensible>)`), '');
    d.pkg.set(extensiblePath, xml);
  }
}

export async function writeComments(d: WordDoc, body: PMNode) {
  const anchored = anchoredIds(body);
  const all = [...d.comments.values()];
  const removed = all.filter((c) => c.anchored && !anchored.has(c.id));
  const kept = all.filter((c) => !c.anchored || anchored.has(c.id));
  const signature = JSON.stringify(kept.map((c) => [c.id, c.dirty, c.done, c.text]));
  if (!d.commentsWritten) {
    if (!removed.length && !kept.some((c) => c.dirty)) { d.commentsWritten = signature; return; }
  } else if (d.commentsWritten === signature) return;
  d.commentsWritten = signature;

  const w = d.w;
  if (!d.commentsPath) {
    const dir = d.mainPath.slice(0, d.mainPath.lastIndexOf('/') + 1);
    d.commentsPath = `${dir}comments.xml`;
    const ns = w ? `xmlns:${w.slice(0, -1)}` : 'xmlns';
    d.commentsHead = `${decl}<${w}comments ${ns}="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14">`;
    d.commentsTail = `</${w}comments>`;
    await d.pkg.addRel(d.mainPath, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'comments.xml');
    await d.pkg.ensureContentOverride(`/${d.commentsPath}`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
  }
  d.commentsHead = withNamespaces(d.commentsHead, w);
  d.pkg.set(d.commentsPath, d.commentsHead + kept.map((c) => commentXml(d, c, w)).join('') + d.commentsTail);

  if (!d.extendedPath) {
    const dir = d.mainPath.slice(0, d.mainPath.lastIndexOf('/') + 1);
    d.extendedPath = `${dir}commentsExtended.xml`;
    d.extendedHead = `${decl}<w15:commentsEx xmlns:w15="${W15}" xmlns:mc="${MC}" mc:Ignorable="w15">`;
    d.extendedTail = '</w15:commentsEx>';
    await d.pkg.addRel(d.mainPath, EXTENDED_REL, 'commentsExtended.xml');
    await d.pkg.ensureContentOverride(`/${d.extendedPath}`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml');
  }
  const prefix = /<([\w.-]+):commentsEx\b/.exec(d.extendedHead)?.[1] ?? 'w15';
  const parentPara = (c: Comment) => (c.parent ? d.comments.get(c.parent)?.paraId ?? null : null);
  d.pkg.set(d.extendedPath, d.extendedHead + kept.filter((c) => c.paraId).map((c) => {
    const parent = parentPara(c);
    return `<${prefix}:commentEx ${prefix}:paraId="${c.paraId}"${parent ? ` ${prefix}:paraIdParent="${parent}"` : ''} ${prefix}:done="${c.done ? 1 : 0}"/>`;
  }).join('') + d.extendedTail);
  await pruneIds(d, removed);
}
