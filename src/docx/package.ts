import JSZip from 'jszip';
import { attr, child, elements, escAttr, local, parseXml, XEl } from './xml';

const MAX_PART = 512 * 1024 * 1024;
const MAX_TOTAL = 2 * 1024 * 1024 * 1024;

export interface Rel { id: string; type: string; target: string; external: boolean }

export class DocxError extends Error {}

/** The OPC package. Parts nobody asked to change stay as the bytes they arrived as. */
export class Package {
  private constructor(public zip: JSZip) {}

  static empty() { return new Package(new JSZip()); }

  static async load(bytes: Uint8Array): Promise<Package> {
    if (bytes[0] === 0xd0 && bytes[1] === 0xcf)
      throw new DocxError('This file is password-protected or uses the older .doc format. Folio opens .docx, .docm and .dotx files. Remove the password or convert it to .docx in Word first.');
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new DocxError('This file is not a Word document. It has no .docx package structure.');
    let zip: JSZip;
    try { zip = await JSZip.loadAsync(bytes, { createFolders: false }); }
    catch { throw new DocxError('This file is damaged. Its .docx package could not be read.'); }
    let total = 0;
    zip.forEach((path, f) => {
      const size = (f as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize ?? 0;
      total += size;
      if (size > MAX_PART || total > MAX_TOTAL) throw new DocxError('This file expands to more data than Folio will open (2 GB).');
    });
    if (!zip.file('[Content_Types].xml')) throw new DocxError('This file is not a Word document. It is a zip archive without Word content.');
    return new Package(zip);
  }

  has(path: string) { return !!this.zip.file(path); }
  async text(path: string): Promise<string | null> {
    const f = this.zip.file(path);
    if (!f) return null;
    const bytes = await f.async('uint8array');
    return decodeXml(bytes);
  }
  bytes(path: string) { return this.zip.file(path)?.async('uint8array') ?? Promise.resolve(null); }
  set(path: string, content: string | Uint8Array) { this.zip.file(path, content, { createFolders: false }); }
  list() { return Object.keys(this.zip.files).filter((p) => !this.zip.files[p].dir); }

  async rels(ownerPath: string): Promise<Rel[]> {
    const xml = await this.text(relsPath(ownerPath));
    if (!xml) return [];
    return elements(parseXml(xml)).filter((e) => local(e.name) === 'Relationship').map((e) => ({
      id: e.attrs.Id, type: e.attrs.Type, target: e.attrs.Target, external: e.attrs.TargetMode === 'External',
    }));
  }

  async addRel(ownerPath: string, type: string, target: string, external = false): Promise<string> {
    const path = relsPath(ownerPath);
    let xml = await this.text(path);
    if (!xml) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    let max = 0;
    for (const m of xml.matchAll(/\bId="rId(\d+)"/g)) max = Math.max(max, +m[1]);
    const id = `rId${max + 1}`;
    const el = `<Relationship Id="${id}" Type="${escAttr(type)}" Target="${escAttr(target)}"${external ? ' TargetMode="External"' : ''}/>`;
    if (xml.includes('</Relationships>')) xml = xml.replace('</Relationships>', el + '</Relationships>');
    else xml = xml.replace(/<Relationships([^>]*)\/>/, `<Relationships$1>${el}</Relationships>`);
    this.set(path, xml);
    return id;
  }

  async ensureContentDefault(ext: string, type: string) {
    let xml = (await this.text('[Content_Types].xml'))!;
    if (new RegExp(`<Default[^>]*Extension="${ext}"`, 'i').test(xml)) return;
    xml = xml.replace('</Types>', `<Default Extension="${ext}" ContentType="${type}"/></Types>`);
    this.set('[Content_Types].xml', xml);
  }
  async ensureContentOverride(part: string, type: string) {
    let xml = (await this.text('[Content_Types].xml'))!;
    if (xml.includes(`PartName="${part}"`)) return;
    xml = xml.replace('</Types>', `<Override PartName="${part}" ContentType="${type}"/></Types>`);
    this.set('[Content_Types].xml', xml);
  }

  async mainPart(): Promise<string> {
    const rels = await this.rels('');
    const r = rels.find((x) => x.type.endsWith('/officeDocument'));
    const target = r ? resolveTarget('', r.target) : 'word/document.xml';
    if (!this.has(target)) throw new DocxError('This file has no main document part, so it cannot be opened.');
    return target;
  }

  generate(): Promise<Uint8Array> {
    return this.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }
}

export function relsPath(owner: string) {
  if (!owner) return '_rels/.rels';
  const i = owner.lastIndexOf('/');
  return `${owner.slice(0, i + 1)}_rels/${owner.slice(i + 1)}.rels`;
}

export function resolveTarget(owner: string, target: string) {
  if (target.startsWith('/')) return target.slice(1);
  const base = owner.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') base.pop();
    else if (seg !== '.') base.push(seg);
  }
  return base.join('/');
}

function decodeXml(bytes: Uint8Array) {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  // ignoreBOM keeps a leading BOM in the string, so re-encoding gives the same bytes.
  return new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes);
}

export { attr, child };
export type { XEl };
