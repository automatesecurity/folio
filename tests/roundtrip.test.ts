import { describe, it, expect, beforeAll } from 'vitest';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import { EditorState, TextSelection } from 'prosemirror-state';
import { fillPageFields, WordDoc } from '../src/docx/document';
import { documentXml, exportDocx } from '../src/docx/export';
import { blankDocument } from '../src/docx/template';
import { schema } from '../src/editor/schema';

const dir = mkdtempSync(join(tmpdir(), 'folio-'));
const fixture = join(dir, 'complex.docx');

const py = String.raw`
import sys, io, struct, zlib
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls
def png():
    def chunk(t,d):
        c=struct.pack('>I',len(d))+t+d
        return c+struct.pack('>I',zlib.crc32(t+d)&0xffffffff)
    raw=b''.join(b'\0'+b'\xff\x00\x00'*8 for _ in range(8))
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',8,8,8,2,0,0,0))+chunk(b'IDAT',zlib.compress(raw))+chunk(b'IEND',b'')
d=Document()
d.core_properties.title='Fixture'
s=d.sections[0]
s.header.paragraphs[0].text='Confidential header'
fp=s.footer.paragraphs[0]; fp.text='Page '
r=fp.add_run()
for t,x in (('begin',None),(None,' PAGE '),('end',None)):
    if t:
        e=OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'),t); r._r.append(e)
    else:
        e=OxmlElement('w:instrText'); e.text=x; r._r.append(e)
d.add_heading('Quarterly report',0)
d.add_heading('Overview',1)
p=d.add_paragraph('Plain ')
p.add_run('bold').bold=True
p.add_run(' and ')
i=p.add_run('italic red'); i.italic=True; i.font.color.rgb=RGBColor(255,0,0)
p.add_run(' sized').font.size=Pt(18)
p.add_run(' tail text.')
d.add_paragraph('First bullet',style='List Bullet')
d.add_paragraph('Second bullet',style='List Bullet')
d.add_paragraph('Step one',style='List Number')
p=d.add_paragraph('Image here: ')
p.add_run().add_picture(io.BytesIO(png()),width=Inches(1))
t=d.add_table(rows=3,cols=3); t.style='Table Grid'
for r_ in range(3):
    for c in range(3): t.cell(r_,c).text=f'R{r_}C{c}'
t.cell(0,0).merge(t.cell(0,1))
t.cell(1,2).merge(t.cell(2,2))
irregular=d.add_table(rows=2,cols=2)
tr=irregular.rows[1]._tr
trPr=tr.get_or_add_trPr(); g=OxmlElement('w:gridBefore'); g.set(qn('w:val'),'1'); trPr.append(g)
tr.remove(tr.tc_lst[0])
irregular.cell(0,0).text='irregular'
p=d.add_paragraph('Tracked: ')
p._p.append(parse_xml('<w:ins %s w:id="901" w:author="Dan" w:date="2026-01-01T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins>'%nsdecls('w')))
p._p.append(parse_xml('<w:del %s w:id="902" w:author="Dan" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>deleted</w:delText></w:r></w:del>'%nsdecls('w')))
p=d.add_paragraph('Bookmarked ')
p._p.append(parse_xml('<w:bookmarkStart %s w:id="77" w:name="mark"/>'%nsdecls('w')))
p.add_run('words')
p._p.append(parse_xml('<w:bookmarkEnd %s w:id="77"/>'%nsdecls('w')))
sd=parse_xml('<w:sdt %s><w:sdtPr><w:alias w:val="TOC"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Control content</w:t></w:r></w:p></w:sdtContent></w:sdt>'%nsdecls('w'))
d.element.body.insert(len(d.element.body)-1,sd)
d.add_paragraph('The final paragraph.')
d.save(sys.argv[1])
`;

async function parts(bytes: Uint8Array) {
  const z = await JSZip.loadAsync(bytes);
  const out: Record<string, Buffer> = {};
  for (const n of Object.keys(z.files)) if (!z.files[n].dir) out[n] = Buffer.from(await z.files[n].async('uint8array'));
  return out;
}
const text = (path: string) => execFileSync('python3', ['-c', `
import sys
from docx import Document
d=Document(sys.argv[1])
print('\\n'.join(p.text for p in d.paragraphs))
print('TABLES',len(d.tables))
`, path]).toString();

let original: Uint8Array;
beforeAll(() => {
  writeFileSync(join(dir, 'make.py'), py);
  execFileSync('python3', [join(dir, 'make.py'), fixture]);
  original = readFileSync(fixture);
});

function edit(doc: WordDoc, fn: (s: EditorState) => EditorState) {
  return fn(EditorState.create({ doc: doc.body, schema }));
}
const findText = (node: import('prosemirror-model').Node, needle: string) => {
  let at = -1;
  node.descendants((n, pos) => { if (at < 0 && n.isText && n.text!.includes(needle)) at = pos + n.text!.indexOf(needle); });
  return at;
};

describe('round trip', () => {
  it('reproduces document.xml byte for byte when nothing changed', async () => {
    const d = await WordDoc.open(original);
    const before = await parts(original);
    expect(documentXml(d, d.body)).toBe(before['word/document.xml'].toString('utf8'));
  });

  it('keeps every package part identical on an untouched save', async () => {
    const d = await WordDoc.open(original);
    const after = await parts(await exportDocx(d, d.body));
    const before = await parts(original);
    expect(Object.keys(after).sort()).toEqual(Object.keys(before).sort());
    for (const k of Object.keys(before)) expect(after[k].equals(before[k]), k).toBe(true);
  });

  it('changes only the edited paragraph', async () => {
    const d = await WordDoc.open(original);
    const s = edit(d, (st) => {
      const at = findText(st.doc, 'tail text');
      return st.apply(st.tr.insertText('EDITED ', at));
    });
    const out = await exportDocx(d, s.doc);
    const a = (await parts(original))['word/document.xml'].toString('utf8');
    const b = (await parts(out))['word/document.xml'].toString('utf8');
    const pa = a.match(/<w:p[ >][\s\S]*?<\/w:p>/g)!, pb = b.match(/<w:p[ >][\s\S]*?<\/w:p>/g)!;
    expect(pa.length).toBe(pb.length);
    expect(pa.filter((p, i) => p !== pb[i]).length).toBe(1);
    const t = text(writeTmp(out));
    expect(t).toContain('EDITED tail text.');
    expect(b).toContain('<w:b/>');
    expect(b).toContain('w:val="FF0000"');
    expect(b).toContain('w:ins ');
    expect(b).toContain('w:del ');
    expect(b).toContain('Control content');
    expect(b).toContain('gridBefore');
    const prior = await parts(original);
    const now = await parts(out);
    for (const k of Object.keys(prior)) if (k !== 'word/document.xml') expect(now[k].equals(prior[k]), k).toBe(true);
  });

  it('survives inserting a table row and regenerates vertical merges', async () => {
    const d = await WordDoc.open(original);
    const s = edit(d, (st) => {
      let tablePos = -1;
      st.doc.forEach((n, off) => { if (n.type.name === 'table' && tablePos < 0) tablePos = off; });
      const table = st.doc.nodeAt(tablePos)!;
      const row = table.child(0);
      const cloned = schema.nodes.table_row.create(null, [
        schema.nodes.table_cell.createAndFill({ colspan: 2 })!,
        schema.nodes.table_cell.createAndFill()!,
      ]);
      return st.apply(st.tr.insert(tablePos + 1 + row.nodeSize, cloned));
    });
    const out = await exportDocx(d, s.doc);
    const xml = (await parts(out))['word/document.xml'].toString('utf8');
    expect(xml).toContain('w:vMerge');
    expect(text(writeTmp(out))).toContain('TABLES 2');
    const rows = execFileSync('python3', ['-c', `
import sys
from docx import Document
print(len(Document(sys.argv[1]).tables[0].rows))`, writeTmp(out)]).toString().trim();
    expect(rows).toBe('4');
  });

  it('removes the partner when half a bookmark is deleted', async () => {
    const d = await WordDoc.open(original);
    const s = edit(d, (st) => {
      let pos = -1;
      st.doc.descendants((n, p) => { if (n.type.name === 'opaque_inline' && n.attrs.xml.includes('bookmarkStart') && n.attrs.xml.includes('"mark"')) pos = p; });
      return st.apply(st.tr.delete(pos, pos + 1));
    });
    const xml = (await parts(await exportDocx(d, s.doc)))['word/document.xml'].toString('utf8');
    expect(xml).not.toContain('bookmarkEnd w:id="77"');
  });

  it('does not duplicate paragraph identity on split', async () => {
    const d = await WordDoc.open(original);
    const s = edit(d, (st) => st.apply(st.tr.split(findText(st.doc, 'tail text'))));
    const xml = (await parts(await exportDocx(d, s.doc)))['word/document.xml'].toString('utf8');
    const ids = [...xml.matchAll(/w14:paraId="([^"]+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('builds a blank document Word tooling can open, with list, link and image', async () => {
    const d = await blankDocument();
    const para = schema.nodes.paragraph;
    const numId = d.numbering.newList('ordered', d.w, true);
    const bullet = d.numbering.newList('bullet', d.w, true);
    const { rId } = await d.addImage(Buffer.from('89504e470d0a1a0a', 'hex'), 'png');
    const body = schema.nodes.doc.create(null, [
      para.create({ pStyle: 'Heading1' }, schema.text('Title & <things>')),
      para.create({ numId, ilvl: 0 }, schema.text('one')),
      para.create({ numId: bullet, ilvl: 0 }, schema.text('dot')),
      para.create(null, [schema.text('link', [schema.marks.link.create({ href: 'https://example.com/?a=1&b=2' })]), schema.text('\ttabbed')]),
      para.create(null, schema.nodes.opaque_inline.create({ kind: 'drawing', xml: d.drawingXml(rId, 100, 100, 1, 'pic') })),
    ]);
    const out = await exportDocx(d, body);
    const p = writeTmp(out);
    const t = text(p);
    expect(t).toContain('Title & <things>');
    const reopened = await WordDoc.open(out);
    expect(reopened.numbering.kindOf(numId)).toBe('ordered');
    expect(reopened.numbering.kindOf(bullet)).toBe('bullet');
    expect((await parts(out))['word/_rels/document.xml.rels'].toString()).toContain('example.com/?a=1&amp;b=2');
  });

  it('reads headers and live page fields', async () => {
    const d = await WordDoc.open(original);
    expect(d.findPart('header', 'default')!.doc.textContent).toBe('Confidential header');
    const fields: string[] = [];
    d.findPart('footer', 'default')!.doc.descendants((n) => { if (n.attrs.instr) fields.push(n.attrs.instr); });
    expect(fields).toEqual(['PAGE']);
  });

  it('edits a header and leaves every other part alone', async () => {
    const d = await WordDoc.open(original);
    const part = d.findPart('header', 'default')!;
    const at = findText(part.doc, 'Confidential');
    part.doc = EditorState.create({ doc: part.doc, schema }).apply(EditorState.create({ doc: part.doc, schema }).tr.insertText('Very ', at)).doc;
    const out = await exportDocx(d, d.body);
    const before = await parts(original), after = await parts(out);
    const changed = Object.keys(before).filter((k) => !before[k].equals(after[k]));
    expect(changed).toEqual([part.path]);
    expect(pyHeader(writeTmp(out))).toContain('Very Confidential header');
  });

  it('adds new header, footer and first-page parts to a blank document', async () => {
    const d = await blankDocument();
    const para = schema.nodes.paragraph;
    const set = (kind: 'header' | 'footer', slot: 'default' | 'first', content: import('prosemirror-model').Node[]) => {
      const part = d.hfPart(kind, slot);
      part.doc = schema.nodes.doc.create(null, para.create({ pStyle: kind === 'header' ? 'Header' : 'Footer', jc: 'right' }, content));
    };
    set('header', 'default', [schema.text('Kalystron report')]);
    set('header', 'first', [schema.text('Cover header')]);
    set('footer', 'default', [schema.text('Page '), schema.nodes.opaque_inline.create({ kind: 'field', xml: '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>', label: '1', instr: 'PAGE' })]);
    d.hfPart('footer', 'first'); // opened but left empty
    const body = schema.nodes.doc.create({ ...d.body.attrs, titlePg: true }, d.body.content);
    const out = await exportDocx(d, body);
    const p = await parts(out);
    expect(Object.keys(p).filter((k) => /word\/(header|footer)\d/.test(k)).sort()).toEqual(['word/footer1.xml', 'word/header1.xml', 'word/header2.xml']);
    expect(p['[Content_Types].xml'].toString()).toContain('/word/header2.xml');
    const xml = p['word/document.xml'].toString();
    expect(xml).toMatch(/<w:sectPr><w:\w+Reference[\s\S]*<w:pgMar[^>]*\/><w:titlePg\/><\/w:sectPr>/);
    const r = pyHeader(writeTmp(out));
    expect(r).toContain('Kalystron report');
    expect(r).toContain('Cover header');
    // A second save must not stack references or lose them.
    const again = (await parts(await exportDocx(d, body)))['word/document.xml'].toString();
    expect(again.match(/headerReference/g)!.length).toBe(2);
    const reopened = await WordDoc.open(await exportDocx(d, body));
    expect(reopened.body.attrs.titlePg).toBe(true);
    expect(reopened.findPart('header', 'first')!.doc.textContent).toBe('Cover header');
  });

  it('applies margins to every section and removes the first-page flag', async () => {
    const z = await JSZip.loadAsync(original);
    let xml = await z.file('word/document.xml')!.async('string');
    const sect = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(xml)![0].replace(/<w:\w+Reference[^>]*\/>/g, '');
    xml = xml.replace('<w:p><w:r><w:t>The final paragraph.', `<w:p><w:pPr>${sect}</w:pPr></w:p><w:p><w:r><w:t>The final paragraph.`);
    expect(xml).toContain('<w:pPr><w:sectPr');
    z.file('word/document.xml', xml.replace('</w:sectPr></w:body>', '<w:titlePg/></w:sectPr></w:body>'));
    const two = await z.generateAsync({ type: 'uint8array' });
    const d = await WordDoc.open(two);
    expect(d.body.attrs.titlePg).toBe(true);
    const m = { ...d.body.attrs.pgMar, left: 720, right: 2160, top: 1000 };
    const out = await exportDocx(d, schema.nodes.doc.create({ pgMar: m, titlePg: false }, d.body.content));
    const margins = execFileSync('python3', ['-c', `
import sys
from docx import Document
print([(s.left_margin.twips, s.right_margin.twips, s.top_margin.twips, s.different_first_page_header_footer) for s in Document(sys.argv[1]).sections])`, writeTmp(out)]).toString().trim();
    expect(margins).toBe('[(720, 2160, 1000, False), (720, 2160, 1000, False)]');
  });

  it('rejects files that are not documents and refuses DTDs', async () => {
    await expect(WordDoc.open(new TextEncoder().encode('hello'))).rejects.toThrow(/not a Word/);
    const z = await JSZip.loadAsync(original);
    z.file('word/document.xml', '<!DOCTYPE x [<!ENTITY a "b">]><w:document/>');
    await expect(WordDoc.open(await z.generateAsync({ type: 'uint8array' }))).rejects.toThrow(/damaged/);
  });
});

const pyHeader = (path: string) => execFileSync('python3', ['-c', `
import sys
from docx import Document
s=Document(sys.argv[1]).sections[0]
print(s.header.paragraphs[0].text, '|', s.first_page_header.paragraphs[0].text)`, path]).toString();

function writeTmp(b: Uint8Array) {
  const p = join(dir, `o${Math.random().toString(36).slice(2)}.docx`);
  writeFileSync(p, b);
  return p;
}
void TextSelection;
