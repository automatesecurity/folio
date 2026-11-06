import { Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { authorColor, schema } from '../editor/schema';

export const commentIdOf = (xml: string) => /\bid="([^"]+)"/.exec(xml)?.[1] ?? '';

function build(doc: PMNode) {
  const decos: Decoration[] = [];
  const open = new Map<string, number>();
  doc.descendants((n, pos) => {
    if (n.type === schema.nodes.opaque_inline && n.attrs.kind === 'range' && /commentRange/.test(n.attrs.xml)) {
      const id = commentIdOf(n.attrs.xml);
      if (/commentRangeStart/.test(n.attrs.xml)) open.set(id, pos + 1);
      else if (open.has(id) && open.get(id)! < pos) decos.push(Decoration.inline(open.get(id)!, pos, { class: 'c-hl', 'data-cid': id }));
    }
    if (n.type === schema.nodes.table_row && n.attrs.rowChange) {
      decos.push(Decoration.node(pos, pos + n.nodeSize, { class: `row-${n.attrs.rowChange.kind}` }));
    }
    if (n.type !== schema.nodes.paragraph) return true;
    const pm = n.attrs.pmark;
    if (!pm) return true;
    const kind = pm.del ? 'del' : 'ins';
    const rev = pm.del ?? pm.ins;
    decos.push(Decoration.widget(pos + n.nodeSize - 1, () => {
      const s = document.createElement('span');
      s.className = `p-mark t-${kind} t-a${authorColor(rev.author)}`;
      s.textContent = '¶';
      return s;
    }, { side: 1, key: `pm-${kind}-${rev.author}` }));
    let whole = n.content.size > 0 || !!pm;
    n.forEach((c) => {
      const hidden = c.type === schema.nodes.opaque_inline && ['bookmark', 'range', 'comment'].includes(c.attrs.kind);
      if (!hidden && !(pm.del ? schema.marks.tracked_del : schema.marks.tracked_ins).isInSet(c.marks)) whole = false;
    });
    if (whole) decos.push(Decoration.node(pos, pos + n.nodeSize, { class: pm.del ? 'p-all-del' : 'p-all-ins' }));
    return true;
  });
  return DecorationSet.create(doc, decos);
}

export function reviewPlugin() {
  return new Plugin<DecorationSet>({
    key: new PluginKey('review'),
    state: { init: (_, s) => build(s.doc), apply: (tr, old) => (tr.docChanged ? build(tr.doc) : old) },
    props: { decorations(s) { return this.getState(s); } },
  });
}
