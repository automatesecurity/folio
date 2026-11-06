import type { Numbering, StyleSheet } from '../docx/styles';

/** State the schema's render functions need about the open document. */
class Session {
  id = 0;
  styles: StyleSheet | null = null;
  numbering: Numbering | null = null;
  assets = new Map<string, string>();
  pasteVault = new Map<string, Record<string, unknown>>();
  private counter = 0;
  /** Set by the open document: allocates a list definition while parsing pasted HTML. */
  newList: (kind: 'bullet' | 'ordered') => string = () => '1';
  private lists = new WeakMap<object, string>();

  reset() {
    this.id++;
    for (const url of this.assets.values()) if (url.startsWith('blob:')) URL.revokeObjectURL(url);
    this.assets.clear();
    this.pasteVault.clear();
  }
  stash(attrs: Record<string, unknown>) {
    const key = String(++this.counter);
    this.pasteVault.set(key, { ...attrs });
    return key;
  }
  listFor(kind: 'bullet' | 'ordered', listEl: object | null) {
    if (!listEl) return this.newList(kind);
    let id = this.lists.get(listEl);
    if (!id) { id = this.newList(kind); this.lists.set(listEl, id); }
    return id;
  }
}
export const session = new Session();
