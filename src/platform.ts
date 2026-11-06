// The one place that knows whether we are inside the desktop shell.
import { invoke } from '@tauri-apps/api/core';
import { getCurrentWindow } from '@tauri-apps/api/window';
import { listen } from '@tauri-apps/api/event';

export const isDesktop = '__TAURI_INTERNALS__' in window;

export interface OpenedFile { path: string | null; name: string; bytes: Uint8Array }
const baseName = (p: string) => p.split(/[\\/]/).pop() ?? p;

function pickInBrowser(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.onchange = () => resolve(input.files?.[0] ?? null);
    input.oncancel = () => resolve(null);
    input.click();
  });
}
function download(bytes: Uint8Array, name: string, type: string) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

async function readPath(path: string): Promise<OpenedFile> {
  const buf = await invoke<ArrayBuffer>('read_document', { path });
  return { path, name: baseName(path), bytes: new Uint8Array(buf) };
}

export async function pickDocument(): Promise<OpenedFile | null> {
  if (!isDesktop) {
    const f = await pickInBrowser('.docx,.docm,.dotx');
    return f ? { path: null, name: f.name, bytes: new Uint8Array(await f.arrayBuffer()) } : null;
  }
  const path = await invoke<string | null>('pick_open_path');
  return path ? readPath(path) : null;
}
export async function launchDocument(): Promise<OpenedFile | null> {
  if (!isDesktop) return null;
  const path = await invoke<string | null>('launch_path');
  return path ? readPath(path) : null;
}
export async function pickImage(): Promise<{ name: string; bytes: Uint8Array } | null> {
  if (!isDesktop) {
    const f = await pickInBrowser('image/png,image/jpeg,image/gif,image/webp');
    return f ? { name: f.name, bytes: new Uint8Array(await f.arrayBuffer()) } : null;
  }
  const path = await invoke<string | null>('pick_image_path');
  return path ? readPath(path) : null;
}
/** Returns the chosen path, or null if cancelled. In a browser the name is the destination. */
export async function pickSavePath(defaultName: string, kind: 'docx' | 'pdf'): Promise<string | null> {
  if (!isDesktop) return defaultName;
  return invoke<string | null>('pick_save_path', { defaultName, kind });
}
export async function writeFile(path: string, bytes: Uint8Array, kind: 'docx' | 'pdf') {
  if (!isDesktop) {
    download(bytes, path, kind === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    return;
  }
  await invoke('write_document', bytes, { headers: { 'x-path': encodeURIComponent(path) } });
}
export async function libreOfficePdf(docx: Uint8Array): Promise<Uint8Array | null> {
  if (!isDesktop) return null;
  const available = await invoke<boolean>('libreoffice_available');
  if (!available) return null;
  const buf = await invoke<ArrayBuffer>('convert_with_libreoffice', docx);
  return new Uint8Array(buf);
}
export async function printPage() {
  // window.print() does nothing in macOS webviews, so the desktop app goes through the shell.
  if (isDesktop) await invoke('print_window');
  else window.print();
}
export function openExternal(url: string) {
  if (isDesktop) void invoke('open_link', { url });
  else window.open(url, '_blank', 'noopener');
}
export function setTitle(title: string) {
  document.title = title;
  if (isDesktop) void getCurrentWindow().setTitle(title);
}
export function onMenu(fn: (id: string) => void) {
  if (isDesktop) void listen<string>('menu', (e) => fn(e.payload));
}
export function onCloseRequest(fn: () => Promise<boolean>) {
  if (!isDesktop) {
    window.addEventListener('beforeunload', (e) => { void fn; if ((window as unknown as { __dirty?: boolean }).__dirty) e.preventDefault(); });
    return;
  }
  const win = getCurrentWindow();
  void win.onCloseRequested(async (e) => {
    e.preventDefault();
    if (await fn()) await win.destroy();
  });
}
