# Folio

A lightweight, bloat-free desktop word processor for Microsoft Word files.

Folio creates, opens, edits and saves `.docx` files without touching the parts of a document it doesn't need to change. It's a small native app built with Rust, Tauri and ProseMirror, and it has no accounts, subscriptions or cloud service.

## Features

- **Word files:** create, open, edit and save `.docx`, `.docm` and `.dotx`
- **Formatting:** paragraph styles, fonts, sizes, colors, highlights, alignment and indents
- **Content:** lists, tables, images, links, page breaks, and find and replace
- **Page layout:** page view, margins, editable headers and footers, page numbers, and a different first page
- **Review:** Track Changes, threaded comments, and accept or reject, saved in Word's own format
- **Output:** print, or export to PDF

## Getting started

### Download

Download the latest build from the [Releases](../../releases) page, if one is available for your platform.

Builds aren't code-signed yet, so Windows SmartScreen and macOS Gatekeeper may warn you the first time you open one. On Windows, you can right-click the downloaded file, choose **Properties**, tick **Unblock**, and then run it.

On Windows, Folio needs the Microsoft Edge WebView2 Runtime, which is included with Windows 11.

### Build from source

Install [Node 22+](https://nodejs.org), [Rust](https://rustup.rs) and the [Tauri prerequisites](https://tauri.app/start/prerequisites/) for your platform. Then run:

```sh
npm install
npm run tauri dev      # run the app with live reload
npm run tauri build    # build an installer in src-tauri/target/release/bundle
```

The tests need Python 3 with `python-docx`:

```sh
pip install python-docx
npm test
```

Pushing a `v*` tag, or running the workflow by hand, builds Windows, macOS and Linux installers with GitHub Actions.

To cross-compile a single-file Windows `.exe` from Linux with MinGW, run `packaging/windows-portable/build.sh`. That build unpacks Microsoft's `WebView2Loader.dll` to `%LOCALAPPDATA%\com.kalystron.folio` on first launch, and WebView2 keeps its browser data in the same folder.

## Using Folio

### Headers and footers

Double-click the top or bottom margin of a page to edit in place. You can also use the toolbar button or the **Layout** menu. Press Esc to return to the document.

### Review

Select **Review** to open the Review bar and pane.

- **Track Changes** records insertions, deletions, paragraph breaks and whole table rows. The setting is saved with the document.
- **Show** switches between All Markup, No Markup and Original. Original is read-only.
- **Comments** support replies, editing, resolving and deleting.
- **Accept** and **Reject** work on one change, a selection, or every change, and you can undo them.
- **Back to editing** closes Review and hides markup. Your tracked changes and comments stay in the document.

Printing and PDF export follow the current view.

### PDF export

If LibreOffice is installed, Folio uses it for the closest match to Word's layout. Otherwise it uses a built-in exporter with embedded fonts, which handles text, styles, lists, images, tables, headers, footers and page numbers.

### Keyboard shortcuts

Use Cmd instead of Ctrl on macOS.

| Action | Shortcut |
|---|---|
| New, open, save | Ctrl+N, Ctrl+O, Ctrl+S |
| Save as | Ctrl+Shift+S |
| Export as PDF | Ctrl+Shift+E |
| Print | Ctrl+P |
| Bold, italic, underline | Ctrl+B, Ctrl+I, Ctrl+U |
| Undo, redo | Ctrl+Z, Ctrl+Y |
| Find and replace | Ctrl+F |
| Insert link | Ctrl+K |
| Page break | Ctrl+Enter |
| New comment | Ctrl+Alt+M |

## How files stay intact

Folio doesn't rebuild a document from its own model.

1. `src/docx/xml.ts` parses XML and keeps each element's position in the original file.
2. `src/docx/document.ts` maps supported content to ProseMirror nodes. Each paragraph, table, row and cell remembers its exact source text.
3. `src/docx/export.ts` writes unchanged blocks back byte for byte. It regenerates only edited blocks and carries unknown formatting along.
4. Other package parts stay as they arrived unless an edit needs them.

Content the editor can't represent stays in the document as locked objects, including fields, equations, shapes, content controls and irregular tables. The status bar shows how many are in the open file.

Tracked changes and comments use standard `w:ins`, `w:del`, `comments.xml` and `commentsExtended.xml`.

## Compatibility

Output is checked with OOXML schema validation, python-docx, pandoc and LibreOffice. Testing in Microsoft Word is still in progress, so keep a copy of important documents and please [report anything that looks wrong](../../issues).

## Known limitations

**Editing and layout**

- On-screen pagination is approximate. Paragraphs move to the next page whole instead of splitting.
- Footnotes are displayed and preserved, but they can't be edited.
- Custom tab stops show at default positions on screen and in the built-in PDF. They're kept in the file.
- Each document uses one page setup on screen and in the built-in PDF. Mixed sections are preserved in the file.
- Margin changes apply to the whole document and replace any per-section margins.
- If sections have different headers, Folio shows and edits one set. The others stay in the file.
- Table style banding isn't drawn on screen, and the built-in PDF skips paragraph borders, table styles and floating objects.
- Password-protected documents and legacy `.doc` files can't be opened.

**Review**

- Formatting changes, table column changes, and edits in headers and footers aren't tracked. Formatting changes tracked in Word can still be accepted or rejected.
- Deleting a locked block with tracking on removes it outright.
- Tracked moves from Word stay as locked items.
- A deleted paragraph break still shows as a line break in No Markup until it's accepted.

## Security

Folio treats every document as untrusted input.

- The app runs under a strict content security policy and doesn't load remote content.
- XML document type declarations are rejected, and zip sizes are capped.
- Links open only for `http`, `https` and `mailto`.
- The Rust side reads and writes only files you chose in a dialog or opened with the app.
- Saves go through a temporary file and a rename, so a failed write can't damage the original.

## Fonts

Carlito, Caladea and Liberation are bundled as metric-compatible stand-ins for Calibri, Cambria, Arial, Times New Roman and Courier New. Their licenses are in `public/fonts/licenses`.

## Author

Developed by Daniel Wood. Find him on [LinkedIn](https://www.linkedin.com/in/danielewood) and [GitHub](https://github.com/automatesecurity).

## License

See [LICENSE](LICENSE).
