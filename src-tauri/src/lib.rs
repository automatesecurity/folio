use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::Mutex;
use std::time::{Duration, Instant};

use percent_encoding::percent_decode_str;
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::menu::{MenuBuilder, MenuItem, PredefinedMenuItem, SubmenuBuilder};
use tauri::{AppHandle, Emitter, Manager, State};
use tauri_plugin_dialog::DialogExt;
use tauri_plugin_opener::OpenerExt;

const MAX_FILE: u64 = 512 * 1024 * 1024;
const LIBREOFFICE_TIMEOUT: Duration = Duration::from_secs(120);

/// Files the user has actually chosen. The webview can only touch these, so a
/// script bug or a hostile document can never read or overwrite anything else.
#[derive(Default)]
struct Granted(Mutex<HashSet<PathBuf>>);

impl Granted {
    fn add(&self, p: PathBuf) -> String {
        let shown = p.to_string_lossy().into_owned();
        self.0.lock().unwrap().insert(p);
        shown
    }
    fn check(&self, path: &str) -> Result<PathBuf, String> {
        let p = PathBuf::from(path);
        if self.0.lock().unwrap().contains(&p) {
            Ok(p)
        } else {
            Err("That file was not chosen in a file dialog.".into())
        }
    }
}

struct LaunchFile(Mutex<Option<PathBuf>>);

// The dialog commands are async so they run off the main thread. A blocking
// dialog on the main thread would deadlock the event loop that has to draw it.
#[tauri::command]
async fn pick_open_path(app: AppHandle) -> Option<String> {
    let f = app
        .dialog()
        .file()
        .add_filter("Word documents", &["docx", "docm", "dotx"])
        .blocking_pick_file()?;
    Some(app.state::<Granted>().add(f.into_path().ok()?))
}

#[tauri::command]
async fn pick_image_path(app: AppHandle) -> Option<String> {
    let f = app
        .dialog()
        .file()
        .add_filter("Images", &["png", "jpg", "jpeg", "gif", "bmp", "webp"])
        .blocking_pick_file()?;
    Some(app.state::<Granted>().add(f.into_path().ok()?))
}

#[tauri::command]
async fn pick_save_path(app: AppHandle, default_name: String, kind: String) -> Option<String> {
    let (label, exts): (&str, &[&str]) = if kind == "pdf" {
        ("PDF", &["pdf"])
    } else {
        ("Word documents", &["docx", "docm", "dotx"])
    };
    let f = app
        .dialog()
        .file()
        .set_file_name(&default_name)
        .add_filter(label, exts)
        .blocking_save_file()?;
    let mut p = f.into_path().ok()?;
    if p.extension().is_none() {
        p.set_extension(exts[0]);
    }
    Some(app.state::<Granted>().add(p))
}

#[tauri::command]
fn launch_path(launch: State<LaunchFile>, granted: State<Granted>) -> Option<String> {
    let p = launch.0.lock().unwrap().take()?;
    Some(granted.add(p))
}

#[tauri::command]
fn read_document(path: String, granted: State<Granted>) -> Result<Response, String> {
    let p = granted.check(&path)?;
    let size = fs::metadata(&p).map_err(|e| e.to_string())?.len();
    if size > MAX_FILE {
        return Err("This file is larger than the 512 MB Folio will open.".into());
    }
    Ok(Response::new(fs::read(&p).map_err(|e| e.to_string())?))
}

/// Writes beside the target and renames over it, so a crash or full disk can't
/// leave a half-written document where the good one used to be.
fn atomic_write(path: &Path, bytes: &[u8]) -> Result<(), String> {
    use std::io::Write;
    let dir = path.parent().filter(|d| !d.as_os_str().is_empty()).unwrap_or(Path::new("."));
    let mut tmp = tempfile::Builder::new()
        .prefix(".folio-")
        .tempfile_in(dir)
        .map_err(|e| e.to_string())?;
    tmp.write_all(bytes).map_err(|e| e.to_string())?;
    tmp.as_file().sync_all().map_err(|e| e.to_string())?;
    if let Ok(meta) = fs::metadata(path) {
        let _ = fs::set_permissions(tmp.path(), meta.permissions());
    }
    tmp.persist(path).map_err(|e| e.error.to_string())?;
    Ok(())
}

#[tauri::command]
fn write_document(request: Request<'_>, granted: State<Granted>) -> Result<(), String> {
    let header = request
        .headers()
        .get("x-path")
        .and_then(|v| v.to_str().ok())
        .ok_or("Missing path.")?;
    let path = percent_decode_str(header).decode_utf8().map_err(|e| e.to_string())?;
    let p = granted.check(&path)?;
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("Expected file bytes.".into());
    };
    atomic_write(&p, bytes)
}

fn find_libreoffice() -> Option<PathBuf> {
    for name in ["soffice", "libreoffice"] {
        if let Ok(p) = which::which(name) {
            return Some(p);
        }
    }
    let known: &[&str] = if cfg!(target_os = "macos") {
        &["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    } else if cfg!(target_os = "windows") {
        &[
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    } else {
        &["/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice", "/snap/bin/libreoffice"]
    };
    known.iter().map(PathBuf::from).find(|p| p.exists())
}

#[tauri::command]
fn libreoffice_available() -> bool {
    find_libreoffice().is_some()
}

#[tauri::command]
async fn convert_with_libreoffice(request: Request<'_>) -> Result<Response, String> {
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("Expected file bytes.".into());
    };
    let bytes = bytes.clone();
    tauri::async_runtime::spawn_blocking(move || run_libreoffice(&bytes))
        .await
        .map_err(|e| e.to_string())?
        .map(Response::new)
}

fn run_libreoffice(docx: &[u8]) -> Result<Vec<u8>, String> {
    let exe = find_libreoffice().ok_or("LibreOffice is not installed.")?;
    let work = tempfile::tempdir().map_err(|e| e.to_string())?;
    let input = work.path().join("document.docx");
    fs::write(&input, docx).map_err(|e| e.to_string())?;
    let profile = work.path().join("profile");
    // Its own profile, so this works while the user has LibreOffice open.
    let profile_url = format!("file:///{}", profile.to_string_lossy().replace('\\', "/").trim_start_matches('/'));
    let mut child = Command::new(exe)
        .arg(format!("-env:UserInstallation={profile_url}"))
        .args(["--headless", "--norestore", "--nolockcheck", "--convert-to", "pdf", "--outdir"])
        .arg(work.path())
        .arg(&input)
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null())
        .spawn()
        .map_err(|e| e.to_string())?;
    let started = Instant::now();
    loop {
        match child.try_wait().map_err(|e| e.to_string())? {
            Some(_) => break,
            None if started.elapsed() > LIBREOFFICE_TIMEOUT => {
                let _ = child.kill();
                return Err("LibreOffice took too long to convert the document.".into());
            }
            None => std::thread::sleep(Duration::from_millis(100)),
        }
    }
    fs::read(work.path().join("document.pdf")).map_err(|_| "LibreOffice did not produce a PDF.".to_string())
}

#[tauri::command]
fn open_link(app: AppHandle, url: String) -> Result<(), String> {
    let lower = url.to_ascii_lowercase();
    if !(lower.starts_with("http://") || lower.starts_with("https://") || lower.starts_with("mailto:")) {
        return Err("Only web and email links can be opened.".into());
    }
    app.opener().open_url(url, None::<&str>).map_err(|e| e.to_string())
}

#[tauri::command]
fn print_window(window: tauri::WebviewWindow) -> Result<(), String> {
    window.print().map_err(|e| e.to_string())
}

fn build_menu(app: &tauri::App) -> tauri::Result<()> {
    let h = app.handle();
    let item = |id: &str, label: &str, accel: Option<&str>| MenuItem::with_id(h, id, label, true, accel);
    let file = SubmenuBuilder::new(h, "File")
        .item(&item("new", "New", Some("CmdOrCtrl+N"))?)
        .item(&item("open", "Open...", Some("CmdOrCtrl+O"))?)
        .separator()
        .item(&item("save", "Save", Some("CmdOrCtrl+S"))?)
        .item(&item("save_as", "Save As...", Some("CmdOrCtrl+Shift+S"))?)
        .item(&item("export_pdf", "Export as PDF...", Some("CmdOrCtrl+Shift+E"))?)
        .item(&item("print", "Print...", Some("CmdOrCtrl+P"))?)
        .separator()
        .item(&PredefinedMenuItem::close_window(h, Some("Close"))?)
        .build()?;
    let edit = SubmenuBuilder::new(h, "Edit")
        .item(&PredefinedMenuItem::cut(h, None)?)
        .item(&PredefinedMenuItem::copy(h, None)?)
        .item(&PredefinedMenuItem::paste(h, None)?)
        .item(&PredefinedMenuItem::select_all(h, None)?)
        .separator()
        .item(&item("find", "Find and Replace...", None)?)
        .build()?;
    let view = SubmenuBuilder::new(h, "View")
        .item(&item("zoom_in", "Zoom In", Some("CmdOrCtrl+="))?)
        .item(&item("zoom_out", "Zoom Out", Some("CmdOrCtrl+-"))?)
        .item(&item("zoom_reset", "Actual Size", Some("CmdOrCtrl+0"))?)
        .build()?;
    let layout = SubmenuBuilder::new(h, "Layout")
        .item(&item("margins", "Margins...", None)?)
        .item(&item("edit_header", "Edit Header", None)?)
        .item(&item("edit_footer", "Edit Footer", None)?)
        .build()?;
    let review = SubmenuBuilder::new(h, "Review")
        .item(&item("review", "Show or Hide Review", None)?)
        .item(&item("track", "Track Changes", None)?)
        .item(&item("new_comment", "New Comment", Some("CmdOrCtrl+Alt+M"))?)
        .separator()
        .item(&item("accept_all", "Accept All Changes...", None)?)
        .item(&item("reject_all", "Reject All Changes...", None)?)
        .build()?;
    let help = SubmenuBuilder::new(h, "Help")
        .item(&item("about", "About Folio", None)?)
        .build()?;
    #[allow(unused_mut)]
    let mut menu = MenuBuilder::new(h);
    #[cfg(target_os = "macos")]
    {
        let about = SubmenuBuilder::new(h, "Folio")
            .item(&PredefinedMenuItem::hide(h, None)?)
            .item(&PredefinedMenuItem::quit(h, None)?)
            .build()?;
        menu = menu.item(&about);
    }
    let menu = menu.item(&file).item(&edit).item(&view).item(&layout).item(&review).item(&help).build()?;
    app.set_menu(menu)?;
    Ok(())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let launch = std::env::args_os()
        .skip(1)
        .map(PathBuf::from)
        .find(|p| p.is_file());
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .manage(Granted::default())
        .manage(LaunchFile(Mutex::new(launch)))
        .setup(|app| {
            build_menu(app)?;
            Ok(())
        })
        .on_menu_event(|app, e| {
            let _ = app.emit("menu", e.id().as_ref());
            if let Some(w) = app.get_webview_window("main") {
                let _ = w.set_focus();
            }
        })
        .invoke_handler(tauri::generate_handler![
            pick_open_path,
            pick_image_path,
            pick_save_path,
            launch_path,
            read_document,
            write_document,
            libreoffice_available,
            convert_with_libreoffice,
            print_window,
            open_link
        ])
        .run(tauri::generate_context!())
        .expect("failed to start Folio");
}
