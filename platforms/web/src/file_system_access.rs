//! **File System Access** — WHATWG's picker API, `showOpenFilePicker()`,
//! `showSaveFilePicker()` and `showDirectoryPicker()`.
//!
//! The three ways a page asks the user agent for a file, and the web's answer
//! to what WinForms spells `OpenFileDialog` / `SaveFileDialog` /
//! `FolderBrowserDialog`. A picker is not something a page draws: it is the
//! agent's own chrome, which is why these are host functions rather than
//! elements — the same reason `<dialog>` IS an element and this is not.
//!
//! ## Why here
//!
//! These used to be registered by `crates/vybex` — the RUNNER — reaching
//! `vybe_widgets::dialogs` directly. The runner is the user agent: it should
//! not be publishing page APIs, because then two crates own the browser's
//! surface and neither owns it fully.
//!
//! `platforms/web` owns the relationship to `vybe_widgets`. Everything a page
//! can call belongs here.
//!
//! ## Where this diverges from the spec, deliberately
//!
//! The real API answers `FileSystemFileHandle` objects and is asynchronous —
//! a Promise, because a picker blocks on a human. These answer the PATH as a
//! string, synchronously, because that is what the callers need today: a
//! WinForms `FileName` property is a path, and the guest has no Promise to
//! await. The handle objects are the spec-shaped upgrade and would sit on top
//! of exactly these calls.
//!
//! Recorded rather than hidden: a program written against the real API would
//! not work here yet. What it would NOT do is silently get a wrong answer.

use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

use vybe_widgets::dialogs::{FileDialog, FileFilter, FolderDialog};

/// `accept` as the spec spells it is a list of `{description, accept}` entries;
/// a WinForms `Filter` is `"Text|*.txt|All|*.*"`. Both reduce to the same pair,
/// so the string form is accepted directly and anything richer is ignored
/// rather than half-read.
fn filters_from(spec: &str) -> Vec<FileFilter> {
    let parts: Vec<&str> = spec.split('|').collect();
    let mut out = Vec::new();
    for pair in parts.chunks(2) {
        let [name, patterns] = pair else { continue };
        let extensions: Vec<&str> = patterns
            .split(';')
            .filter_map(|p| p.trim().rsplit('.').next())
            .filter(|e| !e.is_empty() && *e != "*")
            .collect();
        if !extensions.is_empty() {
            out.push(FileFilter::new(name.to_string(), &extensions));
        }
    }
    out
}

fn string_arg(args: &[Value], index: usize) -> String {
    match args.get(index) {
        Some(Value::String(s)) => s.to_string(),
        Some(Value::Undefined) | Some(Value::Null) | None => String::new(),
        Some(other) => format!("{}", other),
    }
}

/// A picker the user dismissed answers NOTHING, and that is the whole of the
/// error path: the spec rejects with `AbortError`, and `null` is the value a
/// synchronous caller can test. `""` would be a path.
fn path_or_null(path: Option<std::path::PathBuf>) -> Value {
    match path {
        Some(p) => Value::String(p.to_string_lossy().to_string().into()),
        None => Value::Null,
    }
}

fn build(title: &str, filter: &str, directory: &str) -> FileDialog {
    let mut dialog = FileDialog::new(if title.is_empty() { "Open" } else { title });
    for f in filters_from(filter) {
        dialog = dialog.with_filter(f);
    }
    if !directory.is_empty() {
        dialog = dialog.with_starting_directory(directory);
    }
    dialog
}

/// `showOpenFilePicker` as a Rust call.
///
/// Public because the RUNNER needs the same three pickers, and it must reach
/// them through this crate rather than around it: `platforms/web` owns the
/// relationship to `vybe_widgets`, and a second crate calling
/// `vybe_widgets::dialogs` directly is how that ownership stops being true.
pub fn open_file(title: &str, filter: &str, directory: &str) -> Option<std::path::PathBuf> {
    build(title, filter, directory).open()
}

/// `showSaveFilePicker` as a Rust call.
pub fn save_file(
    title: &str,
    filter: &str,
    directory: &str,
    suggested: &str,
) -> Option<std::path::PathBuf> {
    let mut dialog = build(title, filter, directory);
    if !suggested.is_empty() {
        dialog = dialog.with_filename(suggested);
    }
    dialog.save()
}

/// `showDirectoryPicker` as a Rust call.
pub fn pick_directory(title: &str, directory: &str) -> Option<std::path::PathBuf> {
    let mut dialog = FolderDialog::new(if title.is_empty() {
        "Select Folder"
    } else {
        title
    });
    if !directory.is_empty() {
        dialog = dialog.with_starting_directory(directory);
    }
    dialog.pick()
}

pub fn register(vm: &mut VM) {
    // `showOpenFilePicker(options)` — one file, or a list when `multiple`.
    //
    // The arguments are flattened rather than taking the spec's options
    // dictionary, because the emitters that call this have the values as
    // separate properties already and packing them into an object here only to
    // unpack them there would be ceremony with a chance of disagreeing.
    vm.register_host_fn(
        "web:file-system-access",
        "showOpenFilePicker",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let (title, filter, directory) = (
                string_arg(args, 0),
                string_arg(args, 1),
                string_arg(args, 2),
            );
            let multiple = matches!(args.get(3), Some(Value::Bool(true)));
            if !multiple {
                return path_or_null(open_file(&title, &filter, &directory));
            }
            let dialog = build(&title, &filter, &directory);
            // The spec always answers a LIST here; the single case above is the
            // convenience the callers actually use.
            let items: Vec<Value> = dialog
                .open_multiple()
                .unwrap_or_default()
                .into_iter()
                .map(|p| Value::String(p.to_string_lossy().to_string().into()))
                .collect();
            Value::Object(vybe_runtime::heap::alloc(Object::new_array(items)))
        }),
    );

    vm.register_host_fn(
        "web:file-system-access",
        "showSaveFilePicker",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            path_or_null(save_file(
                &string_arg(args, 0),
                &string_arg(args, 1),
                &string_arg(args, 2),
                &string_arg(args, 3),
            ))
        }),
    );

    vm.register_host_fn(
        "web:file-system-access",
        "showDirectoryPicker",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            path_or_null(pick_directory(
                &string_arg(args, 0),
                &string_arg(args, 1),
            ))
        }),
    );
}
