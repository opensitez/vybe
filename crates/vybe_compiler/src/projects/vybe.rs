//! Loader for `.vybe` project files (TOML format).
//!
//! ```toml
//! [project]
//! name = "My App"
//! entry = "main.vb"
//!
//! [sources]
//! files = ["main.vb", "math_utils.js", "math.wasm"]
//!
//! [host]
//! gui = false
//! ```
//!
//! The entry file's extension picks the *entry* language. Sources in any other
//! language become their own units of the resulting [`Program`] — they cannot
//! share one parse, because each language has its own grammar, profile and
//! prelude. Linking happens in the VM at run time (shared globals and host),
//! never by concatenating source text across languages.

use super::Program;
use crate::bundle::{Bundle, EntryPoint, SourceFile, WasmFile};
use std::path::{Path, PathBuf};

/// The entry-language unit alone. Kept for callers that compile or inspect a
/// single bundle (`--dump`, `--emit-wasm`, hot reload, the debugger's
/// evaluator); use [`load_program`] to get the other languages too.
pub fn load(path: &Path) -> Result<Bundle, String> {
    Ok(load_program(path)?.into_entry())
}

pub fn load_program(path: &Path) -> Result<Program, String> {
    let content = std::fs::read_to_string(path)
        .map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;

    let project_dir = path.parent().unwrap_or(Path::new("."));

    let root: toml::Value = toml::from_str(&content)
        .map_err(|e| format!("TOML parse error in {}: {}", path.display(), e))?;

    // [project]
    let project = root
        .get("project")
        .ok_or("Missing [project] section in .vybe file")?;

    let name = project
        .get("name")
        .and_then(|v| v.as_str())
        .unwrap_or("project")
        .to_string();

    let entry_str = project.get("entry").and_then(|v| v.as_str()).unwrap_or("");

    // [sources]
    let files: Vec<String> = root
        .get("sources")
        .and_then(|s| s.get("files"))
        .and_then(|f| f.as_array())
        .map(|arr| {
            arr.iter()
                .filter_map(|v| v.as_str().map(|s| s.to_string()))
                .collect()
        })
        .unwrap_or_default();

    // If no [sources].files, use the entry file alone
    let file_list = if files.is_empty() && !entry_str.is_empty() {
        vec![entry_str.to_string()]
    } else {
        files
    };

    if file_list.is_empty() {
        return Err("No source files specified in .vybe project".into());
    }

    // The entry file names the entry language; without one, the first
    // compilable source does.
    let detect_from = if !entry_str.is_empty() {
        entry_str.to_string()
    } else {
        file_list
            .iter()
            .find(|f| !f.ends_with(".wasm"))
            .cloned()
            .unwrap_or_default()
    };
    let entry_lang = language_of(&detect_from)?;

    // Read every listed file, keeping WASM binaries aside — they are already
    // compiled, carry no language, and link as chunks on the entry unit.
    let mut wasm_files = Vec::new();
    let mut grouped: Vec<(String, Vec<SourceFile>)> = Vec::new();
    for rel_path in &file_list {
        let full_path = project_dir.join(rel_path);
        if rel_path.ends_with(".wasm") {
            let data = std::fs::read(&full_path)
                .map_err(|e| format!("Cannot read WASM '{}': {}", rel_path, e))?;
            wasm_files.push(WasmFile {
                path: full_path,
                data,
            });
            continue;
        }
        let code = std::fs::read_to_string(&full_path)
            .map_err(|e| format!("Cannot read '{}': {}", rel_path, e))?;
        let lang = language_of(rel_path)?;
        let source = SourceFile {
            path: full_path,
            code,
        };
        match grouped.iter_mut().find(|(n, _)| *n == lang.name) {
            Some((_, sources)) => sources.push(source),
            None => grouped.push((lang.name.to_string(), vec![source])),
        }
    }

    if grouped.is_empty() {
        return Err("No compilable source files in project".into());
    }

    Ok(build_program(name, entry_lang.name, grouped, wasm_files))
}

/// Assemble the ordered unit list: every secondary language in the order it
/// first appears, then the entry language LAST, so a library is defined before
/// the program that uses it.
pub(super) fn build_program(
    name: String,
    entry_lang: &str,
    grouped: Vec<(String, Vec<SourceFile>)>,
    wasm_files: Vec<WasmFile>,
) -> Program {
    let mut units = Vec::new();
    let mut entry_unit = None;

    for (lang_name, sources) in grouped {
        let language = crate::languages::find_by_name(&lang_name)
            .expect("language was resolved when the group was built");
        let is_entry = lang_name == entry_lang;
        let bundle = Bundle {
            name: if is_entry {
                name.clone()
            } else {
                format!("{name}:{lang_name}")
            },
            language,
            sources,
            // Pre-compiled WASM links onto the entry unit; it is
            // language-neutral and appending it twice would duplicate chunks.
            wasm_files: if is_entry {
                wasm_files.clone()
            } else {
                Vec::new()
            },
            entry_point: EntryPoint::Auto,
        };
        if is_entry {
            entry_unit = Some(bundle);
        } else {
            units.push(bundle);
        }
    }

    if let Some(entry) = entry_unit {
        units.push(entry);
    }
    Program { units }
}

pub(super) fn language_of(file: &str) -> Result<crate::languages::Language, String> {
    let ext = Path::new(file)
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();
    crate::languages::find_by_extension(&ext).ok_or_else(|| {
        format!("Cannot detect language from '{file}' — unsupported extension '.{ext}'")
    })
}

/// Group already-read sources by language, preserving first-appearance order.
pub(super) fn group_by_language(
    sources: Vec<(PathBuf, String)>,
) -> Result<Vec<(String, Vec<SourceFile>)>, String> {
    let mut grouped: Vec<(String, Vec<SourceFile>)> = Vec::new();
    for (path, code) in sources {
        let lang = language_of(&path.to_string_lossy())?;
        let source = SourceFile { path, code };
        match grouped.iter_mut().find(|(n, _)| *n == lang.name) {
            Some((_, group)) => group.push(source),
            None => grouped.push((lang.name.to_string(), vec![source])),
        }
    }
    Ok(grouped)
}
