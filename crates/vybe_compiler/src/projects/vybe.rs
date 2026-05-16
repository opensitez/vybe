//! Loader for `.vybe` project files (TOML format).
//!
//! ```toml
//! [project]
//! name = "My App"
//! entry = "main.vb"
//!
//! [sources]
//! files = ["main.vb", "utils.vb", "math.wasm"]
//!
//! [host]
//! gui = false
//! ```
//!
//! Language is auto-detected from the entry file extension.

use std::path::Path;
use crate::bundle::{Bundle, EntryPoint, SourceFile, WasmFile};

pub fn load(path: &Path) -> Result<Bundle, String> {
    let content = std::fs::read_to_string(path)
        .map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;

    let project_dir = path.parent().unwrap_or(Path::new("."));

    let root: toml::Value = toml::from_str(&content)
        .map_err(|e| format!("TOML parse error in {}: {}", path.display(), e))?;

    // [project]
    let project = root.get("project")
        .ok_or("Missing [project] section in .vybe file")?;

    let name = project.get("name")
        .and_then(|v| v.as_str())
        .unwrap_or("project")
        .to_string();

    let entry_str = project.get("entry")
        .and_then(|v| v.as_str())
        .unwrap_or("");

    let entry_point = if entry_str.is_empty() || entry_str.eq_ignore_ascii_case("auto") {
        EntryPoint::Auto
    } else {
        EntryPoint::Auto
    };

    // [sources]
    let files: Vec<String> = root.get("sources")
        .and_then(|s| s.get("files"))
        .and_then(|f| f.as_array())
        .map(|arr| arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
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

    // Detect language from entry file or first non-wasm source
    let detect_from = if !entry_str.is_empty() {
        entry_str.to_string()
    } else {
        file_list.iter()
            .find(|f| !f.ends_with(".wasm"))
            .cloned()
            .unwrap_or_default()
    };

    let lang_ext = Path::new(&detect_from)
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();

    let lang = crate::languages::find_by_extension(&lang_ext).ok_or_else(|| {
        format!("Cannot detect language from '{}' — unsupported extension '.{}'", detect_from, lang_ext)
    })?;

    // Load source files and WASM binaries
    let mut sources = Vec::new();
    let mut wasm_files = Vec::new();
    for rel_path in &file_list {
        let full_path = project_dir.join(rel_path);
        if rel_path.ends_with(".wasm") {
            let data = std::fs::read(&full_path)
                .map_err(|e| format!("Cannot read WASM '{}': {}", rel_path, e))?;
            wasm_files.push(WasmFile { path: full_path, data });
        } else {
            let code = std::fs::read_to_string(&full_path)
                .map_err(|e| format!("Cannot read '{}': {}", rel_path, e))?;
            sources.push(SourceFile { path: full_path, code });
        }
    }

    if sources.is_empty() {
        return Err("No compilable source files in project".into());
    }

    Ok(Bundle { name, language: lang, sources, wasm_files, entry_point })
}
