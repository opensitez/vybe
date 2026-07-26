//! Project loading — takes any path, returns a `Bundle`.
//!
//! Dispatches by file extension to the appropriate loader.
//! Single source files are also handled here (a project of one).
//!
//! Also provides the IDE project model (`Project`, `FormModule`, etc.)
//! for the code editor — load, edit, and save `.vbproj` projects with
//! full form designer round-trip support.

pub mod encoding;
pub mod form_modules;
mod managed_msbuild;
mod single_file;
pub mod vbforms;
mod vbproj;
mod vybe;

pub use encoding::read_text_file;
pub use vbforms::*;

use crate::bundle::Bundle;
use std::path::Path;

/// Load any supported file or project → `Bundle`.
pub fn load(path: &Path) -> Result<Bundle, String> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    match ext.as_str() {
        "vybe" => vybe::load(path),
        "vbproj" => vbproj::load(path),
        "csproj" | "pyproj" | "ipyproj" => managed_msbuild::load(path),
        _ => single_file::load(path, &ext),
    }
}

/// Load one or more paths given on the command line.
///
/// A single path keeps the existing behaviour exactly (project files dispatch
/// by extension; a bare source file becomes a one-file bundle). Several paths
/// link together into one multi-source bundle via
/// [`single_file::load_many`] — the entry file is first.
///
/// Project files describe their own source list, so they cannot be combined
/// with additional paths on the command line.
pub fn load_many(paths: &[std::path::PathBuf]) -> Result<Bundle, String> {
    match paths {
        [] => Err("no input files".to_string()),
        [single] => load(single),
        _ => {
            for path in paths {
                let ext = path
                    .extension()
                    .and_then(|e| e.to_str())
                    .map(|e| e.to_lowercase())
                    .unwrap_or_default();
                if matches!(
                    ext.as_str(),
                    "vybe" | "vbproj" | "csproj" | "pyproj" | "ipyproj"
                ) {
                    return Err(format!(
                        "{} is a project file and already lists its sources; \
pass it on its own.",
                        path.display()
                    ));
                }
            }
            single_file::load_many(paths)
        }
    }
}

/// List all supported extensions (languages + project formats).
pub fn supported_extensions() -> Vec<String> {
    let mut exts = crate::languages::supported_extensions();
    exts.extend_from_slice(&[
        "vybe".into(),
        "vbproj".into(),
        "csproj".into(),
        "pyproj".into(),
        "ipyproj".into(),
    ]);
    exts
}
