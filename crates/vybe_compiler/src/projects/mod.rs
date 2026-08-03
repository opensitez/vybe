//! Project loading — takes any path, returns a `Bundle`.
//!
//! Dispatches by file extension to the appropriate loader.
//! Single source files are also handled here (a project of one).
//!
//! Also provides the IDE project model (`Project`, `FormModule`, etc.)
//! for the code editor — load, edit, and save `.vbproj` projects with
//! full form designer round-trip support.

pub mod encoding;
mod managed_msbuild;
mod single_file;
mod vbproj;
mod vybe;

pub use encoding::read_text_file;

use crate::bundle::Bundle;
use std::path::Path;

/// One or more bundles that run together in a single VM.
///
/// Multi-language is why this type exists. A `.vybe` project may list
/// `main.vb` beside `math_utils.js`, and those cannot share one parse — each
/// language has its own grammar, profile and injected prelude, so feeding both
/// to one front-end is what produced "expected EOI" on line 21 of the JS file.
/// Each language group is therefore its own [`Bundle`], compiled on its own
/// terms, and the units are linked where linking actually belongs: in the VM,
/// sharing globals and host functions.
///
/// `units` is ordered — every secondary language first, the entry language
/// LAST — so libraries are defined before the program that uses them.
pub struct Program {
    pub units: Vec<Bundle> }

impl Program {
    /// A single-language program.
    pub fn single(bundle: Bundle) -> Self {
        Program { units: vec![bundle] }
    }

    /// The entry-language unit — the one whose entry point starts the program.
    pub fn entry(&self) -> &Bundle {
        self.units.last().expect("a program has at least one unit")
    }

    pub fn into_entry(mut self) -> Bundle {
        self.units.pop().expect("a program has at least one unit")
    }

    /// The units that run before the entry unit.
    pub fn secondaries(&self) -> &[Bundle] {
        &self.units[..self.units.len() - 1]
    }

    pub fn is_multi_language(&self) -> bool {
        self.units.len() > 1
    }
}

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
        _ => single_file::load(path, &ext) }
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

/// Load one or more paths as a [`Program`] — the multi-language form of
/// [`load_many`].
///
/// Only `.vybe` projects and bare source paths can span languages; the
/// MSBuild-family project formats describe a single language by construction,
/// so they yield a one-unit program.
pub fn load_program(paths: &[std::path::PathBuf]) -> Result<Program, String> {
    match paths {
        [] => Err("no input files".to_string()),
        [single] => {
            let ext = single
                .extension()
                .and_then(|e| e.to_str())
                .map(|e| e.to_lowercase())
                .unwrap_or_default();
            match ext.as_str() {
                "vybe" => vybe::load_program(single),
                _ => load(single).map(Program::single) }
        }
        _ => {
            reject_project_files(paths)?;
            single_file::load_many_program(paths)
        }
    }
}

fn reject_project_files(paths: &[std::path::PathBuf]) -> Result<(), String> {
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
    Ok(())
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
