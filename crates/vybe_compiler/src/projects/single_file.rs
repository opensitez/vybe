//! Loader for source files given directly on the command line.
//!
//! One file wraps into a one-file `Bundle`; several files link together into a
//! single multi-source `Bundle`, the same shape the project loaders
//! (`.vybe`/`.vbproj`/`.csproj`) produce — so `vybex a.c b.c` links like a C
//! compiler invocation rather than compiling only the first file.

use crate::bundle::{Bundle, EntryPoint, SourceFile};
use std::path::Path;

pub fn load(path: &Path, ext: &str) -> Result<Bundle, String> {
    let lang = crate::languages::find_by_extension(ext).ok_or_else(|| {
        let exts = crate::projects::supported_extensions();
        let list: Vec<String> = exts.iter().map(|e| format!(".{e}")).collect();
        format!(
            "Unknown file extension: .{ext}\nSupported: {}",
            list.join(", ")
        )
    })?;

    let code = std::fs::read_to_string(path)
        .map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;

    let name = path
        .file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "main".into());

    Ok(Bundle {
        name,
        language: lang,
        sources: vec![SourceFile {
            path: path.to_path_buf(),
            code,
        }],
        wasm_files: vec![],
        entry_point: EntryPoint::Auto,
    })
}

/// Load several source files as ONE bundle, linked together.
///
/// The first path is the entry file: it names the bundle and fixes the
/// language. Every other file must share that language — mixing front-ends in
/// a single link step has no defined semantics (each language lowers through
/// its own walker/profile), so it's rejected with a clear error rather than
/// silently compiling the first one.
///
/// `EntryPoint::Auto` still applies: the entry is inferred from the code
/// (`main()`, `Sub Main`, top-level statements) exactly as for one file.
/// Load several source files as a [`Program`], one unit per language.
///
/// `vybex main.vb math_utils.js` used to be rejected outright ("all source
/// files in one command must be the same language"). Mixing front-ends in a
/// single *parse* still has no meaning — each language lowers through its own
/// walker and profile — so instead of concatenating them, each language group
/// is compiled on its own and the units are linked in the VM.
///
/// The first path is the entry file: it names the program and fixes the entry
/// language, whose unit runs last.
pub fn load_many_program(paths: &[std::path::PathBuf]) -> Result<crate::projects::Program, String> {
    let Some((first, _)) = paths.split_first() else {
        return Err("no source files given".to_string());
    };

    let entry_lang = super::vybe::language_of(&first.to_string_lossy())?;
    let name = first
        .file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "main".into());

    let mut read = Vec::new();
    for path in paths {
        let code = std::fs::read_to_string(path)
            .map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;
        read.push((path.to_path_buf(), code));
    }

    let grouped = super::vybe::group_by_language(read)?;
    Ok(super::vybe::build_program(
        name,
        entry_lang.name,
        grouped,
        Vec::new(),
    ))
}

pub fn load_many(paths: &[std::path::PathBuf]) -> Result<Bundle, String> {
    let Some((first, rest)) = paths.split_first() else {
        return Err("no source files given".to_string());
    };

    let ext_of = |p: &Path| -> String {
        p.extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_lowercase())
            .unwrap_or_default()
    };

    // The entry file fixes the language for the whole link step.
    let entry_ext = ext_of(first);
    let mut bundle = load(first, &entry_ext)?;

    for path in rest {
        let ext = ext_of(path);
        let lang = crate::languages::find_by_extension(&ext)
            .ok_or_else(|| format!("Unknown file extension: .{ext} ({})", path.display()))?;
        if lang.name != bundle.language.name {
            return Err(format!(
                "Cannot link {} ({}) with {} ({}): all source files in one command must \
be the same language. Use a project file (.vybe) to combine languages.",
                path.display(),
                lang.name,
                first.display(),
                bundle.language.name,
            ));
        }
        let code = std::fs::read_to_string(path)
            .map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;
        bundle.sources.push(SourceFile {
            path: path.to_path_buf(),
            code,
        });
    }

    Ok(bundle)
}
