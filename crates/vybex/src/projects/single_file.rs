//! Loader for a single source file — wraps it into a one-file `Bundle`.

use std::path::Path;
use crate::bundle::{Bundle, EntryPoint, SourceFile};

pub fn load(path: &Path, ext: &str) -> Result<Bundle, String> {
    let lang = crate::languages::find_by_extension(ext).ok_or_else(|| {
        let exts = crate::projects::supported_extensions();
        let list: Vec<String> = exts.iter().map(|e| format!(".{e}")).collect();
        format!("Unknown file extension: .{ext}\nSupported: {}", list.join(", "))
    })?;

    let code = std::fs::read_to_string(path)
        .map_err(|e| format!("Cannot read {}: {}", path.display(), e))?;

    let name = path.file_stem()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| "main".into());

    Ok(Bundle {
        name,
        language: lang,
        sources: vec![SourceFile { path: path.to_path_buf(), code }],
        entry_point: EntryPoint::Auto,
    })
}
