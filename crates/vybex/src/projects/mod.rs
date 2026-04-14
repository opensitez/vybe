//! Project loading — takes any path, returns a `Bundle`.
//!
//! Dispatches by file extension to the appropriate loader.
//! Single source files are also handled here (a project of one).

mod single_file;
mod vybe;
mod vbproj;

use std::path::Path;
use crate::bundle::Bundle;

/// Load any supported file or project → `Bundle`.
pub fn load(path: &Path) -> Result<Bundle, String> {
    let ext = path.extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    match ext.as_str() {
        "vybe"   => vybe::load(path),
        "vbproj" => vbproj::load(path),
        _        => single_file::load(path, &ext),
    }
}

/// List all supported extensions (languages + project formats).
pub fn supported_extensions() -> Vec<String> {
    let mut exts = crate::languages::supported_extensions();
    exts.extend_from_slice(&["vybe".into(), "vbproj".into()]);
    exts
}
