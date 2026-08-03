//! Directory mode URL → filesystem path resolution.
//!
//! Rules:
//! 1. Strip query string, percent-decode the path.
//! 2. Reject `..` segments and encoded `%2F` slashes.
//! 3. Join to the configured root, canonicalize, reject if the result
//!    escapes the root (symlink / `..` bypass).
//! 4. If the resolved path is a file, return it.
//! 5. If it's a directory, try each configured index filename in order.
//! 6. Otherwise return `Resolution::NotFound`.

use std::path::{Path, PathBuf};

use super::config::ServeConfig;

#[derive(Debug)]
pub enum Resolution {
    File(PathBuf),
    NotFound,
    Forbidden }

pub fn resolve(url_path: &str, config: &ServeConfig) -> Resolution {
    let decoded = match percent_encoding::percent_decode_str(url_path).decode_utf8() {
        Ok(s) => s.into_owned(),
        Err(_) => return Resolution::Forbidden };

    // Reject encoded / and parent traversal segments before canonicalize.
    if url_path.contains("%2f") || url_path.contains("%2F") {
        return Resolution::Forbidden;
    }
    for seg in decoded.split('/') {
        if seg == ".." {
            return Resolution::Forbidden;
        }
    }

    let rel = decoded.trim_start_matches('/');
    let candidate = config.root.join(rel);

    let canonical = match candidate.canonicalize() {
        Ok(p) => p,
        Err(_) => return Resolution::NotFound };
    let canonical_root = match config.root.canonicalize() {
        Ok(p) => p,
        Err(_) => return Resolution::NotFound };

    if !canonical.starts_with(&canonical_root) {
        return Resolution::Forbidden;
    }

    if canonical.is_file() {
        return Resolution::File(canonical);
    }

    if canonical.is_dir() {
        for index in &config.index_files {
            let candidate = canonical.join(index);
            if candidate.is_file() {
                return Resolution::File(candidate);
            }
        }
        return Resolution::NotFound;
    }

    Resolution::NotFound
}

/// Script extensions that trigger compile-and-run dispatch.
///
/// Directory mode serves browser assets from mixed webroots, so `.js`
/// and `.mjs` stay on the static-file path rather than being executed as
/// server scripts.
pub fn is_script(path: &Path) -> bool {
    let ext = path.extension().and_then(|e| e.to_str()).unwrap_or("");
    matches!(
        ext,
        "php" | "vb" | "py" | "rb" | "cs" | "dart" | "cob" | "cbl" | "f90" | "pas"
    )
}
