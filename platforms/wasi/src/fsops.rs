//! The filesystem operations themselves — **one implementation, several
//! surfaces**.
//!
//! There were THREE independent filesystem implementations in this host, each
//! calling `std::fs` directly and each free to drift from the others:
//!
//! | file | lines | surface |
//! |---|---|---|
//! | `filesystem.rs` | 1471 | `wasi:filesystem/types` — the SPEC one, descriptor-based, 0.3.1 |
//! | `fs.rs` | 657 | `readFile`/`writeFile`/`lineInput`/`pathCombine` — INVENTED names under the `wasi:filesystem` package, and what the compiler actually emits |
//! | `../../node/src/fs.rs` | 1080 | `node:fs` — Node shape, `std::fs` direct |
//!
//! Same split as `primitives/io.rs` made for output: the QUIRKS live in each
//! surface (Node throws an `Error` with a `code`; WASI answers an
//! `error-code`; the invented layer answers a string), and there is ONE thing
//! underneath doing the work. `node:crypto` already composes this way —
//! `use vybe_platform_wasi::crypto::{md5_hex, sha256_hex}` — so the layering
//! is the established pattern here, not a new one.
//!
//! Nothing in this module registers a host function or knows what a `Value`
//! is. It is deliberately the boring layer: paths in, bytes and metadata out,
//! `std::io::Error` on failure so each surface can shape the error its own
//! way.

use std::fs;
use std::io::Write;
use std::path::Path;

/// Whole-file read.
pub fn read(path: &str) -> std::io::Result<Vec<u8>> {
    fs::read(path)
}

/// Whole-file write, truncating.
pub fn write(path: &str, data: &[u8]) -> std::io::Result<()> {
    fs::write(path, data)
}

/// Append, creating the file when absent.
pub fn append(path: &str, data: &[u8]) -> std::io::Result<()> {
    let mut file = fs::OpenOptions::new().create(true).append(true).open(path)?;
    file.write_all(data)
}

/// Positioned write — the operation behind 0.3.1's
/// `descriptor.write-via-stream(data, offset)`.
///
/// Writing past the end extends the file and zero-fills the gap, which is what
/// `File::seek` past EOF plus a write already does and what the WIT requires.
pub fn write_at(path: &str, data: &[u8], offset: u64) -> std::io::Result<()> {
    use std::io::{Seek, SeekFrom};
    let mut file = fs::OpenOptions::new().write(true).open(path)?;
    file.seek(SeekFrom::Start(offset))?;
    file.write_all(data)
}

/// Positioned read of at most `len` bytes — the operation behind
/// `descriptor.read-via-stream(offset)`.
///
/// A SHORT read is not an error and not necessarily EOF: it means the file had
/// fewer bytes left than asked for. Callers that need "no such record" must
/// compare the length themselves rather than reading a flag, because the two
/// are genuinely different conditions.
pub fn read_at(path: &str, offset: u64, len: usize) -> std::io::Result<Vec<u8>> {
    use std::io::{Read, Seek, SeekFrom};
    let mut file = fs::File::open(path)?;
    file.seek(SeekFrom::Start(offset))?;
    let mut buf = vec![0u8; len];
    let n = file.read(&mut buf)?;
    buf.truncate(n);
    Ok(buf)
}

pub fn metadata(path: &str) -> std::io::Result<fs::Metadata> {
    fs::metadata(path)
}

/// `lstat` — metadata WITHOUT following a final symlink.
pub fn symlink_metadata(path: &str) -> std::io::Result<fs::Metadata> {
    fs::symlink_metadata(path)
}

pub fn exists(path: &str) -> bool {
    Path::new(path).exists()
}

pub fn remove_file(path: &str) -> std::io::Result<()> {
    fs::remove_file(path)
}

pub fn remove_dir_all(path: &str) -> std::io::Result<()> {
    fs::remove_dir_all(path)
}

pub fn create_dir_all(path: &str) -> std::io::Result<()> {
    fs::create_dir_all(path)
}

pub fn rename(from: &str, to: &str) -> std::io::Result<()> {
    fs::rename(from, to)
}

pub fn copy(from: &str, to: &str) -> std::io::Result<u64> {
    fs::copy(from, to)
}

/// Directory entry names, unsorted — the order the OS gives them.
pub fn read_dir_names(path: &str) -> std::io::Result<Vec<String>> {
    let mut names = Vec::new();
    for entry in fs::read_dir(path)? {
        names.push(entry?.file_name().to_string_lossy().into_owned());
    }
    Ok(names)
}

/// The POSIX `errno` NAME for an I/O error — `"ENOENT"`, `"EACCES"`, ….
///
/// This is the vocabulary **Node** uses: a thrown `fs` error carries
/// `err.code === 'ENOENT'`, and real programs branch on exactly that string.
/// WASI's `error-code` vocabulary is different (`no-entry`, `access`, …) and
/// lives in `filesystem.rs::map_io_error` — the two must NOT be merged, since
/// each is the correct answer for its own surface. Same error, two spellings,
/// which is precisely the kind of thing that belongs in the surface and not
/// down here.
pub fn errno_name(e: &std::io::Error) -> &'static str {
    use std::io::ErrorKind::*;
    match e.kind() {
        NotFound => "ENOENT",
        PermissionDenied => "EACCES",
        AlreadyExists => "EEXIST",
        WouldBlock => "EAGAIN",
        InvalidInput | InvalidData => "EINVAL",
        BrokenPipe => "EPIPE",
        Interrupted => "EINTR",
        Unsupported => "ENOTSUP",
        OutOfMemory => "ENOMEM",
        // ENOTEMPTY / EISDIR / ENOTDIR have no stable `ErrorKind` variant, so
        // they are recovered from the raw OS number — different per platform.
        _ => match e.raw_os_error() {
            Some(39) | Some(66) | Some(145) => "ENOTEMPTY",
            Some(21) => "EISDIR",
            Some(20) => "ENOTDIR",
            Some(28) => "ENOSPC",
            _ => "EIO",
        },
    }
}
