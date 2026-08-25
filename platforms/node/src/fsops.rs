//! The filesystem operations themselves — paths in, bytes and metadata out.
//!
//! Lives in `platforms/node` because its only surface is `node:fs`. It sat in
//! `platforms/wasi` when three filesystem implementations shared it, and the
//! other two are gone: the invented-verb layer (`wasi/src/fs.rs`,
//! `readFile`/`writeFile`/`lineInput`/`pathCombine` under the `wasi:filesystem`
//! PACKAGE name) is deleted, and `wasi:filesystem/types` does not compose this
//! at all — 0.3.1 is descriptor- and STREAM-based, where this is path-based, so
//! there is no shared layer left to justify the cross-platform home.
//!
//! Nothing here registers a host function or knows what a `Value` is. It is
//! deliberately the boring layer: `std::io::Error` on failure, so the surface
//! above shapes the error its own way (Node throws an `Error` with a `code`).

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
