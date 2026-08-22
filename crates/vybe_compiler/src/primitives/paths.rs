//! Path STRING manipulation — `System.IO.Path`, `os.path`, `pathinfo`.
//!
//! # Why this is not in `fs_path.rs`, and not in WASI at all
//!
//! None of these touch the filesystem. `Path.GetExtension("a/b.txt")` is `".txt"`
//! whether or not that file exists, and asking the host would be both slower and
//! wrong — `.NET` answers from the string alone and so must we.
//!
//! WASI has no path interface, and that is deliberate rather than an omission:
//! path *syntax* is a language and platform concern (`.NET` accepts `\` as a
//! separator, POSIX does not; `Path.GetExtension` and `os.path.splitext`
//! disagree about a leading dot), while WASI's `-at` calls take a path as an
//! opaque string to be resolved by the host. So ten verbs — `pathCombine`,
//! `pathGetFileName`, `pathGetExtension`, … — sat in `wasi:filesystem` naming
//! functions that no WIT declares and no conforming host could serve, to do work
//! that never needed a host call in the first place.
//!
//! # ⚠Two of the ten are NOT pure, and binning them with the rest strands them
//!
//! [`emit_full_path`] resolves against the current directory and
//! [`emit_temp_path`] reads the environment. Those are real capability reads
//! that happen to return a string, so they go through `wasi:cli/environment`
//! rather than pretending to be string manipulation.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::primitives::{ops, strings};

/// The separator these verbs produce. Both `/` and `\` are ACCEPTED on input —
/// `.NET` code carries Windows paths and `Path.GetFileName(@"C:\a\b.txt")` must
/// still answer `b.txt` — but output is normalised, because the host resolves
/// against a POSIX preopen.
const SEP: &str = "/";

/// Index one past the last separator, or 0. Stack: `[]` → `[f64]`.
///
/// Both separators are searched and the later one wins, so a mixed path
/// (`C:\dir/file`) splits at the true last component rather than at whichever
/// flavour happened to be looked for first.
fn emit_last_sep(chunk: &mut Chunk, s: u16, line: u32) -> u16 {
    let fwd = chunk.alloc_scratch(1);
    let back = chunk.alloc_scratch(1);
    let cut = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("/", line);
    strings::emit_last_index_of(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, fwd, line);

    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("\\", line);
    strings::emit_last_index_of(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, back, line);

    chunk.emit_op_u16(Op::LOCAL_GET, fwd, line);
    chunk.emit_op_u16(Op::LOCAL_GET, back, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, back, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, fwd, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, cut, line);
    cut
}

/// `Path.GetFileName(p)` — everything after the last separator.
/// Stack: `[path]` → `[string]`.
pub fn emit_file_name(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    let cut = emit_last_sep(chunk, s, line);

    chunk.emit_op_u16(Op::LOCAL_GET, cut, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // No separator: the whole string IS the file name.
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, s, line);
        chunk.emit_op_u16(Op::LOCAL_GET, cut, line);
        chunk.emit_f64_const(1.0, line);
        ops::emit_dyn_add(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, s, line);
        strings::emit_length(chunk, line);
        strings::emit_substring(chunk, line);
    }
    chunk.emit_end(line);
}

/// `Path.GetDirectoryName(p)` — everything before the last separator, with the
/// separator dropped. `""` when there is none.
/// Stack: `[path]` → `[string]`.
pub fn emit_directory(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    let cut = emit_last_sep(chunk, s, line);

    chunk.emit_op_u16(Op::LOCAL_GET, cut, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("", line);
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, s, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op_u16(Op::LOCAL_GET, cut, line);
        strings::emit_substring(chunk, line);
    }
    chunk.emit_end(line);
}

/// Index of the extension's dot within the whole path, or `-1`.
///
/// Searched in the FILE NAME, not the path: `a.b/c` has no extension, and
/// looking for the last `.` in the full string would answer `.b/c`.
fn emit_ext_dot(chunk: &mut Chunk, s: u16, line: u32) -> u16 {
    let cut = emit_last_sep(chunk, s, line);
    let dot = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const(".", line);
    strings::emit_last_index_of(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dot, line);

    // A dot at or before the last separator belongs to a DIRECTORY name.
    chunk.emit_op_u16(Op::LOCAL_GET, dot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cut, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(-1.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dot, line);
    chunk.emit_end(line);
    dot
}

/// `Path.GetExtension(p)` — the last `.` in the file name and everything after
/// it, INCLUDING the dot. `""` when the file name has none.
/// Stack: `[path]` → `[string]`.
pub fn emit_extension(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    let dot = emit_ext_dot(chunk, s, line);

    chunk.emit_op_u16(Op::LOCAL_GET, dot, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("", line);
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, s, line);
        chunk.emit_op_u16(Op::LOCAL_GET, dot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, s, line);
        strings::emit_length(chunk, line);
        strings::emit_substring(chunk, line);
    }
    chunk.emit_end(line);
}

/// `Path.GetFileNameWithoutExtension(p)`. Stack: `[path]` → `[string]`.
pub fn emit_file_stem(chunk: &mut Chunk, line: u32) {
    emit_file_name(chunk, line);
    let name = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, name, line);
    let dot = emit_ext_dot(chunk, name, line);

    chunk.emit_op_u16(Op::LOCAL_GET, dot, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, name, line);
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, name, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op_u16(Op::LOCAL_GET, dot, line);
        strings::emit_substring(chunk, line);
    }
    chunk.emit_end(line);
}

/// `Path.HasExtension(p)`. Stack: `[path]` → `[bool]`.
pub fn emit_has_extension(chunk: &mut Chunk, line: u32) {
    emit_extension(chunk, line);
    strings::emit_length(chunk, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_gt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    ops::emit_i32_to_bool(chunk, line);
}

/// `Path.ChangeExtension(p, ext)`. Stack: `[path, ext]` → `[string]`.
///
/// `.NET` accepts the new extension with or without its dot and does not double
/// it, so both spellings are normalised to one leading dot.
pub fn emit_change_extension(chunk: &mut Chunk, line: u32) {
    let ext = chunk.alloc_scratch(1);
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, ext, line);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    let dot = emit_ext_dot(chunk, s, line);

    // base = path with any existing extension removed
    chunk.emit_op_u16(Op::LOCAL_GET, dot, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, s, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op_u16(Op::LOCAL_GET, dot, line);
        strings::emit_substring(chunk, line);
    }
    chunk.emit_end(line);

    // + "." + ext, with the caller's leading dot dropped if present
    chunk.emit_op_u16(Op::LOCAL_GET, ext, line);
    chunk.emit_string_const(".", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ext, line);
    chunk.emit_else(line);
    {
        chunk.emit_string_const(".", line);
        chunk.emit_op_u16(Op::LOCAL_GET, ext, line);
        strings::emit_concat(chunk, 2, line);
    }
    chunk.emit_end(line);
    strings::emit_concat(chunk, 2, line);
}

/// `Path.IsPathRooted(p)` — POSIX absolute, or a Windows drive/UNC prefix.
/// Stack: `[path]` → `[bool]`.
pub fn emit_is_rooted(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);
    // A leading separator of either flavour. A drive letter (`C:`) is a colon
    // at index 1, which is the same test `.NET` makes.
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("/", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("\\", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const(":", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(1.0, line);
    ops::emit_dyn_eq(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
    ops::emit_i32_to_bool(chunk, line);
}

/// `Path.Combine(a, b, …)` — join with one separator.
/// Stack: `[seg…]` → `[string]`.
///
/// A ROOTED segment discards everything to its left, which is `.NET`'s rule and
/// POSIX's: `Path.Combine("/tmp", "/etc")` is `/etc`, not `/tmp/etc`. Getting
/// that wrong is how a program that means to write to an absolute path quietly
/// writes inside a scratch directory instead.
pub fn emit_combine(chunk: &mut Chunk, argc: u8, line: u32) {
    let n = argc as u16;
    let base = chunk.alloc_scratch(n.max(1));
    for offset in (0..n).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    let acc = chunk.alloc_scratch(1);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

    for i in 0..n {
        let seg = base + i;
        // Skip empty segments outright — `Combine("a", "", "b")` is `a/b`.
        chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
        strings::emit_length(chunk, line);
        chunk.emit_f64_const(0.0, line);
        ops::emit_dyn_gt(chunk, line);
        ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        {
            chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
            emit_is_rooted(chunk, line);
            ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
            strings::emit_length(chunk, line);
            chunk.emit_f64_const(0.0, line);
            ops::emit_dyn_eq(chunk, line);
            ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_op(Op::I32_OR, line);
            chunk.emit_if_value(line);
            // rooted, or nothing accumulated yet: this segment replaces
            chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
            chunk.emit_else(line);
            {
                // The separator goes in only if the accumulator does not
                // already end with one. `Path.Combine("/a/b/", "c")` is
                // `/a/b/c` in .NET, and `os.path.join`/`path.join` agree — no
                // consumer wants the doubled slash, so this is a bug rather
                // than a dialect. It was visible as `/var/.../T//name`, which
                // still resolves on POSIX and so hid until a test compared
                // the STRING.
                chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
                // "ends with SEP" as `lastIndexOf(SEP) == length - 1`, built
                // from helpers this module already uses. The empty-accumulator
                // case cannot arrive here — the branch above claims it — which
                // matters, because for `""` both sides are -1 and this would
                // answer "already separated".
                chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
                chunk.emit_string_const(SEP, line);
                strings::emit_last_index_of(chunk, line);
                chunk.emit_f64_const(1.0, line);
                ops::emit_dyn_add(chunk, line);
                chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
                strings::emit_length(chunk, line);
                ops::emit_dyn_eq(chunk, line);
                ops::emit_dyn_to_bool(chunk, line);
                chunk.emit_if_value(line);
                chunk.emit_string_const("", line);
                chunk.emit_else(line);
                chunk.emit_string_const(SEP, line);
                chunk.emit_end(line);
                chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
                strings::emit_concat(chunk, 3, line);
            }
            chunk.emit_end(line);
            chunk.emit_op_u16(Op::LOCAL_SET, acc, line);
        }
        chunk.emit_end(line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
}

// ── The two that are capability reads, not string manipulation ────────────

/// `Path.GetFullPath(p)` — absolute form, resolved against the process's
/// initial working directory. Stack: `[path]` → `[string]`.
///
/// NOT a string op: it needs `wasi:cli/environment.get-initial-cwd`, which
/// answers `option<string>`. A relative path with no cwd available is left
/// as-is rather than guessed at.
pub fn emit_full_path(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(1);
    let cwd = chunk.alloc_scratch(1);
    let get_cwd = chunk.add_import("wasi:cli/environment", "get-initial-cwd");
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);

    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    emit_is_rooted(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_else(line);
    {
        chunk.emit_call(get_cwd, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, cwd, line);
        chunk.emit_op_u16(Op::LOCAL_GET, cwd, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, s, line);
        chunk.emit_else(line);
        {
            chunk.emit_op_u16(Op::LOCAL_GET, cwd, line);
            chunk.emit_string_const(SEP, line);
            chunk.emit_op_u16(Op::LOCAL_GET, s, line);
            strings::emit_concat(chunk, 3, line);
        }
        chunk.emit_end(line);
    }
    chunk.emit_end(line);
}

/// `Path.GetTempPath()` — the scratch directory, from the environment.
/// Stack: `[]` → `[string]`.
///
/// Also NOT a string op: it reads `TMPDIR` out of
/// `wasi:cli/environment.get-environment`, whose shape is
/// `list<tuple<string, string>>`. Falls back to `/tmp`.
pub fn emit_temp_path(chunk: &mut Chunk, line: u32) {
    let env = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);
    let n = chunk.alloc_scratch(1);
    let pair = chunk.alloc_scratch(1);
    let out = chunk.alloc_scratch(1);
    let get_env = chunk.add_import("wasi:cli/environment", "get-environment");
    let len = chunk.add_import("ecma:array", "length");
    let to_i32 = chunk.add_import("wasm:js-number", "toI32");

    chunk.emit_string_const("/tmp", line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);

    chunk.emit_call(get_env, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, env, line);
    chunk.emit_op_u16(Op::LOCAL_GET, env, line);
    chunk.emit_call(len, 1, line);
    chunk.emit_call(to_i32, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);

    let done = chunk.emit_block(line);
    let (scan, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, env, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pair, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pair, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_string_const("TMPDIR", line);
    strings::emit_str_equals(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, pair, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(scan);
    chunk.emit_end(line);
    chunk.patch_block(done);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}
