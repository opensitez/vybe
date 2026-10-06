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

/// Index of the last separator, or -1. Stored as a numeric UTF-16 index.
///
/// Both separators are searched and the later one wins, so a mixed path
/// (`C:\dir/file`) splits at the true last component rather than at whichever
/// flavour happened to be looked for first.
fn emit_last_sep(chunk: &mut Chunk, s: u16, line: u32) -> u16 {
    let cut = chunk.alloc_scratch(1);
    // Both ECMA searches return numeric UTF-16 indices (or -1), never NaN.
    // Keep the first result on the stack while obtaining the second one.
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("/", line);
    strings::emit_last_index_of(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("\\", line);
    strings::emit_last_index_of(chunk, line);
    chunk.emit_op(Op::F64_MAX, line);
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
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    // No separator: the whole string IS the file name.
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_else(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, s, line);
        chunk.emit_op_u16(Op::LOCAL_GET, cut, line);
        chunk.emit_f64_const(1.0, line);
        chunk.emit_op(Op::F64_ADD, line);
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
    chunk.emit_op(Op::F64_LT, line);
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
    chunk.emit_op(Op::F64_LT, line);
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
    chunk.emit_op(Op::F64_LT, line);
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
    let dot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_GET, name, line);
    chunk.emit_string_const(".", line);
    strings::emit_last_index_of(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, dot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
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
    // Keep extension's existing strict string length/substring behavior for
    // non-string callers, rather than inferring a result solely from a dot.
    emit_extension(chunk, line);
    strings::emit_length(chunk, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_GT, line);
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
    chunk.emit_op(Op::F64_LT, line);
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
    chunk.emit_op(Op::F64_EQ, line);
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
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const("\\", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    chunk.emit_string_const(":", line);
    strings::emit_index_of(chunk, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_EQ, line);
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
fn emit_combine_small(chunk: &mut Chunk, argc: u8, line: u32) {
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
        chunk.emit_op(Op::F64_GT, line);
        chunk.emit_if(line);
        {
            chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
            emit_is_rooted(chunk, line);
            ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
            strings::emit_length(chunk, line);
            chunk.emit_f64_const(0.0, line);
            chunk.emit_op(Op::F64_EQ, line);
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
                chunk.emit_op(Op::F64_ADD, line);
                chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
                strings::emit_length(chunk, line);
                chunk.emit_op(Op::F64_EQ, line);
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

const COMBINE_LONG_LOCAL_BUDGET: u16 = 15;
fn combine_get(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn combine_set(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
fn combine_int(c: &mut Chunk, n: i32, l: u32) {
    c.emit_i32_const(n, l);
}
fn combine_init(c: &mut Chunk, slot: u16, n: i32, l: u32) {
    combine_int(c, n, l);
    combine_set(c, slot, l);
}
fn combine_step(c: &mut Chunk, slot: u16, n: i32, l: u32) {
    combine_get(c, slot, l);
    combine_int(c, n, l);
    c.emit_op(Op::I32_ADD, l);
    combine_set(c, slot, l);
}
fn combine_read(c: &mut Chunk, buf: u16, idx: u16, dst: u16, l: u32) {
    combine_get(c, buf, l);
    combine_get(c, idx, l);
    c.emit_op(Op::ARRAY_GET, l);
    combine_set(c, dst, l);
}
fn combine_write(c: &mut Chunk, buf: u16, idx: u16, src: u16, l: u32) {
    combine_get(c, buf, l);
    combine_get(c, idx, l);
    combine_get(c, src, l);
    c.emit_op(Op::ARRAY_SET, l);
}
fn combine_positive(c: &mut Chunk, seg: u16, l: u32) {
    combine_get(c, seg, l);
    strings::emit_length(c, l);
    c.emit_f64_const(0.0, l);
    c.emit_op(Op::F64_GT, l);
}
fn combine_trailing(c: &mut Chunk, seg: u16, length: Option<u16>, l: u32) {
    combine_get(c, seg, l);
    c.emit_string_const("/", l);
    strings::emit_last_index_of(c, l);
    c.emit_f64_const(1.0, l);
    c.emit_op(Op::F64_ADD, l);
    if let Some(length) = length {
        combine_get(c, length, l);
    } else {
        combine_get(c, seg, l);
        strings::emit_length(c, l);
    }
    c.emit_op(Op::F64_EQ, l);
}
fn combine_finish_loop(c: &mut Chunk, done: usize, again: usize, l: u32) {
    c.emit_br(0, l);
    c.emit_end(l);
    c.patch_loop(again);
    c.emit_end(l);
    c.patch_block(done);
}
pub fn emit_combine(c: &mut Chunk, argc: u8, l: u32) {
    if argc < 16 {
        emit_combine_small(c, argc, l);
        return;
    }
    let before = c.local_count;
    let end = before
        .checked_add(COMBINE_LONG_LOCAL_BUDGET)
        .expect("path local slot overflow");
    let base = c.alloc_scratch(12);
    let buf = base;
    let idx = base + 1;
    let seg = base + 2;
    let total = base + 3;
    let full = base + 4;
    let tail = base + 5;
    let first = base + 6;
    let safe = base + 7;
    let acc = base + 8;
    let count = base + 9;
    let out = base + 10;
    let right = base + 11;
    // Existing GC value-array type zero; no escape or callback route. Keep all
    // evaluated original arguments intact until the fast path is selected.
    c.emit_array_new_fixed(0, argc as u16, l);
    combine_set(c, buf, l);
    combine_init(c, idx, 0, l);
    combine_init(c, tail, 0, l);
    combine_init(c, first, 0, l);
    combine_init(c, safe, 1, l);
    c.emit_f64_const(0.0, l);
    combine_set(c, total, l);
    let done = c.emit_block(l);
    let (again, _) = c.emit_loop_s(l);
    combine_get(c, idx, l);
    combine_int(c, argc as i32, l);
    c.emit_op(Op::I32_GE_S, l);
    c.emit_br_if(1, l);
    combine_read(c, buf, idx, seg, l);
    combine_positive(c, seg, l);
    c.emit_if(l); // Exact strict String gate first.
    combine_get(c, seg, l);
    let length = c.add_import("ecma:string", "length");
    c.emit_call(length, 1, l);
    combine_set(c, full, l);
    combine_init(c, safe, 0, l);
    combine_get(c, full, l);
    c.emit_f64_const(i32::MAX as f64, l);
    c.emit_op(Op::F64_GT, l);
    c.emit_br_if(2, l);
    combine_init(c, safe, 1, l);
    combine_get(c, seg, l);
    emit_is_rooted(c, l);
    ops::emit_dyn_to_bool(c, l);
    combine_get(c, total, l);
    c.emit_f64_const(0.0, l);
    c.emit_op(Op::F64_EQ, l);
    c.emit_op(Op::I32_OR, l);
    c.emit_if_value(l);
    combine_get(c, idx, l);
    combine_set(c, first, l);
    combine_get(c, full, l);
    c.emit_else(l);
    combine_get(c, total, l);
    combine_get(c, tail, l);
    c.emit_if_value(l);
    c.emit_f64_const(0.0, l);
    c.emit_else(l);
    c.emit_f64_const(1.0, l);
    c.emit_end(l);
    c.emit_op(Op::F64_ADD, l);
    combine_get(c, full, l);
    c.emit_op(Op::F64_ADD, l);
    c.emit_end(l);
    combine_set(c, total, l);
    combine_init(c, safe, 0, l);
    combine_get(c, total, l);
    c.emit_f64_const(i32::MAX as f64, l);
    c.emit_op(Op::F64_GT, l);
    c.emit_br_if(2, l);
    combine_init(c, safe, 1, l);
    combine_trailing(c, seg, Some(full), l);
    combine_set(c, tail, l);
    c.emit_end(l);
    combine_step(c, idx, 1, l);
    combine_finish_loop(c, done, again, l);
    combine_get(c, safe, l);
    c.emit_if_value(l);
    // Validation complete: reuse the private buffer slots as String pieces.
    combine_get(c, first, l);
    combine_set(c, idx, l);
    combine_get(c, first, l);
    combine_set(c, out, l);
    combine_init(c, total, 0, l);
    combine_init(c, tail, 0, l);
    let done = c.emit_block(l);
    let (again, _) = c.emit_loop_s(l);
    combine_get(c, idx, l);
    combine_int(c, argc as i32, l);
    c.emit_op(Op::I32_GE_S, l);
    c.emit_br_if(1, l);
    combine_read(c, buf, idx, seg, l);
    combine_positive(c, seg, l);
    c.emit_if(l);
    combine_get(c, total, l);
    c.emit_if(l);
    combine_get(c, tail, l);
    c.emit_if_value(l);
    c.emit_string_const("", l);
    c.emit_else(l);
    c.emit_string_const("/", l);
    c.emit_end(l);
    combine_get(c, seg, l);
    strings::emit_concat(c, 2, l);
    combine_set(c, seg, l);
    c.emit_end(l);
    combine_trailing(c, seg, None, l);
    combine_set(c, tail, l);
    combine_init(c, total, 1, l);
    // Keep ONLY positive pieces. The output cursor cannot overtake the read
    // cursor, and skipping empty strings preserves the single-input Arc.
    combine_write(c, buf, out, seg, l);
    combine_step(c, out, 1, l);
    c.emit_end(l);
    combine_step(c, idx, 1, l);
    combine_finish_loop(c, done, again, l);
    // Merge adjacent pieces in-place, one balanced level per outer iteration.
    combine_get(c, out, l);
    combine_get(c, first, l);
    c.emit_op(Op::I32_SUB, l);
    combine_set(c, count, l);
    let levels = c.emit_block(l);
    let (level_again, _) = c.emit_loop_s(l);
    combine_get(c, count, l);
    combine_int(c, 1, l);
    c.emit_op(Op::I32_LE_S, l);
    c.emit_br_if(1, l);
    combine_get(c, first, l);
    combine_set(c, idx, l);
    combine_get(c, first, l);
    combine_set(c, out, l);
    let pairs = c.emit_block(l);
    let (pair_again, _) = c.emit_loop_s(l);
    combine_get(c, idx, l);
    combine_get(c, first, l);
    combine_get(c, count, l);
    c.emit_op(Op::I32_ADD, l);
    c.emit_op(Op::I32_GE_S, l);
    c.emit_br_if(1, l);
    combine_read(c, buf, idx, acc, l);
    combine_get(c, idx, l);
    combine_int(c, 1, l);
    c.emit_op(Op::I32_ADD, l);
    combine_set(c, right, l);
    combine_get(c, right, l);
    combine_get(c, first, l);
    combine_get(c, count, l);
    c.emit_op(Op::I32_ADD, l);
    c.emit_op(Op::I32_LT_S, l);
    c.emit_if(l);
    combine_get(c, acc, l);
    combine_get(c, buf, l);
    combine_get(c, right, l);
    c.emit_op(Op::ARRAY_GET, l);
    strings::emit_concat(c, 2, l);
    combine_set(c, acc, l);
    c.emit_end(l);
    combine_write(c, buf, out, acc, l);
    combine_step(c, out, 1, l);
    combine_step(c, idx, 2, l);
    combine_finish_loop(c, pairs, pair_again, l);
    combine_get(c, out, l);
    combine_get(c, first, l);
    c.emit_op(Op::I32_SUB, l);
    combine_set(c, count, l);
    combine_finish_loop(c, levels, level_again, l);
    combine_get(c, count, l);
    c.emit_if_value(l);
    combine_get(c, buf, l);
    combine_get(c, first, l);
    c.emit_op(Op::ARRAY_GET, l);
    c.emit_else(l);
    // Same imported empty constant as the frozen accumulator's initial value.
    c.emit_string_const("", l);
    c.emit_end(l);
    c.emit_else(l);
    // Rare signed-length-wrap fallback: original string-operation order over
    // the untouched evaluated argument buffer, now a compact runtime loop.
    combine_init(c, idx, 0, l);
    c.emit_string_const("", l);
    combine_set(c, acc, l);
    let done = c.emit_block(l);
    let (again, _) = c.emit_loop_s(l);
    combine_get(c, idx, l);
    combine_int(c, argc as i32, l);
    c.emit_op(Op::I32_GE_S, l);
    c.emit_br_if(1, l);
    combine_read(c, buf, idx, seg, l);
    combine_positive(c, seg, l);
    c.emit_if(l);
    combine_get(c, seg, l);
    emit_is_rooted(c, l);
    ops::emit_dyn_to_bool(c, l);
    combine_get(c, acc, l);
    strings::emit_length(c, l);
    c.emit_f64_const(0.0, l);
    c.emit_op(Op::F64_EQ, l);
    c.emit_op(Op::I32_OR, l);
    c.emit_if_value(l);
    combine_get(c, seg, l);
    c.emit_else(l);
    combine_get(c, acc, l);
    combine_trailing(c, acc, None, l);
    c.emit_if_value(l);
    c.emit_string_const("", l);
    c.emit_else(l);
    c.emit_string_const("/", l);
    c.emit_end(l);
    combine_get(c, seg, l);
    strings::emit_concat(c, 3, l);
    c.emit_end(l);
    combine_set(c, acc, l);
    c.emit_end(l);
    combine_step(c, idx, 1, l);
    combine_finish_loop(c, done, again, l);
    combine_get(c, acc, l);
    c.emit_end(l);
    assert_eq!(c.local_count, end, "complete compact body budget");
}

// ── The two that are capability reads, not string manipulation ────────────

/// Canonicalize an absolute path: drop `.` and empty segments, resolve `..`
/// against what precedes it, and collapse runs of separators.
///
/// ⛔THIS IS PART OF `GetFullPath`, NOT AN EXTRA. .NET's `GetFullPath` is
/// documented to return a CANONICAL path, and without this step it only made a
/// path absolute — `Convert-Path .` answered `<cwd>/.` and a doubled separator
/// survived as `tests//powershell`.
///
/// Stack: `[path]` → `[path]`.
fn emit_canonicalize(chunk: &mut Chunk, line: u32) {
    let input = chunk.alloc_scratch(5);
    let parts = input + 1;
    let out = input + 2;
    let i = input + 3;
    let seg = input + 4;

    let split = chunk.add_import("ecma:string", "split");
    let arr_new = chunk.add_import("ecma:array", "new");
    let arr_len = chunk.add_import("ecma:array", "length");
    let arr_get = chunk.add_import("ecma:array", "get");
    let arr_push = chunk.add_import("ecma:array", "push");
    let arr_pop = chunk.add_import("ecma:array", "pop");
    let arr_join = chunk.add_import("ecma:array", "join");

    chunk.emit_op_u16(Op::LOCAL_SET, input, line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_string_const(SEP, line);
    chunk.emit_call(split, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, parts, line);
    chunk.emit_call(arr_new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);

    chunk.emit_block(line);
    chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, parts, line);
    chunk.emit_call(arr_len, 1, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, parts, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_call(arr_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, seg, line);

    // `..` pops what precedes it; an empty segment or `.` contributes nothing;
    // anything else is a real component.
    chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
    chunk.emit_string_const("..", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_call(arr_len, 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_call(arr_pop, 1, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
    chunk.emit_string_const("", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
    chunk.emit_string_const(".", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, seg, line);
    chunk.emit_call(arr_push, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // The leading separator was the empty first segment, which the loop drops,
    // so it is put back rather than preserved.
    chunk.emit_string_const(SEP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_string_const(SEP, line);
    chunk.emit_call(arr_join, 2, line);
    strings::emit_concat(chunk, 2, line);

    // ⛔A TRAILING SEPARATOR IS SIGNIFICANT and survives canonicalization —
    // `GetFullPath("tests/")` keeps the slash, marking the path as a directory.
    // The loop drops it with the empty final segment, so it is restored from
    // the INPUT rather than inferred from the result.
    let ends = chunk.add_import("ecma:string", "endsWith");
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_string_const(SEP, line);
    chunk.emit_call(ends, 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_string_const(SEP, line);
    strings::emit_concat(chunk, 2, line);
    chunk.emit_end(line);
}

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

    emit_canonicalize(chunk, line);
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
