//! Path-addressed filesystem verbs, lowered onto `wasi:filesystem@0.3.1`.
//!
//! # Why this module exists
//!
//! `platforms/wasi/src/fs.rs` registers thirty flat, path-taking verbs —
//! `readFile`, `writeFile`, `openFile`, `lineInput`, `pathCombine` — under the
//! REAL `wasi:filesystem` namespace. None of them appear in the WIT. Twenty-six
//! spec names and twenty-six emitted names overlap in exactly one place, and
//! even that one (`stat`) is a bare name where the spec says
//! `[method]descriptor.stat`. So the namespace was squatted: a component built
//! against real WASI cannot satisfy a single one of those imports, and a
//! conforming host cannot serve them.
//!
//! This module is the honest path they migrate ONTO. Every function here
//! composes only names that exist in `proposals/WASI/proposals/filesystem/wit/`.
//!
//! # The shape the spec forces
//!
//! There is no absolute open in WASI, by design — `open-at` resolves relative
//! to a parent descriptor, so every path verb begins by asking
//! `preopens.get-directories()` for a root. Vybe preopens the process cwd, and
//! `PathBuf::join` lets an absolute child replace the base, so both relative and
//! absolute paths resolve through the same first preopen.
//!
//! Reads and writes then go through the 0.3.1 stream pair, NOT the positioned
//! `descriptor.read`/`descriptor.write` — 0.3.1 deleted both. The WIT likens
//! `read-via-stream(offset)` and `write-via-stream(data, offset)` to
//! `pread`/`pwrite`: the offset positions that one transfer and there is no
//! descriptor-wide cursor to disturb.
//!
//! # ⚠Every call here can answer an error instead of the thing you asked for
//!
//! `open-at` on a missing file answers an `error-code`, not a descriptor — and
//! `read-via-stream` on THAT answers another error rather than a
//! `tuple<stream, future>`, so element 0 is not a handle and `canon stream.read`
//! traps with "handle is not a readable stream end". A missing input file is an
//! ordinary condition, not a crash, so every lowering below tests for the error
//! shape before it uses the result. [`emit_ok_test`] is that test, and it is not
//! optional anywhere.

use crate::primitives::class_slots;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::HT_EXTERN;

use crate::primitives::{canon_marshal, collections, io};

/// `wasi:filesystem/types` — the descriptor surface.
const TYPES: &str = "wasi:filesystem/types";
/// `wasi:filesystem/preopens` — where a rooted descriptor comes from.
const PREOPENS: &str = "wasi:filesystem/preopens";

// ── Flag bits ─────────────────────────────────────────────────────────────
//
// `open-flags` in WIT declaration order: create, directory, exclusive,
// truncate. `descriptor-flags`: read, write, file-integrity-sync,
// data-integrity-sync, requested-write-sync, mutate-directory. These mirror
// `platforms/wasi/src/filesystem.rs`, which decodes the same bits.

const OPEN_CREATE: i32 = 1;
const OPEN_DIRECTORY: i32 = 2;
const OPEN_TRUNCATE: i32 = 8;

const DESC_READ: i32 = 1;
const DESC_WRITE: i32 = 2;

/// The property a WASI error answer carries. Absent on every success value,
/// which is what makes [`emit_ok_test`] a total test rather than a guess.
const ERROR_KEY: &str = "__wasi_error";

// ── Building blocks ───────────────────────────────────────────────────────

/// Push the first preopen's descriptor. Stack: `[]` → `[descriptor]`.
///
/// `get-directories()` answers `list<tuple<descriptor, string>>`; element 0 of
/// the list is the root Vybe preopened, and element 0 of that tuple is its
/// descriptor. Taking the first rather than longest-prefix-matching the path is
/// correct for this host specifically: it preopens the cwd, and joining an
/// absolute child onto it replaces the base outright.
fn emit_preopen_root(chunk: &mut Chunk, line: u32) {
    let preopens = chunk.add_import(PREOPENS, "get-directories");
    let at = chunk.add_import("ecma:array", "at");
    chunk.emit_call(preopens, 0, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(at, 2, line); // first tuple<descriptor, string>
    chunk.emit_i32_const(0, line);
    chunk.emit_call(at, 2, line); // its descriptor
}

/// `open-at` on the preopen root, path taken from `path_slot`.
/// Stack: `[]` → `[descriptor | error]`.
fn emit_open_at(chunk: &mut Chunk, line: u32, path_slot: u16, open_flags: i32, desc_flags: i32) {
    let open_at = chunk.add_import(TYPES, "[method]descriptor.open-at");
    emit_preopen_root(chunk, line);
    chunk.emit_i32_const(0, line); // path-flags: no symlink-follow
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_i32_const(open_flags, line);
    chunk.emit_i32_const(desc_flags, line);
    chunk.emit_call(open_at, 5, line);
}

/// `stat-at` on the preopen root. Stack: `[]` → `[descriptor-stat | error]`.
///
/// The cheapest question in the module, and the one every metadata verb is
/// built from: `exists`, `isFile`, `isDir`, `fileSize` and `stat` differ only
/// in which field of the answer they read.
fn emit_stat_at(chunk: &mut Chunk, line: u32, path_slot: u16) {
    let stat_at = chunk.add_import(TYPES, "[method]descriptor.stat-at");
    emit_preopen_root(chunk, line);
    chunk.emit_i32_const(0, line); // path-flags
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_call(stat_at, 3, line);
}

/// Park a WASI answer in `slot` and push whether it SUCCEEDED.
/// Stack: `[value]` → `[i32]`, with `value` left in `slot`.
///
/// The discriminator is structural: `platforms/wasi/src/filesystem.rs::err`
/// builds `{ __wasi_error: <code> }` and every success answer — descriptor,
/// stat record, stream tuple — lacks that key. Reading an absent key yields
/// `Undefined`, which `REF_IS_NULL` reports as null, so "no error property" and
/// "not an object at all" both come out as success, and only a real error
/// answer comes out as failure.
fn emit_ok_test(chunk: &mut Chunk, line: u32, slot: u16) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_string_const(ERROR_KEY, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
}

/// Read a named field out of a WASI record. Stack: `[]` → `[value]`.
fn emit_field(chunk: &mut Chunk, line: u32, slot: u16, field: &str) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_string_const(field, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// Compare a `descriptor-stat`'s `type` against one of the WIT's
/// `descriptor-type` spellings. Stack: `[]` → `[i32]`.
fn emit_type_is(chunk: &mut Chunk, line: u32, stat_slot: u16, want: &str) {
    emit_field(chunk, line, stat_slot, "type");
    chunk.emit_string_const(want, line);
    crate::primitives::strings::emit_str_equals(chunk, line);
}

/// Open for reading and drain the whole file. Stack: `[]` → `[byte array]`,
/// or the empty array when the file cannot be opened or read.
///
/// Factored out because `readFile` and `readFileBytes` differ only in whether a
/// `TextDecoder` pass runs on top, and `copy` needs the bytes with no decode at
/// all. Decoding per chunk would corrupt any multi-byte sequence straddling a
/// buffer boundary, so the decode belongs above the drain, never inside it.
fn emit_drain_file_bytes(chunk: &mut Chunk, line: u32, path_slot: u16) {
    let read_via = chunk.add_import(TYPES, "[method]descriptor.read-via-stream");
    let at = chunk.add_import("ecma:array", "at");
    let desc = chunk.alloc_scratch(1);
    let tuple = chunk.alloc_scratch(1);

    emit_open_at(chunk, line, path_slot, 0, DESC_READ);
    emit_ok_test(chunk, line, desc);
    chunk.emit_if_value(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
        chunk.emit_i32_const(0, line); // offset — read from the start
        chunk.emit_call(read_via, 2, line);
        emit_ok_test(chunk, line, tuple);
        chunk.emit_if_value(line);
        {
            // tuple<stream<u8>, future<result<_, error-code>>>: the readable
            // end is element 0. The future carries a PARTIAL read's failure and
            // is deliberately not consulted — a short read still hands over the
            // bytes it got, which is the whole reason 0.3.1 splits the two.
            chunk.emit_op_u16(Op::LOCAL_GET, tuple, line);
            chunk.emit_i32_const(0, line);
            chunk.emit_call(at, 2, line);
            io::emit_read_stream_to_bytes(chunk, line);
        }
        chunk.emit_else(line);
        canon_marshal::emit_new_bytes(chunk, line);
        chunk.emit_end(line);
    }
    chunk.emit_else(line);
    canon_marshal::emit_new_bytes(chunk, line);
    chunk.emit_end(line);
}

/// Open for writing and push the contents through the stream pair.
/// Stack: `[]` → `[i32]` — 1 on success, 0 if the file could not be opened.
///
/// `append` picks `append-via-stream` (no offset — the position is "the end",
/// fixed by the open mode) over `write-via-stream(data, 0)`.
fn emit_stream_out(
    chunk: &mut Chunk,
    line: u32,
    path_slot: u16,
    data_slot: u16,
    append: bool,
    payload: io::Payload,
) -> bool {
    let method = if append {
        "[method]descriptor.append-via-stream"
    } else {
        "[method]descriptor.write-via-stream"
    };
    let sink = chunk.add_import(TYPES, method);
    let desc = chunk.alloc_scratch(1);
    let rd = chunk.alloc_scratch(1);
    let wr = chunk.alloc_scratch(1);

    // TRUNCATE on a plain write, not on an append: `writeFile` REPLACES.
    // Without it a short write over a long file would leave the old tail
    // behind, which reads as corruption rather than as a failed write.
    let open_flags = if append {
        OPEN_CREATE
    } else {
        OPEN_CREATE | OPEN_TRUNCATE
    };
    emit_open_at(chunk, line, path_slot, open_flags, DESC_WRITE);
    emit_ok_test(chunk, line, desc);
    chunk.emit_if_value(line);
    {
        io::emit_payload_via_stream(
            chunk,
            rd,
            wr,
            line,
            payload,
            |chunk| chunk.emit_op_u16(Op::LOCAL_GET, data_slot, line),
            |chunk, rd| {
                chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
                chunk.emit_op_u16(Op::LOCAL_GET, rd, line);
                if append {
                    chunk.emit_call(sink, 2, line);
                } else {
                    chunk.emit_i32_const(0, line); // offset
                    chunk.emit_call(sink, 3, line);
                }
                // The `future<result<_, error-code>>` is discarded: the host
                // resolves it before returning, and the verbs above report
                // success as a bool that the open already decided.
                chunk.emit_op(Op::DROP, line);
            },
        );
        chunk.emit_bool_const(true, line);
    }
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
    true
}

// ── The verbs ─────────────────────────────────────────────────────────────
//
// Each preserves the RETURN CONTRACT of the invented verb it replaces, so a
// call site migrates by changing which emitter it calls and nothing else. The
// contracts are not all good — `readFile` answering the string `"Error: …"`
// instead of failing is a real defect, and every language's own file API wants
// something different (Node throws `ENOENT`, PHP returns `false`, Python raises
// `FileNotFoundError`). Fixing that is a SEPARATE change with its own
// measurement; folding it in here would make a migration that should be
// byte-neutral impossible to verify.

/// `readFile(path)` → contents as text, or NULL when the file cannot be read.
/// Stack: `[path]` → `[string | null]`.
///
/// ⚠The verb this replaces answered the STRING `"Error: No such file or
/// directory (os error 2)"` — as the file's contents. Every caller that did not
/// inspect it got the error message where the data should be, and no caller
/// could tell an error from a file that happens to start with "Error: ".
/// PHP's `file_get_contents` returns `false`, Python's `open` raises,
/// Node's `readFileSync` throws — not one of them wants that string.
///
/// So failure is `null` and each language maps it to its own idiom. That is a
/// deliberate behaviour change, not a mechanical port: the old contract was
/// invented, and inventing is what this migration exists to stop.
pub fn emit_read_file(chunk: &mut Chunk, line: u32) {
    let path = chunk.alloc_scratch(1);
    let desc = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);

    // The open is tested HERE as well as inside the shared drain, because the
    // drain answers an empty array for both "empty file" and "no such file"
    // and this verb has to tell those apart.
    emit_open_at(chunk, line, path, 0, DESC_READ);
    emit_ok_test(chunk, line, desc);
    chunk.emit_if_value(line);
    {
        emit_drain_file_bytes(chunk, line, path);
        canon_marshal::emit_decode_utf8(chunk, line);
    }
    chunk.emit_else(line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_end(line);
}

/// `readFileBytes(path)` → `[byte array]`, or null when the file is unreadable.
/// Stack: `[path]` → `[array | null]`.
pub fn emit_read_file_bytes(chunk: &mut Chunk, line: u32) {
    let path = chunk.alloc_scratch(1);
    let desc = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);

    emit_open_at(chunk, line, path, 0, DESC_READ);
    emit_ok_test(chunk, line, desc);
    chunk.emit_if_value(line);
    emit_drain_file_bytes(chunk, line, path);
    chunk.emit_else(line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_end(line);
}

/// `writeFile(path, data)` → bool. Stack: `[path, data]` → `[bool]`.
pub fn emit_write_file(chunk: &mut Chunk, line: u32) {
    emit_write_common(chunk, line, false, io::Payload::Text);
}

/// `appendFile(path, data)` → bool. Stack: `[path, data]` → `[bool]`.
pub fn emit_append_file(chunk: &mut Chunk, line: u32) {
    emit_write_common(chunk, line, true, io::Payload::Text);
}

/// `writeFileBytes(path, bytes)` → bool. Stack: `[path, array]` → `[bool]`.
///
/// The mirror of [`emit_read_file_bytes`], and NOT a convenience wrapper over
/// the text form: `io::Payload` exists because a byte array sent down the text
/// path is `TextEncoder`'d, so the file receives the array's DECIMAL RENDERING
/// — `"72,101,108"` — with a plausible length and a successful return.
///
/// The callers that need this are the ones whose language has a byte-oriented
/// write at all: `File.WriteAllBytes` in .NET, which until now round-tripped
/// its bytes through `String.fromCharCode` and a UTF-8 encode, so every byte
/// above 0x7F was written as two.
pub fn emit_write_file_bytes(chunk: &mut Chunk, line: u32) {
    emit_write_common(chunk, line, false, io::Payload::Bytes);
}

fn emit_write_common(chunk: &mut Chunk, line: u32, append: bool, payload: io::Payload) {
    let data = chunk.alloc_scratch(1);
    let path = chunk.alloc_scratch(1);
    // Arguments land on the stack in call order, so the LAST one pops FIRST.
    chunk.emit_op_u16(Op::LOCAL_SET, data, line);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    emit_stream_out(chunk, line, path, data, append, payload);
}

/// `exists(path)` → bool. Stack: `[path]` → `[bool]`.
///
/// The smallest end-to-end proof of this whole module: one `stat-at`, no
/// streams, no decode. If this does not round-trip, nothing above it will.
pub fn emit_exists(chunk: &mut Chunk, line: u32) {
    let path = chunk.alloc_scratch(1);
    let st = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    emit_stat_at(chunk, line, path);
    emit_ok_test(chunk, line, st);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `isFile(path)` → bool. Stack: `[path]` → `[bool]`.
pub fn emit_is_file(chunk: &mut Chunk, line: u32) {
    emit_type_probe(chunk, line, "regular-file");
}

/// `isDir(path)` → bool. Stack: `[path]` → `[bool]`.
pub fn emit_is_dir(chunk: &mut Chunk, line: u32) {
    emit_type_probe(chunk, line, "directory");
}

fn emit_type_probe(chunk: &mut Chunk, line: u32, want: &str) {
    let path = chunk.alloc_scratch(1);
    let st = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    emit_stat_at(chunk, line, path);
    emit_ok_test(chunk, line, st);
    chunk.emit_if_value(line);
    emit_type_is(chunk, line, st, want);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

/// `fileSize(path)` → f64 byte count, `-1` when unknown.
/// Stack: `[path]` → `[f64]`.
pub fn emit_file_size(chunk: &mut Chunk, line: u32) {
    let path = chunk.alloc_scratch(1);
    let st = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    emit_stat_at(chunk, line, path);
    emit_ok_test(chunk, line, st);
    chunk.emit_if_value(line);
    emit_field(chunk, line, st, "size");
    chunk.emit_else(line);
    chunk.emit_f64_const(-1.0, line);
    chunk.emit_end(line);
}

/// `stat(path)` → `{ size, isFile, isDir, modified }`, or null.
/// Stack: `[path]` → `[object | null]`.
///
/// Deliberately re-shaped rather than passed through: the spec answer is a
/// `descriptor-stat` with `type` / `link-count` / `size` / three timestamps,
/// and the callers of this verb read `isFile` and `modified`. Translating here
/// keeps the WIT record honest at the boundary AND leaves the call sites
/// untouched.
pub fn emit_stat(chunk: &mut Chunk, line: u32) {
    let path = chunk.alloc_scratch(1);
    let st = chunk.alloc_scratch(1);
    let out = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    emit_stat_at(chunk, line, path);
    emit_ok_test(chunk, line, st);
    chunk.emit_if_value(line);
    {
        // ⚠A STRUCT, not a Map. The verb this replaces answered a host
        // `Object` whose properties the consumers read with `Op::STRUCT_GET`,
        // and `ecma:map.new` + `ecma:array.set` writes an ENTRY rather than a
        // property — a different store that `STRUCT_GET` cannot see. Building
        // the wrong one is silent: `st_size` read back as `Undefined` and only
        // surfaced two frames later as
        // `wasm:js-number.toF64 — not a number`.
        class_slots::emit_class_alloc(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, out, line);

        for (key, emit) in [
            ("size", StatField::Field("size")),
            ("isFile", StatField::TypeIs("regular-file")),
            ("isDir", StatField::TypeIs("directory")),
            ("modified", StatField::Field("data-modification-timestamp")),
        ] {
            chunk.emit_op_u16(Op::LOCAL_GET, out, line);
            match emit {
                StatField::Field(name) => emit_field(chunk, line, st, name),
                StatField::TypeIs(want) => emit_type_is(chunk, line, st, want),
            }
            let k = class_slots::resolve(
                &class_slots::ClassSlot::internal(key),
                &class_slots::PlainNames,
            );
            class_slots::emit_class_set(
                chunk,
                class_slots::ObjSource::Stack,
                &k,
                class_slots::ValueSource::Stack,
                line,
            );
        }
        chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    }
    chunk.emit_else(line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_end(line);
}

/// How one output field of [`emit_stat`] is derived from the WIT record.
enum StatField {
    /// Copied straight across under a different name.
    Field(&'static str),
    /// Derived by comparing `descriptor-stat.type`.
    TypeIs(&'static str),
}

/// `mkdir(path)` → bool. Stack: `[path]` → `[bool]`.
///
/// ⚠NOT recursive, unlike the `create_dir_all` the invented verb used.
/// `create-directory-at` is one level, matching POSIX `mkdirat`; WASI has no
/// recursive form because each level is a separate capability check. Callers
/// that need `mkdir -p` have to walk the path themselves.
pub fn emit_mkdir(chunk: &mut Chunk, line: u32) {
    let mkdir = chunk.add_import(TYPES, "[method]descriptor.create-directory-at");
    let path = chunk.alloc_scratch(1);
    let res = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    emit_preopen_root(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path, line);
    chunk.emit_call(mkdir, 2, line);
    emit_ok_test(chunk, line, res);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// The recursive form — `mkdir -p`, `os.makedirs`, `create_dir_all`.
/// Stack: `[path]` → `[bool]`.
///
/// WASI has no recursive create and cannot have one: every level is a separate
/// capability check, which is the whole point of `-at` addressing. So the walk
/// is the GUEST's job, and this is it — split on `/`, create each prefix in
/// turn, ignore the `exist` failures that intermediate levels legitimately
/// produce, and answer by asking whether the leaf is now a directory.
///
/// ⚠This is NOT the same verb as [`emit_mkdir`], and conflating them is a live
/// bug this migration inherits: `languages/python/src/profile` binds BOTH
/// `os.mkdir` and `os.makedirs` to the recursive shim, so
/// `os.mkdir("a/b/c")` with no `a/b` silently builds the tree where CPython
/// raises `FileNotFoundError`. Splitting the two here is what lets the
/// non-recursive one start telling the truth.
pub fn emit_mkdir_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let line_ = line;
    let chunk = &mut chunks[current];
    let mkdir = chunk.add_import(TYPES, "[method]descriptor.create-directory-at");
    let split = chunk.add_import("ecma:string", "split");
    let len = chunk.add_import("ecma:array", "length");
    let to_i32 = chunk.add_import("wasm:js-number", "toI32");

    let path = chunk.alloc_scratch(1);
    let parts = chunk.alloc_scratch(1);
    let n = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);
    let acc = chunk.alloc_scratch(1);
    let part = chunk.alloc_scratch(1);
    let st = chunk.alloc_scratch(1);
    let sink = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, path, line_);

    chunk.emit_op_u16(Op::LOCAL_GET, path, line_);
    chunk.emit_string_const("/", line_);
    chunk.emit_call(split, 2, line_);
    chunk.emit_op_u16(Op::LOCAL_SET, parts, line_);

    chunk.emit_op_u16(Op::LOCAL_GET, parts, line_);
    chunk.emit_call(len, 1, line_);
    chunk.emit_call(to_i32, 1, line_);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line_);

    chunk.emit_i32_const(0, line_);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line_);
    chunk.emit_string_const("", line_);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line_);

    let done = chunk.emit_block(line_);
    let (walk, _) = chunk.emit_loop_s(line_);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line_);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line_);
    chunk.emit_op(Op::I32_GE_S, line_);
    chunk.emit_br_if(1, line_);

    chunk.emit_op_u16(Op::LOCAL_GET, parts, line_);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line_);
    chunk.emit_op(Op::ARRAY_GET, line_);
    chunk.emit_op_u16(Op::LOCAL_SET, part, line_);

    // acc grows by "/" + part, except for the very first segment which
    // contributes no separator. An absolute path's leading segment is EMPTY,
    // so `"" + "/" + "tmp"` reproduces the leading slash for free — and the
    // empty prefix it would otherwise try to create is skipped below.
    chunk.emit_op_u16(Op::LOCAL_GET, i, line_);
    chunk.emit_i32_const(0, line_);
    chunk.emit_op(Op::I32_EQ, line_);
    chunk.emit_if_value(line_);
    chunk.emit_op_u16(Op::LOCAL_GET, part, line_);
    chunk.emit_else(line_);
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line_);
    chunk.emit_string_const("/", line_);
    chunk.emit_op_u16(Op::LOCAL_GET, part, line_);
    crate::primitives::strings::emit_concat(chunk, 3, line_);
    chunk.emit_end(line_);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line_);

    // Skip empty prefixes: a leading "/" and any "//" produce one, and asking
    // to create "" is not a level at all.
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line_);
    chunk.emit_string_const("", line_);
    crate::primitives::strings::emit_str_equals(chunk, line_);
    chunk.emit_op(Op::I32_EQZ, line_);
    chunk.emit_if(line_);
    emit_preopen_root(chunk, line_);
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line_);
    chunk.emit_call(mkdir, 2, line_);
    // An intermediate level that already exists answers `exist`, which is the
    // ordinary case on the second call and NOT a failure of the whole walk.
    // The leaf's own outcome is decided by the `stat-at` below instead.
    chunk.emit_op_u16(Op::LOCAL_SET, sink, line_);
    chunk.emit_end(line_);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line_);
    chunk.emit_i32_const(1, line_);
    chunk.emit_op(Op::I32_ADD, line_);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line_);
    chunk.emit_br(0, line_);
    chunk.emit_end(line_);
    chunk.patch_loop(walk);
    chunk.emit_end(line_);
    chunk.patch_block(done);

    // Success is "the leaf is a directory now", not "no call errored" —
    // `create_dir_all` has the same contract and it is the only one that stays
    // true when part of the tree already existed.
    emit_stat_at(chunk, line_, path);
    emit_ok_test(chunk, line_, st);
    chunk.emit_if_value(line_);
    emit_type_is(chunk, line_, st, "directory");
    crate::primitives::ops::emit_i32_to_bool(chunk, line_);
    chunk.emit_else(line_);
    chunk.emit_bool_const(false, line_);
    chunk.emit_end(line_);
}

/// `remove(path)` → bool. Stack: `[path]` → `[bool]`.
///
/// WASI splits removal by kind — `unlink-file-at` and `remove-directory-at` are
/// separate calls, as in POSIX — so this asks `stat-at` first. The old verb
/// used `remove_dir_all` for directories; `remove-directory-at` answers
/// `not-empty` instead, so a non-empty directory now reports failure rather
/// than recursively deleting. That is the spec's behaviour and the safer of
/// the two.
pub fn emit_remove(chunk: &mut Chunk, line: u32) {
    let unlink = chunk.add_import(TYPES, "[method]descriptor.unlink-file-at");
    let rmdir = chunk.add_import(TYPES, "[method]descriptor.remove-directory-at");
    let path = chunk.alloc_scratch(1);
    let st = chunk.alloc_scratch(1);
    let res = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);

    emit_stat_at(chunk, line, path);
    emit_ok_test(chunk, line, st);
    chunk.emit_if_value(line);
    {
        emit_type_is(chunk, line, st, "directory");
        chunk.emit_if_value(line);
        {
            emit_preopen_root(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_GET, path, line);
            chunk.emit_call(rmdir, 2, line);
        }
        chunk.emit_else(line);
        {
            emit_preopen_root(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_GET, path, line);
            chunk.emit_call(unlink, 2, line);
        }
        chunk.emit_end(line);
        emit_ok_test(chunk, line, res);
        crate::primitives::ops::emit_i32_to_bool(chunk, line);
    }
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

/// `unlink-file-at(path)` → bool. Stack: `[path]` → `[bool]`.
///
/// File-only removal, which is what POSIX `unlink` and PHP's `unlink()` both
/// mean — a directory argument FAILS rather than being removed. [`emit_remove`]
/// is the permissive form that dispatches on the kind; a language whose verb
/// draws the file/directory line itself wants this one, so that the line keeps
/// being drawn.
pub fn emit_unlink(chunk: &mut Chunk, line: u32) {
    emit_removal(chunk, line, "[method]descriptor.unlink-file-at");
}

/// `remove-directory-at(path)` → bool. Stack: `[path]` → `[bool]`.
///
/// Directory-only, and NON-recursive: a non-empty directory answers
/// `not-empty`, exactly as POSIX `rmdir` and PHP's `rmdir()` do.
pub fn emit_rmdir(chunk: &mut Chunk, line: u32) {
    emit_removal(chunk, line, "[method]descriptor.remove-directory-at");
}

fn emit_removal(chunk: &mut Chunk, line: u32, method: &str) {
    let remove = chunk.add_import(TYPES, method);
    let path = chunk.alloc_scratch(1);
    let res = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    emit_preopen_root(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path, line);
    chunk.emit_call(remove, 2, line);
    emit_ok_test(chunk, line, res);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `rmtree(path)` — remove a directory AND everything under it.
/// Stack: `[path]` → `[bool]`.
///
/// WASI has no recursive delete, for the same reason it has no `mkdir -p`:
/// every level is a separate capability check. So the walk is the guest's, and
/// this is it — breadth-first to collect directories, unlinking files as it
/// goes, then removing the directories in REVERSE discovery order so children
/// always precede their parents. `remove-directory-at` answers `not-empty`
/// otherwise, and the whole tree would survive.
///
/// This is the verb that could not be written until `read-directory` had a
/// guest-side reader: `shutil.rmtree` and `Directory.Delete(recursive)` are a
/// LOOP over a listing, so a filesystem that cannot enumerate cannot delete a
/// tree. It was the single binding left on the retired shim while that was
/// true, and the single measured regression from retiring it.
pub fn emit_remove_all(chunk: &mut Chunk, line: u32) {
    let push = chunk.add_import("ecma:array", "push");
    let new_arr = chunk.add_import("vybe:js-array", "newWithLength");
    let at = chunk.add_import("ecma:array", "at");
    let length = chunk.add_import("ecma:array", "length");
    let to_i32 = chunk.add_import("wasm:js-number", "toI32");

    let root = chunk.alloc_scratch(1);
    let dirs = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);
    let d = chunk.alloc_scratch(1);
    let entries = chunk.alloc_scratch(1);
    let j = chunk.alloc_scratch(1);
    let n = chunk.alloc_scratch(1);
    let e = chunk.alloc_scratch(1);
    let child = chunk.alloc_scratch(1);
    let st = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, root, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_call(new_arr, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dirs, line);
    chunk.emit_op_u16(Op::LOCAL_GET, dirs, line);
    chunk.emit_op_u16(Op::LOCAL_GET, root, line);
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);

    // ── pass 1: walk, unlinking files and collecting directories ───────────
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    let walk_done = chunk.emit_block(line);
    let (walk, _) = chunk.emit_loop_s(line);
    {
        // The bound is re-read each turn ON PURPOSE: the loop appends to the
        // very array it is scanning, which is how the breadth-first walk
        // reaches nested directories without recursion.
        chunk.emit_op_u16(Op::LOCAL_GET, dirs, line);
        chunk.emit_call(length, 1, line);
        chunk.emit_call(to_i32, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, n, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_op_u16(Op::LOCAL_GET, n, line);
        chunk.emit_op(Op::I32_GE_S, line);
        chunk.emit_br_if(1, line);

        chunk.emit_op_u16(Op::LOCAL_GET, dirs, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_call(at, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, d, line);

        chunk.emit_op_u16(Op::LOCAL_GET, d, line);
        emit_read_directory_entries(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_SET, entries, line);

        chunk.emit_i32_const(0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, j, line);
        let inner_done = chunk.emit_block(line);
        let (inner, _) = chunk.emit_loop_s(line);
        {
            chunk.emit_op_u16(Op::LOCAL_GET, j, line);
            chunk.emit_op_u16(Op::LOCAL_GET, entries, line);
            chunk.emit_call(length, 1, line);
            chunk.emit_call(to_i32, 1, line);
            chunk.emit_op(Op::I32_GE_S, line);
            chunk.emit_br_if(1, line);

            chunk.emit_op_u16(Op::LOCAL_GET, entries, line);
            chunk.emit_op_u16(Op::LOCAL_GET, j, line);
            chunk.emit_call(at, 2, line);
            chunk.emit_op_u16(Op::LOCAL_SET, e, line);

            chunk.emit_op_u16(Op::LOCAL_GET, d, line);
            chunk.emit_string_const("/", line);
            emit_field(chunk, line, e, "name");
            crate::primitives::strings::emit_concat(chunk, 3, line);
            chunk.emit_op_u16(Op::LOCAL_SET, child, line);

            // `type` is the WIT's `descriptor-type` case name, which is why the
            // comparison is against the spec spelling rather than a bool.
            emit_field(chunk, line, e, "type");
            chunk.emit_string_const("directory", line);
            crate::primitives::strings::emit_str_equals(chunk, line);
            chunk.emit_if(line);
            {
                chunk.emit_op_u16(Op::LOCAL_GET, dirs, line);
                chunk.emit_op_u16(Op::LOCAL_GET, child, line);
                chunk.emit_call(push, 2, line);
                chunk.emit_op(Op::DROP, line);
            }
            chunk.emit_else(line);
            {
                chunk.emit_op_u16(Op::LOCAL_GET, child, line);
                emit_unlink(chunk, line);
                chunk.emit_op(Op::DROP, line);
            }
            chunk.emit_end(line);

            chunk.emit_op_u16(Op::LOCAL_GET, j, line);
            chunk.emit_i32_const(1, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, j, line);
            chunk.emit_br(0, line);
        }
        chunk.emit_end(line);
        chunk.patch_loop(inner);
        chunk.emit_end(line);
        chunk.patch_block(inner_done);

        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, i, line);
        chunk.emit_br(0, line);
    }
    chunk.emit_end(line);
    chunk.patch_loop(walk);
    chunk.emit_end(line);
    chunk.patch_block(walk_done);

    // ── pass 2: remove directories deepest-first ───────────────────────────
    chunk.emit_op_u16(Op::LOCAL_GET, dirs, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_call(to_i32, 1, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    let rm_done = chunk.emit_block(line);
    let (rm, _) = chunk.emit_loop_s(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_LT_S, line);
        chunk.emit_br_if(1, line);

        chunk.emit_op_u16(Op::LOCAL_GET, dirs, line);
        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_call(at, 2, line);
        emit_rmdir(chunk, line);
        chunk.emit_op(Op::DROP, line);

        chunk.emit_op_u16(Op::LOCAL_GET, i, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_SUB, line);
        chunk.emit_op_u16(Op::LOCAL_SET, i, line);
        chunk.emit_br(0, line);
    }
    chunk.emit_end(line);
    chunk.patch_loop(rm);
    chunk.emit_end(line);
    chunk.patch_block(rm_done);

    // Success is "the root is gone", not "no call errored" — the same contract
    // `remove_dir_all` has, and the only one that survives a partially-shared
    // tree.
    chunk.emit_op_u16(Op::LOCAL_GET, root, line);
    emit_stat_at_stack(chunk, line);
    emit_ok_test(chunk, line, st);
    chunk.emit_op(Op::I32_EQZ, line);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `stat-at` with the path taken from the STACK rather than a slot.
fn emit_stat_at_stack(chunk: &mut Chunk, line: u32) {
    let p = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, p, line);
    emit_stat_at(chunk, line, p);
}

/// `rename(old, new)` → bool. Stack: `[old, new]` → `[bool]`.
///
/// `rename-at(this, old-path, new-descriptor, new-path)` — the destination
/// carries its OWN descriptor, because a rename may cross preopens. Both ends
/// are the same root here.
pub fn emit_rename(chunk: &mut Chunk, line: u32) {
    let rename = chunk.add_import(TYPES, "[method]descriptor.rename-at");
    let new = chunk.alloc_scratch(1);
    let old = chunk.alloc_scratch(1);
    let res = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, new, line);
    chunk.emit_op_u16(Op::LOCAL_SET, old, line);

    emit_preopen_root(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, old, line);
    emit_preopen_root(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, new, line);
    chunk.emit_call(rename, 4, line);
    emit_ok_test(chunk, line, res);
    crate::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `copy(src, dest)` → bool. Stack: `[src, dest]` → `[bool]`.
///
/// WASI has no copy — POSIX has none either — so this is a read followed by a
/// write, which is what `std::fs::copy` does underneath. Byte-exact: the
/// contents never become text, so no encoding can round-trip wrongly.
pub fn emit_copy(chunk: &mut Chunk, line: u32) {
    let dest = chunk.alloc_scratch(1);
    let src = chunk.alloc_scratch(1);
    let bytes = chunk.alloc_scratch(1);
    let desc = chunk.alloc_scratch(1);
    let src_desc = chunk.alloc_scratch(1);
    let rd = chunk.alloc_scratch(1);
    let wr = chunk.alloc_scratch(1);
    let write_via = chunk.add_import(TYPES, "[method]descriptor.write-via-stream");
    chunk.emit_op_u16(Op::LOCAL_SET, dest, line);
    chunk.emit_op_u16(Op::LOCAL_SET, src, line);

    // Report failure on a missing SOURCE rather than silently creating an
    // empty destination — `emit_drain_file_bytes` answers an empty array for
    // both "empty file" and "no such file", so the open has to be tested here
    // to tell them apart.
    emit_open_at(chunk, line, src, 0, DESC_READ);
    emit_ok_test(chunk, line, src_desc);
    chunk.emit_if_value(line);
    {
        emit_drain_file_bytes(chunk, line, src);
        chunk.emit_op_u16(Op::LOCAL_SET, bytes, line);

        emit_open_at(chunk, line, dest, OPEN_CREATE | OPEN_TRUNCATE, DESC_WRITE);
        emit_ok_test(chunk, line, desc);
        chunk.emit_if_value(line);
        {
            io::emit_payload_via_stream(
                chunk,
                rd,
                wr,
                line,
                // BYTES, not text: the drain answers an array and a copied file
                // has no business round-tripping through UTF-8.
                io::Payload::Bytes,
                |chunk| chunk.emit_op_u16(Op::LOCAL_GET, bytes, line),
                |chunk, rd| {
                    chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
                    chunk.emit_op_u16(Op::LOCAL_GET, rd, line);
                    chunk.emit_i32_const(0, line);
                    chunk.emit_call(write_via, 3, line);
                    chunk.emit_op(Op::DROP, line);
                },
            );
            chunk.emit_bool_const(true, line);
        }
        chunk.emit_else(line);
        chunk.emit_bool_const(false, line);
        chunk.emit_end(line);
    }
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

// ── Numbered file handles ────────────────────────────────────────────────
//
// `Open #1 For Output As #1` / `Print #1, x` / `Line Input #1, s` /
// `AssignFile(f, path)` — VB, Pascal and COBOL all address an open file by a
// NUMBER rather than by a descriptor. That is language syntax, and it stays.
// What does not stay is the six host verbs it used to lower to — `openFile`,
// `closeFile`, `printFile`, `writeFile_handle`, `lineInput`, `inputFile` —
// none of which are in any WIT.
//
// So the handle table moves INTO THE GUEST. A file number maps to a record of
// `{path, mode, pos}` in a guest global, and every transfer is `open-at` plus
// the 0.3.1 via-stream pair at an explicit offset. Nothing below calls
// anything the spec does not declare.
//
// ⚠`pos` is a BYTE offset, because `read-via-stream`/`write-via-stream` take
// `filesize` and the WIT likens them to `pread`/`pwrite`. It is NOT a character
// index: advancing it by a string's `.length` would desynchronise the moment a
// file contained one non-ASCII character, and would do so silently — the reads
// would keep succeeding, just from the wrong place.

/// Guest global holding the open-file table: a Map from file number to a
/// `{path, mode, pos}` record. Created on first `open`.
const HANDLES: &str = "__vybe_file_handles";

const H_PATH: &str = "path";
const H_MODE: &str = "mode";
const H_POS: &str = "pos";

/// Pop `argc` call arguments into consecutive scratch locals, leftmost first.
///
/// Arguments arrive on the stack in call order, so the LAST one pops FIRST;
/// this reverses that back into positional slots. Needed by every variadic verb
/// here — `Print #1, a, b, c` has its file number buried under its data.
fn stash(chunk: &mut Chunk, argc: u8, line: u32) -> u16 {
    let base = chunk.alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// Push the handle table, creating it on first use. Stack: `[]` → `[map]`.
fn emit_handle_table(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::primitives::globals::emit_read(&mut chunks[current], HANDLES, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    collections::emit_map_new(chunks, current, line);
    crate::primitives::globals::emit_write(&mut chunks[current], HANDLES, line);
    chunks[current].emit_end(line);
    crate::primitives::globals::emit_read(&mut chunks[current], HANDLES, line);
}

/// Push `handles[fnum]`. Stack: `[]` → `[record | undefined]`.
fn emit_handle_of(chunks: &mut [Chunk], current: usize, fnum: u16, line: u32) {
    emit_handle_table(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fnum, line);
    collections::emit_get(chunks, current, line);
}

/// Read one field of a handle record. Stack: `[]` → `[value]`.
fn emit_handle_field(chunks: &mut [Chunk], current: usize, h: u16, field: &str, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, h, line);
    chunks[current].emit_string_const(field, line);
    collections::emit_get(chunks, current, line);
}

/// Write one field of a handle record, consuming a value from the stack.
fn emit_set_handle_field(chunks: &mut [Chunk], current: usize, h: u16, field: &str, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, h, line);
    chunks[current].emit_string_const(field, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `Open path For <mode> As #n`. Stack: `[path, mode, fnum]` → `[null]`.
///
/// The mode decides two things and only two: whether the file is truncated at
/// open, and where `pos` starts.
///   - `Output` truncates and starts at 0 — a fresh file.
///   - `Append` creates if absent and starts at the CURRENT SIZE, so the first
///     write lands past the existing contents.
///   - `Input` neither creates nor truncates and starts at 0.
///
/// `Append`'s start position is fixed HERE rather than at the first write,
/// because the position is a property of the open: re-deriving it per transfer
/// would let an intervening write by anyone else move it.
pub fn emit_open_file(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(&mut chunks[current], argc, line);
    let (path, mode, fnum) = (base, base + 1, base + 2);
    let entry = chunks[current].alloc_scratch(1);
    let st = chunks[current].alloc_scratch(1);
    let desc = chunks[current].alloc_scratch(1);

    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entry, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    emit_set_handle_field(chunks, current, entry, H_PATH, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mode, line);
    emit_set_handle_field(chunks, current, entry, H_MODE, line);

    // Output: truncate now, so an opened-but-never-written file is empty
    // rather than carrying the previous run's contents.
    chunks[current].emit_op_u16(Op::LOCAL_GET, mode, line);
    chunks[current].emit_string_const("Output", line);
    crate::primitives::strings::emit_str_equals(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_open_at(
        &mut chunks[current],
        line,
        path,
        OPEN_CREATE | OPEN_TRUNCATE,
        DESC_WRITE,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, desc, line);
    chunks[current].emit_end(line);

    // Append: create if absent, and start at the end.
    chunks[current].emit_op_u16(Op::LOCAL_GET, mode, line);
    chunks[current].emit_string_const("Append", line);
    crate::primitives::strings::emit_str_equals(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    {
        emit_open_at(&mut chunks[current], line, path, OPEN_CREATE, DESC_WRITE);
        chunks[current].emit_op_u16(Op::LOCAL_SET, desc, line);
        emit_stat_at(&mut chunks[current], line, path);
        emit_ok_test(&mut chunks[current], line, st);
        chunks[current].emit_if_value(line);
        emit_field(&mut chunks[current], line, st, "size");
        chunks[current].emit_else(line);
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_end(line);
    emit_set_handle_field(chunks, current, entry, H_POS, line);

    emit_handle_table(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fnum, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// `Close #n`, or `Close` with `-1` for every open file.
/// Stack: `[fnum]` → `[null]`.
///
/// There is no descriptor to release: every transfer opens, moves its bytes and
/// lets the descriptor go, because `read-via-stream`/`write-via-stream` are
/// positioned (`pread`/`pwrite`) and carry no cursor of their own. So closing
/// is forgetting the number.
pub fn emit_close_file(chunks: &mut [Chunk], current: usize, line: u32) {
    let fnum = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fnum, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fnum, line);
    chunks[current].emit_f64_const(-1.0, line);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    // `Close` with no argument: a fresh table IS closing every file, since a
    // handle holds no descriptor to release.
    collections::emit_map_new(chunks, current, line);
    crate::primitives::globals::emit_write(&mut chunks[current], HANDLES, line);
    chunks[current].emit_else(line);
    // Nulling the entry rather than removing the key: every reader here already
    // treats a null handle as closed, and a null is one `emit_set` where a real
    // delete would need a whole map-rebuild helper.
    emit_handle_table(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fnum, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// Append `text` (already on the stack) to the file behind `fnum`, advancing
/// its recorded position by the text's UTF-8 BYTE length.
fn emit_append_at_pos(chunks: &mut [Chunk], current: usize, fnum: u16, line: u32) {
    let text = chunks[current].alloc_scratch(1);
    let h = chunks[current].alloc_scratch(1);
    let pos = chunks[current].alloc_scratch(1);
    let desc = chunks[current].alloc_scratch(1);
    let rd = chunks[current].alloc_scratch(1);
    let wr = chunks[current].alloc_scratch(1);
    let write_via = chunks[current].add_import(TYPES, "[method]descriptor.write-via-stream");

    chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);
    emit_handle_of(chunks, current, fnum, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, h, line);

    // An unopened file number is a no-op rather than a trap: `Print #9` with no
    // `Open #9` is a program bug, but crashing the VM is not this layer's call.
    chunks[current].emit_op_u16(Op::LOCAL_GET, h, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    {
        emit_handle_field(chunks, current, h, H_POS, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, pos, line);

        emit_handle_field(chunks, current, h, H_PATH, line);
        let path = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
        emit_open_at(&mut chunks[current], line, path, OPEN_CREATE, DESC_WRITE);
        emit_ok_test(&mut chunks[current], line, desc);
        chunks[current].emit_if(line);
        {
            io::emit_write_via_stream(
                &mut chunks[current],
                rd,
                wr,
                line,
                |chunk| chunk.emit_op_u16(Op::LOCAL_GET, text, line),
                |chunk, rd| {
                    chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
                    chunk.emit_op_u16(Op::LOCAL_GET, rd, line);
                    chunk.emit_op_u16(Op::LOCAL_GET, pos, line);
                    chunk.emit_call(write_via, 3, line);
                    chunk.emit_op(Op::DROP, line);
                },
            );
            // pos += UTF-8 byte length. NOT `.length`: that counts UTF-16 code
            // units, so one non-ASCII character would leave every subsequent
            // write overlapping the previous one.
            chunks[current].emit_op_u16(Op::LOCAL_GET, pos, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
            crate::primitives::strings::emit_byte_length(chunks, current, line);
            crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
            emit_set_handle_field(chunks, current, h, H_POS, line);
        }
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
}

/// `Print #n, a, b, c` — the parts concatenated, then one newline.
/// Stack: `[fnum, item...]` → `[null]`.
pub fn emit_print_file(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(&mut chunks[current], argc, line);
    let fnum = base;
    let items = argc.saturating_sub(1) as usize;
    for i in 0..items {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1 + i as u16, line);
        crate::primitives::strings::emit_to_string(&mut chunks[current], line);
    }
    chunks[current].emit_string_const("\n", line);
    crate::primitives::strings::emit_concat(&mut chunks[current], items + 1, line);
    emit_append_at_pos(chunks, current, fnum, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// `Write #n, a, b, c` — CSV: each part quoted if it contains a comma or a
/// quote (quotes doubled), joined by commas, then one newline.
/// Stack: `[fnum, item...]` → `[null]`.
pub fn emit_write_file_handle(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash(&mut chunks[current], argc, line);
    let fnum = base;
    let items = argc.saturating_sub(1) as usize;
    let mut parts = 0usize;
    for i in 0..items {
        if i > 0 {
            chunks[current].emit_string_const(",", line);
            parts += 1;
        }
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + 1 + i as u16, line);
        crate::primitives::strings::emit_to_string(&mut chunks[current], line);
        emit_csv_quote(chunks, current, line);
        parts += 1;
    }
    chunks[current].emit_string_const("\n", line);
    parts += 1;
    crate::primitives::strings::emit_concat(&mut chunks[current], parts, line);
    emit_append_at_pos(chunks, current, fnum, line);
    chunks[current].emit_ref_null(HT_EXTERN, line);
}

/// Quote one CSV field if it needs it. Stack: `[string]` → `[string]`.
fn emit_csv_quote(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);

    // Needs quoting iff it contains a comma or a quote.
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_string_const(",", line);
    crate::primitives::strings::emit_index_of(&mut chunks[current], line);
    chunks[current].emit_f64_const(0.0, line);
    crate::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_string_const("\"", line);
    crate::primitives::strings::emit_index_of(&mut chunks[current], line);
    chunks[current].emit_f64_const(0.0, line);
    crate::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    {
        chunks[current].emit_string_const("\"", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
        chunks[current].emit_string_const("\"", line);
        chunks[current].emit_string_const("\"\"", line);
        crate::primitives::strings::emit_replace(&mut chunks[current], line);
        chunks[current].emit_string_const("\"", line);
        crate::primitives::strings::emit_concat(&mut chunks[current], 3, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    chunks[current].emit_end(line);
}

/// Read the file behind `fnum` from its recorded position to the end.
/// Stack: `[]` → `[string]`.
fn emit_rest_of_file(chunks: &mut [Chunk], current: usize, h: u16, line: u32) {
    let path = chunks[current].alloc_scratch(1);
    let pos = chunks[current].alloc_scratch(1);
    let desc = chunks[current].alloc_scratch(1);
    let tuple = chunks[current].alloc_scratch(1);
    let read_via = chunks[current].add_import(TYPES, "[method]descriptor.read-via-stream");
    let at = chunks[current].add_import("ecma:array", "at");

    emit_handle_field(chunks, current, h, H_PATH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    emit_handle_field(chunks, current, h, H_POS, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pos, line);

    emit_open_at(&mut chunks[current], line, path, 0, DESC_READ);
    emit_ok_test(&mut chunks[current], line, desc);
    chunks[current].emit_if_value(line);
    {
        // `read-via-stream(offset)` is positioned, so the recorded byte offset
        // IS the seek — there is no cursor to advance and nothing to rewind.
        chunks[current].emit_op_u16(Op::LOCAL_GET, desc, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, pos, line);
        chunks[current].emit_call(read_via, 2, line);
        emit_ok_test(&mut chunks[current], line, tuple);
        chunks[current].emit_if_value(line);
        {
            chunks[current].emit_op_u16(Op::LOCAL_GET, tuple, line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_call(at, 2, line);
            io::emit_read_stream_to_string(&mut chunks[current], line);
        }
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
}

/// One line from `fnum`, trailing `\r\n` or `\n` stripped, position advanced
/// past it. Stack: `[fnum]` → `[string]`.
///
/// ⚠Reads from the position to END OF FILE and keeps the first line, so a loop
/// over an N-line file is O(N²) in bytes moved. The positioned read is what the
/// spec offers — there is no buffered reader in WASI, because a cursor is state
/// and `read-via-stream` is deliberately stateless. A guest-side buffer keyed by
/// file number would fix it and is the obvious next step if it ever matters.
pub fn emit_line_input(chunks: &mut [Chunk], current: usize, line: u32) {
    let fnum = chunks[current].alloc_scratch(1);
    let h = chunks[current].alloc_scratch(1);
    let rest = chunks[current].alloc_scratch(1);
    let nl = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fnum, line);

    emit_handle_of(chunks, current, fnum, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, h, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, h, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    {
        emit_rest_of_file(chunks, current, h, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, rest, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, rest, line);
        chunks[current].emit_string_const("\n", line);
        crate::primitives::strings::emit_index_of(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, nl, line);

        // No newline left: the remainder is the last line.
        chunks[current].emit_op_u16(Op::LOCAL_GET, nl, line);
        chunks[current].emit_f64_const(0.0, line);
        crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
        crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, rest, line);
        chunks[current].emit_else(line);
        {
            chunks[current].emit_op_u16(Op::LOCAL_GET, rest, line);
            chunks[current].emit_f64_const(0.0, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, nl, line);
            crate::primitives::strings::emit_substring(&mut chunks[current], line);
        }
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

        // Advance past the line AND its terminator, in bytes.
        emit_handle_field(chunks, current, h, H_POS, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        crate::primitives::strings::emit_byte_length(chunks, current, line);
        crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_f64_const(1.0, line);
        crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        emit_set_handle_field(chunks, current, h, H_POS, line);

        // A CRLF file leaves the `\r` on the near side of the split.
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_string_const("\r", line);
        chunks[current].emit_string_const("", line);
        crate::primitives::strings::emit_replace(&mut chunks[current], line);
    }
    chunks[current].emit_end(line);
}

/// `Input #n, a, b` — one line split on commas, each field trimmed.
/// Stack: `[fnum]` → `[array]`.
pub fn emit_input_file(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_line_input(chunks, current, line);
    crate::primitives::strings::emit_trim(&mut chunks[current], line);
    chunks[current].emit_string_const(",", line);
    crate::primitives::strings::emit_split(&mut chunks[current], line);
}

// ── Directory enumeration ─────────────────────────────────────────────────
//
// The last verb to migrate, and the only one whose stream carries something
// other than bytes. A const here used to say this was impossible; it stopped
// being true when `canon stream.read` learned element types, and a marker that
// outlives its blocker is worse than no marker at all.

/// `variant descriptor-type` — `types.wit:50`, in DECLARATION ORDER.
///
/// The guest's copy of the host's variant, and it has to be the guest's own:
/// `platforms/wasi` is the other side of this interface, not a library this
/// crate may reach into. Both sides derive their case list from the same WIT,
/// exactly as a bindgen-generated component and its host do.
///
/// ORDER IS THE WIRE FORMAT. `canon stream.read` writes a case's INDEX as the
/// discriminant, so a reordering here does not fail — it silently relabels
/// every entry, turning directories into fifos. The restored per-entry
/// assertions in `platforms/wasi/tests/wasi/filesystem*.rs` are what catch a
/// drift between the two spellings, which is why they assert real names and
/// types rather than a shape.
const DESCRIPTOR_TYPE_CASES: [&str; 8] = [
    "block-device",
    "character-device",
    "directory",
    "fifo",
    "symbolic-link",
    "regular-file",
    "socket",
    "other",
];

/// Field offsets and stride of `record directory-entry { %type:
/// descriptor-type, name: string }` (`types.wit:182`), in bytes.
///
/// Returned as `(stride, disc_offset, disc_bytes, name_offset)` and COMPUTED,
/// never written down: `canon_layout` is the same code the host lowers
/// through, so asking it is the only way the two cannot disagree. Hand-typing
/// `24 / 0 / 1 / 16` would be right today and wrong the day the WIT gains a
/// ninth `descriptor-type` case — which widens the discriminant, moves `name`,
/// and would leave this reading the middle of a string pointer.
fn directory_entry_layout() -> (i32, i32, u32, i32) {
    use vybe_runtime::canon_layout::{align_to, alignment, elem_size, variant_discriminant_size};
    use vybe_runtime::component::ValType;

    let cases: Vec<(String, Option<ValType>)> = DESCRIPTOR_TYPE_CASES
        .iter()
        .map(|name| {
            let payload = (*name == "other").then(|| ValType::Option(Box::new(ValType::String)));
            ((*name).to_string(), payload)
        })
        .collect();
    let disc_bytes = variant_discriminant_size(&cases);
    let type_field = ValType::Variant(cases);

    // `record` field offsets — `CanonicalABI.md` §`store_record`: each field
    // starts at the next offset aligned to ITS OWN alignment.
    let name_offset = align_to(elem_size(&type_field), alignment(&ValType::String));
    let entry = ValType::Record(vec![
        ("type".to_string(), type_field),
        ("name".to_string(), ValType::String),
    ]);

    (
        elem_size(&entry) as i32,
        0, // the discriminant sits at the start of the variant, which starts the record
        disc_bytes,
        name_offset as i32,
    )
}

/// How many `directory-entry` records one `canon stream.read` asks for.
///
/// Bounded and looped rather than one big read, for the reason
/// [`io::emit_read_stream_to_bytes`] gives: the buffer comes from the shared
/// bump allocator, which grows linear memory and never gives it back, so a
/// request sized to the largest imaginable directory would be a permanent
/// cost paid by every program that lists one.
const ENTRIES_PER_READ: i32 = 32;

/// `readDirEntries(path)` → `[{ type, name }]`, the WIT record verbatim.
/// Stack: `[path]` → `[array]`; an unreadable directory answers `[]`.
///
/// `read-directory: func() -> tuple<stream<directory-entry>,
///                                  future<result<_, error-code>>>`
///
/// This is the guest half of the typed-stream path: the host pushes
/// `directory-entry` values into a stream it TYPED with
/// `create_stream_of(Some(directory_entry_type()))`, `canon stream.read` lowers
/// each one into linear memory at its canonical stride, and the loop below
/// lifts them back. No `@N` type immediate is needed — the element type
/// travels with the stream, so the built-in already knows the layout it is
/// writing and this side only has to agree about where the fields land.
///
/// ⚠`type` is the WIT's spelling, not a language's. `isFile`/`isDir` are
/// PHP's and Python's questions, and each adapter derives them from this by
/// comparing against a `descriptor-type` case name. Answering their shape here
/// would put one language's vocabulary in the shared lowering.
pub fn emit_read_directory_entries(chunk: &mut Chunk, line: u32) {
    let (stride, disc_at, disc_bytes, name_at) = directory_entry_layout();

    let read_dir = chunk.add_import(TYPES, "[method]descriptor.read-directory");
    let at = chunk.add_import("ecma:array", "at");
    let push = chunk.add_import("ecma:array", "push");
    let new_arr = chunk.add_import("vybe:js-array", "newWithLength");
    let obj_new = chunk.add_import("ecma:object", "new");
    let obj_set = chunk.add_import("ecma:object", "set");
    let stream_read = chunk.add_import("canon", "stream.read");
    let drop_rd = chunk.add_import("canon", "stream.drop-readable");

    let path = chunk.alloc_scratch(1);
    let desc = chunk.alloc_scratch(1);
    let tuple = chunk.alloc_scratch(1);
    let handle = chunk.alloc_scratch(1);
    let out = chunk.alloc_scratch(1);
    let names = chunk.alloc_scratch(1);
    let packed = chunk.alloc_scratch(1);
    let count = chunk.alloc_scratch(1);
    let idx = chunk.alloc_scratch(1);
    let base = chunk.alloc_scratch(1);
    let entry = chunk.alloc_scratch(1);
    let sptr = chunk.alloc_scratch(1);
    let slen = chunk.alloc_scratch(1);
    let n = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, path, line);

    // OPEN_DIRECTORY, not 0: `open-at` on a directory without it is allowed to
    // fail, and asking for the thing we mean is what makes the error honest
    // when the path turns out to be a file.
    emit_open_at(chunk, line, path, OPEN_DIRECTORY, DESC_READ);
    emit_ok_test(chunk, line, desc);
    chunk.emit_if_value(line);
    {
        chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
        chunk.emit_call(read_dir, 1, line);
        emit_ok_test(chunk, line, tuple);
        chunk.emit_if_value(line);
        {
            // Element 0 of the tuple is the readable end. Element 1 carries a
            // PARTIAL enumeration's failure and is deliberately not consulted,
            // for the same reason `read-via-stream`'s is not: the entries that
            // did arrive are still entries.
            chunk.emit_op_u16(Op::LOCAL_GET, tuple, line);
            chunk.emit_i32_const(0, line);
            chunk.emit_call(at, 2, line);
            chunk.emit_op_u16(Op::LOCAL_SET, handle, line);

            chunk.emit_i32_const(0, line);
            chunk.emit_call(new_arr, 1, line);
            chunk.emit_op_u16(Op::LOCAL_SET, out, line);

            // The discriminant→name table, built once. Indexing it is what
            // turns a wire integer back into the WIT's own spelling.
            chunk.emit_i32_const(0, line);
            chunk.emit_call(new_arr, 1, line);
            chunk.emit_op_u16(Op::LOCAL_SET, names, line);
            for case in DESCRIPTOR_TYPE_CASES {
                chunk.emit_op_u16(Op::LOCAL_GET, names, line);
                chunk.emit_string_const(case, line);
                chunk.emit_call(push, 2, line);
                chunk.emit_op(Op::DROP, line);
            }

            // The buffer. `emit_alloc` hands back an 8-aligned address, which
            // covers `directory-entry`'s 4-byte alignment — `canon stream.read`
            // TRAPS on a misaligned element buffer rather than writing it
            // crooked, so this is a precondition and not a nicety.
            chunk.emit_i32_const(stride * ENTRIES_PER_READ, line);
            chunk.emit_op_u16(Op::LOCAL_SET, n, line);
            let ptr = canon_marshal::emit_alloc(chunk, line, n);

            let done = chunk.emit_block(line);
            let (drain, _) = chunk.emit_loop_s(line);

            chunk.emit_op_u16(Op::LOCAL_GET, handle, line);
            chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
            chunk.emit_i32_const(ENTRIES_PER_READ, line);
            chunk.emit_call(stream_read, 3, line);
            chunk.emit_op_u16(Op::LOCAL_SET, packed, line);

            // No BLOCKED arm: only the `async` variant of `canon stream.read`
            // may answer BLOCKED, and this is the synchronous one — it
            // suspends instead (see `primitives::io::emit_read_stream_to_bytes`
            // for the same change, and the `StreamRead` arm in `dispatch.rs`
            // for the runtime half). Breaking out on BLOCKED here would have
            // ended a directory listing early and reported it as complete.

            // count = packed >> 4, unsigned: the count is the top 28 bits.
            chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
            chunk.emit_i32_const(4, line);
            chunk.emit_op(Op::I32_SHR_U, line);
            chunk.emit_op_u16(Op::LOCAL_SET, count, line);

            chunk.emit_i32_const(0, line);
            chunk.emit_op_u16(Op::LOCAL_SET, idx, line);
            let lifted = chunk.emit_block(line);
            let (lift, _) = chunk.emit_loop_s(line);
            chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
            chunk.emit_op_u16(Op::LOCAL_GET, count, line);
            chunk.emit_op(Op::I32_GE_S, line);
            chunk.emit_br_if(1, line);

            // base = ptr + idx * stride
            chunk.emit_op_u16(Op::LOCAL_GET, ptr, line);
            chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
            chunk.emit_i32_const(stride, line);
            chunk.emit_op(Op::I32_MUL, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, base, line);

            chunk.emit_call(obj_new, 0, line);
            chunk.emit_op_u16(Op::LOCAL_SET, entry, line);

            // %type: the discriminant, then the case name it indexes.
            chunk.emit_op_u16(Op::LOCAL_GET, entry, line);
            chunk.emit_string_const("type", line);
            chunk.emit_op_u16(Op::LOCAL_GET, names, line);
            chunk.emit_op_u16(Op::LOCAL_GET, base, line);
            chunk.emit_i32_const(disc_at, line);
            chunk.emit_op(Op::I32_ADD, line);
            // The load WIDTH follows `variant_discriminant_size`. A fixed
            // 4-byte load would read the discriminant plus three padding bytes
            // — zero on a fresh page and stale entrails the second time round
            // the loop, i.e. a test that passes once and then lies.
            chunk.emit_op(
                match disc_bytes {
                    1 => Op::I32_LOAD8_U,
                    2 => Op::I32_LOAD16_U,
                    _ => Op::I32_LOAD,
                },
                line,
            );
            chunk.emit_call(at, 2, line);
            chunk.emit_call(obj_set, 3, line);
            chunk.emit_op(Op::DROP, line);

            // name: a (ptr, length) pair, decoded out of linear memory.
            chunk.emit_op_u16(Op::LOCAL_GET, base, line);
            chunk.emit_i32_const(name_at, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op(Op::I32_LOAD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, sptr, line);
            chunk.emit_op_u16(Op::LOCAL_GET, base, line);
            chunk.emit_i32_const(name_at + 4, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op(Op::I32_LOAD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, slen, line);

            chunk.emit_op_u16(Op::LOCAL_GET, entry, line);
            chunk.emit_string_const("name", line);
            canon_marshal::emit_load_utf8(chunk, line, sptr, slen);
            chunk.emit_call(obj_set, 3, line);
            chunk.emit_op(Op::DROP, line);

            chunk.emit_op_u16(Op::LOCAL_GET, out, line);
            chunk.emit_op_u16(Op::LOCAL_GET, entry, line);
            chunk.emit_call(push, 2, line);
            chunk.emit_op(Op::DROP, line);

            chunk.emit_op_u16(Op::LOCAL_GET, idx, line);
            chunk.emit_i32_const(1, line);
            chunk.emit_op(Op::I32_ADD, line);
            chunk.emit_op_u16(Op::LOCAL_SET, idx, line);
            chunk.emit_br(0, line);
            chunk.emit_end(line);
            chunk.patch_loop(lift);
            chunk.emit_end(line);
            chunk.patch_block(lifted);

            // Anything but COMPLETED (low nibble 0) ends the drain: DROPPED is
            // the end of the directory, and CANCELLED cannot occur here.
            chunk.emit_op_u16(Op::LOCAL_GET, packed, line);
            chunk.emit_i32_const(0xf, line);
            chunk.emit_op(Op::I32_AND, line);
            chunk.emit_br_if(1, line);

            chunk.emit_br(0, line);
            chunk.emit_end(line);
            chunk.patch_loop(drain);
            chunk.emit_end(line);
            chunk.patch_block(done);

            chunk.emit_op_u16(Op::LOCAL_GET, handle, line);
            chunk.emit_call(drop_rd, 1, line);

            chunk.emit_op_u16(Op::LOCAL_GET, out, line);
        }
        chunk.emit_else(line);
        chunk.emit_i32_const(0, line);
        chunk.emit_call(new_arr, 1, line);
        chunk.emit_end(line);
    }
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(new_arr, 1, line);
    chunk.emit_end(line);
}

/// `listDir(path)` → `[name]` — the names alone, in whatever order the
/// filesystem enumerates them. Stack: `[path]` → `[array of string]`.
///
/// ⚠NOT sorted, and not filtered. WASI's `read-directory` makes no ordering
/// promise, and `.` and `..` are not entries in it — POSIX `readdir` yields
/// them, `wasi:filesystem` does not, so a caller porting from `scandir` gets
/// two fewer entries than it used to. That difference belongs to the spec, not
/// to this lowering.
pub fn emit_list_dir(chunk: &mut Chunk, line: u32) {
    let push = chunk.add_import("ecma:array", "push");
    let new_arr = chunk.add_import("vybe:js-array", "newWithLength");
    let at = chunk.add_import("ecma:array", "at");
    let length = chunk.add_import("ecma:array", "length");

    let entries = chunk.alloc_scratch(1);
    let out = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);

    emit_read_directory_entries(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entries, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(new_arr, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    let done = chunk.emit_block(line);
    let (walk, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entries, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entries, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_call(at, 2, line);
    chunk.emit_string_const("name", line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(walk);
    chunk.emit_end(line);
    chunk.patch_block(done);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}
