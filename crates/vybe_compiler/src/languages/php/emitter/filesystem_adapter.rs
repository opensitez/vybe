//! PHP filesystem helpers — Rust inline opcode emitters.
//!
//! Every function (`basename`, `dirname`, `file_get_contents`,
//! `file_put_contents`, `mkdir`, `file_exists`, `is_file`, `is_dir`,
//! `filesize`, `filemtime`, `unlink`, `file`, `dir` and the
//! `Directory` iterator methods) emits opcodes directly into
//! `chunks[current]`, composing the existing `wasi:filesystem` and
//! `node:fs` host surfaces. No JS polyfill, no shared chunk builders —
//! every call site gets its own inline sequence, mirroring the dotnet
//! datetime_adapter / php datetime_adapter pattern.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[0].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}

/// Push `wasi:filesystem.stat(path)` — leaves the stat record or null
/// on the stack. Stack on entry: `[path]` ; Stack on exit:
/// `[stat_record_or_null]`.
fn emit_stat_at(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "wasi:filesystem", "stat", 1, line);
}

/// PHP `basename($path, $suffix = null)` — current runtime behavior is a
/// direct path filename extraction; optional suffix remains ignored.
pub fn emit_basename(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "pathGetFileName", 1, line);
}

/// PHP `dirname($path, $levels = 1)` — current runtime behavior is a
/// single directory extraction; optional levels remains ignored.
pub fn emit_dirname(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "pathGetDirectory", 1, line);
}

/// PHP `file_get_contents($path, ...)` — current MVP behavior forwards to
/// the text filesystem shim and ignores the optional stream/context args.
pub fn emit_file_get_contents(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "readFile", 1, line);
}

/// PHP `file_put_contents($path, $data, ...)` — forwards the required
/// `(path, data)` pair and ignores the optional flags/context args.
pub fn emit_file_put_contents(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "writeFile", 2, line);
}

/// PHP `mkdir($path, ...)` — forwards the path and ignores optional mode /
/// recursive / context arguments, matching the prior direct-host binding.
pub fn emit_mkdir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "mkdir", 1, line);
}

/// PHP `file_exists($path)` — true iff the filesystem host reports the
/// path exists.
pub fn emit_file_exists(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "wasi:filesystem", "exists", 1, line);
}

/// PHP `is_file($path)` — direct file check through the runtime fs shim.
pub fn emit_is_file(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "wasi:filesystem", "isFile", 1, line);
}

/// PHP `is_dir($path)` — direct directory check through the runtime fs shim.
pub fn emit_is_dir(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "wasi:filesystem", "isDir", 1, line);
}

/// PHP `is_link($path)` — `node:fs.lstatSync(path)` then
/// `_statIsSymbolicLink(stats)` so symlinks are not followed.
pub fn emit_is_link(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let path_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        lset(chunk, s, line);
        s
    };
    let chunk = &mut chunks[current];
    lget(chunk, path_slot, line);
    let _ = chunk;
    call_import(chunks, current, "node:fs", "lstatSync", 1, line);
    call_import(chunks, current, "node:fs", "_statIsSymbolicLink", 1, line);
}

/// PHP `filesize($path)` — direct size query through the runtime fs shim.
pub fn emit_filesize(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "wasi:filesystem", "fileSize", 1, line);
}

/// PHP `filemtime($path)` — stat(path).modified (ms), divided by 1000
/// to match PHP's epoch-seconds result.
pub fn emit_filemtime(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stat_at(chunks, current, line);
    let chunk = &mut chunks[current];
    let key = chunk.add_constant(Value::String(Arc::from("modified")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
}

/// PHP `unlink($path)` — bool success/failure via the runtime fs shim.
pub fn emit_unlink(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // PHP `unlink($path, $context = null)` — drop optional context arg.
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "remove", 1, line);
}

/// PHP `readlink($path)` — Node-aligned readlink surface. Kept here
/// because the higher-level `wasi:filesystem` shim does not expose
/// symlink targets.
pub fn emit_readlink(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "node:fs", "readlinkSync", 1, line);
}

/// PHP `pathinfo($path)` — returns a PHP-shaped record with
/// `dirname`, `basename`, `extension`, and `filename`.
pub fn emit_pathinfo(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line);
    }
    let path_slot = alloc_local(chunk);
    lset(chunk, path_slot, line);

    let dirname_slot = alloc_local(chunk);
    let basename_slot = alloc_local(chunk);
    let extension_slot = alloc_local(chunk);
    let filename_slot = alloc_local(chunk);

    lget(chunk, path_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "pathGetDirectory", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, dirname_slot, line);

    lget(chunk, path_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "pathGetFileName", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, basename_slot, line);

    lget(chunk, path_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "pathGetExtension", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, extension_slot, line);

    lget(chunk, path_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "pathGetFileNameWithoutExt", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, filename_slot, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    lget(chunk, dirname_slot, line);
    let dirname_key = chunk.add_constant(Value::String(Arc::from("dirname")));
    chunk.emit_op_u16(Op::STRUCT_SET, dirname_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::DUP, line);
    lget(chunk, basename_slot, line);
    let basename_key = chunk.add_constant(Value::String(Arc::from("basename")));
    chunk.emit_op_u16(Op::STRUCT_SET, basename_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::DUP, line);
    lget(chunk, extension_slot, line);
    let extension_key = chunk.add_constant(Value::String(Arc::from("extension")));
    chunk.emit_op_u16(Op::STRUCT_SET, extension_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::DUP, line);
    lget(chunk, filename_slot, line);
    let filename_key = chunk.add_constant(Value::String(Arc::from("filename")));
    chunk.emit_op_u16(Op::STRUCT_SET, filename_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// PHP `file($path)` — read whole file, split on `"\n"`, return
/// array of lines. PHP supports flags as args 2/3 (FILE_IGNORE_NEW_LINES,
/// FILE_SKIP_EMPTY_LINES); MVP ignores them.
pub fn emit_file(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // Drop the optional flag args — MVP semantics.
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    let chunk = &mut chunks[current];
    push_str(chunk, "utf8", line);
    let _ = chunk;
    call_import(chunks, current, "node:fs", "readFileSync", 2, line);
    let chunk = &mut chunks[current];
    push_str(chunk, "\n", line);
    chunk.emit_op(Op::STR_SPLIT, line);
}

/// PHP `glob($pattern, $flags = 0)` — current support covers the common
/// single-`*` filename wildcard form by listing the directory and
/// filtering entries with prefix/suffix checks. Returns matching full paths.
pub fn emit_glob(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }

    let pattern_slot = alloc_local(chunk);
    let dir_slot = alloc_local(chunk);
    let file_pattern_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);
    lset(chunk, pattern_slot, line);

    lget(chunk, pattern_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "pathGetDirectory", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, dir_slot, line);

    lget(chunk, pattern_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "pathGetFileName", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, file_pattern_slot, line);

    lget(chunk, file_pattern_slot, line);
    push_str(chunk, "*", line);
    chunk.emit_op(Op::STR_CONTAINS, line);
    let no_wildcard = chunk.emit_jump(Op::BR_IF_FALSE, line);

    let parts_slot = alloc_local(chunk);
    let parts_len_slot = alloc_local(chunk);
    let last_part_idx_slot = alloc_local(chunk);
    let prefix_slot = alloc_local(chunk);
    let suffix_slot = alloc_local(chunk);
    let entries_slot = alloc_local(chunk);
    let entries_len_slot = alloc_local(chunk);
    let index_slot = alloc_local(chunk);
    let entry_slot = alloc_local(chunk);
    let full_path_slot = alloc_local(chunk);

    lget(chunk, file_pattern_slot, line);
    push_str(chunk, "*", line);
    chunk.emit_op(Op::STR_SPLIT, line);
    lset(chunk, parts_slot, line);

    lget(chunk, parts_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, parts_len_slot, line);

    lget(chunk, parts_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, prefix_slot, line);

    lget(chunk, parts_len_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, last_part_idx_slot, line);

    lget(chunk, parts_slot, line);
    lget(chunk, last_part_idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, suffix_slot, line);

    lget(chunk, dir_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "listDir", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, entries_slot, line);

    lget(chunk, entries_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, entries_len_slot, line);

    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, index_slot, line);

    let loop_top = chunk.code.len();
    lget(chunk, index_slot, line);
    lget(chunk, entries_len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let loop_done = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, entries_slot, line);
    lget(chunk, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, entry_slot, line);

    lget(chunk, prefix_slot, line);
    push_str(chunk, "", line);
    chunk.emit_op(Op::DYN_EQ, line);
    let skip_prefix_check = chunk.emit_jump(Op::BR_IF_TRUE, line);
    lget(chunk, entry_slot, line);
    lget(chunk, prefix_slot, line);
    chunk.emit_op(Op::STR_STARTS_WITH, line);
    let next_after_prefix = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.patch_jump(skip_prefix_check);

    lget(chunk, suffix_slot, line);
    push_str(chunk, "", line);
    chunk.emit_op(Op::DYN_EQ, line);
    let skip_suffix_check = chunk.emit_jump(Op::BR_IF_TRUE, line);
    lget(chunk, entry_slot, line);
    lget(chunk, suffix_slot, line);
    chunk.emit_op(Op::STR_ENDS_WITH, line);
    let next_after_suffix = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.patch_jump(skip_suffix_check);

    lget(chunk, dir_slot, line);
    push_str(chunk, "/", line);
    chunk.emit_op(Op::STR_CONCAT, line);
    lget(chunk, entry_slot, line);
    chunk.emit_op(Op::STR_CONCAT, line);
    lset(chunk, full_path_slot, line);

    lget(chunk, full_path_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "isFile", 1, line);
    let chunk = &mut chunks[current];
    let next_after_file_check = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, result_slot, line);
    lget(chunk, full_path_slot, line);
    let _ = chunk;
    crate::emitter::collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    chunk.patch_jump(next_after_prefix);
    chunk.patch_jump(next_after_suffix);
    chunk.patch_jump(next_after_file_check);
    lget(chunk, index_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, index_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(loop_done);

    let wildcard_done = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(no_wildcard);

    lget(chunk, pattern_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "exists", 1, line);
    let chunk = &mut chunks[current];
    let exact_missing = chunk.emit_jump(Op::BR_IF_FALSE, line);

    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);
    lget(chunk, result_slot, line);
    lget(chunk, pattern_slot, line);
    let _ = chunk;
    crate::emitter::collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    let exact_done = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(exact_missing);
    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);
    chunk.patch_jump(exact_done);

    chunk.patch_jump(wildcard_done);
    lget(chunk, result_slot, line);
}

/// PHP `dir($path)` — materialize directory entries and return an
/// iterator object `{__type: "Directory", __entries, __index}`. The walker
/// rewrites `$dir->read()` / `$dir->close()` to `__php_dir_read($dir)`
/// / `__php_dir_close($dir)`, both also inline-emitted.
pub fn emit_dir(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let entries_slot = alloc_local(chunk);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem", "listDir", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, entries_slot, line);

    // wrapper = STRUCT_NEW; wrapper.__type = "Directory";
    // wrapper.__entries = entries; wrapper.__index = 0
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    push_str(chunk, "Directory", line);
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::DUP, line);
    lget(chunk, entries_slot, line);
    let entries_key = chunk.add_constant(Value::String(Arc::from("__entries")));
    chunk.emit_op_u16(Op::STRUCT_SET, entries_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op(Op::DUP, line);
    push_const(chunk, Value::F64(0.0), line);
    let index_key = chunk.add_constant(Value::String(Arc::from("__index")));
    chunk.emit_op_u16(Op::STRUCT_SET, index_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// PHP `Directory->read()` — pulls the next entry name from
/// `$dir->__entries[$dir->__index]`. End-of-stream returns `false` for compatibility
/// with `while ($f = $dir->read()) { ... }`.
pub fn emit_dir_read(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let dir_slot = alloc_local(chunk);
    let entries_slot = alloc_local(chunk);
    let index_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let entry_slot = alloc_local(chunk);
    let entries_key = chunk.add_constant(Value::String(Arc::from("__entries")));
    let index_key = chunk.add_constant(Value::String(Arc::from("__index")));
    lset(chunk, dir_slot, line);

    lget(chunk, dir_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, entries_key, line);
    lset(chunk, entries_slot, line);

    lget(chunk, dir_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, index_key, line);
    lset(chunk, index_slot, line);

    lget(chunk, entries_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    lget(chunk, index_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let has_entry = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, entries_slot, line);
    lget(chunk, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, entry_slot, line);

    lget(chunk, dir_slot, line);
    chunk.emit_op(Op::DUP, line);
    lget(chunk, index_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.emit_op_u16(Op::STRUCT_SET, index_key, line);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, entry_slot, line);
    let done = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(has_entry);
    chunk.emit_op(Op::FALSE, line);
    chunk.patch_jump(done);
}

/// PHP `Directory->close()` — no-op (stream resource is dropped on
/// gc). Returns null for source-compatibility with `$dir->close();`.
pub fn emit_dir_close(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line); // drop receiver
    chunk.emit_op(Op::NULL, line);
}
