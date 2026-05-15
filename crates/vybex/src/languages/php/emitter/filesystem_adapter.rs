//! PHP filesystem helpers — Rust inline opcode emitters.
//!
//! Every function (`file_exists`, `is_file`, `is_dir`, `filesize`,
//! `filemtime`, `unlink`, `file`, `dir` and the `Directory` iterator
//! methods) emits opcodes directly into `chunks[current]`, composing
//! `wasi:filesystem/preopens` + `wasi:filesystem/types` host fns and
//! `ecma:array.get` for the preopens-list indexing. No JS polyfill,
//! no shared chunk builders — every call site gets its own inline
//! sequence, mirroring the dotnet datetime_adapter / php datetime_adapter
//! pattern.

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

/// Push the current preopen descriptor: `get-directories()[0][0]`.
fn emit_preopen_descriptor(chunks: &mut [Chunk], current: usize, line: u32) {
    call_import(chunks, current, "wasi:filesystem/preopens", "get-directories", 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::I32_CONST_0, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "get", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::I32_CONST_0, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "get", 2, line);
}

/// Push `descriptor.stat-at(0, path)` — leaves the stat record on the
/// stack. Stack on entry: `[path]` ; Stack on exit: `[stat_record]`.
fn emit_stat_at(chunks: &mut [Chunk], current: usize, line: u32) {
    let path_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        lset(chunk, s, line);
        s
    };
    emit_preopen_descriptor(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::I32_CONST_0, line); // path-flags
    lget(chunk, path_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem/types", "[method]descriptor.stat-at", 3, line);
}

/// PHP `file_exists($path)` — true iff stat-at returns a record with
/// a non-null `type` field.
pub fn emit_file_exists(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stat_at(chunks, current, line);
    let chunk = &mut chunks[current];
    let key = chunk.add_constant(Value::String(Arc::from("type")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::DYN_NOT, line);
}

/// PHP `is_file($path)` — stat-at then `type === "regular-file"`.
pub fn emit_is_file(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stat_type_match(chunks, current, "regular-file", line);
}

/// PHP `is_dir($path)` — stat-at then `type === "directory"`.
pub fn emit_is_dir(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stat_type_match(chunks, current, "directory", line);
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

fn emit_stat_type_match(chunks: &mut [Chunk], current: usize, expected: &str, line: u32) {
    emit_stat_at(chunks, current, line);
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from("type")));
    chunk.emit_op_u16(Op::STRUCT_GET, type_key, line);
    push_str(chunk, expected, line);
    chunk.emit_op(Op::DYN_EQ, line);
}

/// PHP `filesize($path)` — stat-at then read `size` field.
pub fn emit_filesize(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stat_at(chunks, current, line);
    let chunk = &mut chunks[current];
    let key = chunk.add_constant(Value::String(Arc::from("size")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
}

/// PHP `filemtime($path)` — stat-at, read
/// `data-modification-timestamp` (ms), divide by 1000 → secs.
pub fn emit_filemtime(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stat_at(chunks, current, line);
    let chunk = &mut chunks[current];
    let key = chunk.add_constant(Value::String(Arc::from("data-modification-timestamp")));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
}

/// PHP `unlink($path)` — `[method]descriptor.unlink-file-at(parent, path)`.
/// Returns null on success, error object on failure (matches WIT result).
pub fn emit_unlink(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let path_slot = alloc_local(chunk);
    // PHP `unlink($path, $context = null)` — drop optional context arg.
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line);
    }
    lset(chunk, path_slot, line);
    emit_preopen_descriptor(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, path_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem/types", "[method]descriptor.unlink-file-at", 2, line);
}

/// PHP `readlink($path)` — `[method]descriptor.readlink-at(parent, path)`.
pub fn emit_readlink(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let path_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        lset(chunk, s, line);
        s
    };
    emit_preopen_descriptor(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, path_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem/types", "[method]descriptor.readlink-at", 2, line);
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
    call_import(chunks, current, "wasi:filesystem", "readFile", 1, line);
    let chunk = &mut chunks[current];
    push_str(chunk, "\n", line);
    chunk.emit_op(Op::STR_SPLIT, line);
}

/// PHP `dir($path)` — open the path as a directory and return an
/// iterator object `{__type: "Directory", __stream}`. The walker
/// rewrites `$dir->read()` / `$dir->close()` to `__php_dir_read($dir)`
/// / `__php_dir_close($dir)`, both also inline-emitted.
pub fn emit_dir(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let path_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        lset(chunk, s, line);
        s
    };

    // descriptor = preopen.open-at(0, path, OPEN_DIRECTORY=2, 0)
    emit_preopen_descriptor(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::I32_CONST_0, line);
    lget(chunk, path_slot, line);
    push_const(chunk, Value::I32(2), line); // open-flags::directory
    chunk.emit_op(Op::I32_CONST_0, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem/types", "[method]descriptor.open-at", 5, line);
    let chunk = &mut chunks[current];
    let descriptor_slot = alloc_local(chunk);
    lset(chunk, descriptor_slot, line);

    // stream = descriptor.read-directory()
    lget(chunk, descriptor_slot, line);
    let _ = chunk;
    call_import(chunks, current, "wasi:filesystem/types", "[method]descriptor.read-directory", 1, line);
    let chunk = &mut chunks[current];
    let stream_slot = alloc_local(chunk);
    lset(chunk, stream_slot, line);

    // wrapper = STRUCT_NEW; wrapper.__type = "Directory"; wrapper.__stream = stream
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op(Op::DUP, line);
    push_str(chunk, "Directory", line);
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::DUP, line);
    lget(chunk, stream_slot, line);
    let stream_key = chunk.add_constant(Value::String(Arc::from("__stream")));
    chunk.emit_op_u16(Op::STRUCT_SET, stream_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// PHP `Directory->read()` — pulls the next entry name from
/// `$dir->__stream`. End-of-stream returns `false` for compatibility
/// with `while ($f = $dir->read()) { ... }`.
pub fn emit_dir_read(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // Stack on entry: [$dir]
    let chunk = &mut chunks[current];
    let stream_key = chunk.add_constant(Value::String(Arc::from("__stream")));
    chunk.emit_op_u16(Op::STRUCT_GET, stream_key, line);
    let _ = chunk;
    call_import(
        chunks, current,
        "wasi:filesystem/types",
        "[method]directory-entry-stream.read-directory-entry",
        1, line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let if_null = chunk.emit_jump(Op::BR_IF_TRUE, line);
    // Not null: extract `name`.
    let name_key = chunk.add_constant(Value::String(Arc::from("name")));
    chunk.emit_op_u16(Op::STRUCT_GET, name_key, line);
    let done = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(if_null);
    chunk.emit_op(Op::DROP, line);
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
