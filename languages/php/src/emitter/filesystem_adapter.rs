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

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::{fs_path, paths};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

/// Push `wasi:filesystem.stat(path)` — leaves the stat record or null
/// on the stack. Stack on entry: `[path]` ; Stack on exit:
/// `[stat_record_or_null]`.
fn emit_stat_at(chunks: &mut [Chunk], current: usize, line: u32) {
    fs_path::emit_stat(&mut chunks[current], line);
}

/// PHP `basename($path, $suffix = "")` — filename component, with an
/// optional trailing suffix stripped (unless it equals the whole name).
pub fn emit_basename(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        paths::emit_file_name(&mut chunks[current], line);
        return;
    }
    // stack: [path, suffix]
    let (suffix_slot, path_slot, name_slot) = {
        let chunk = &mut chunks[current];
        (alloc_local(chunk), alloc_local(chunk), alloc_local(chunk))
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, suffix_slot, line);
        lset(chunk, path_slot, line);
        lget(chunk, path_slot, line);
    }
    paths::emit_file_name(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, name_slot, line);

    // if name.endsWith(suffix) && name != suffix { strip suffix }
    lget(chunk, name_slot, line);
    lget(chunk, suffix_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "endsWith");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, name_slot, line);
    lget(chunk, suffix_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // name = name.substring(0, name.length - suffix.length)
    lget(chunk, name_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, name_slot, line);
    str_len(chunk, line);
    lget(chunk, suffix_slot, line);
    str_len(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_neg(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    {
        let idx = chunk.add_import("ecma:string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, name_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    lget(chunk, name_slot, line);
}

/// PHP `dirname($path, $levels = 1)` — walks up `$levels` directory
/// components (default 1).
pub fn emit_dirname(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        paths::emit_directory(&mut chunks[current], line);
        return;
    }
    // stack: [path, levels]
    let (path_slot, n_slot) = {
        let chunk = &mut chunks[current];
        (alloc_local(chunk), alloc_local(chunk))
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, n_slot, line);
        lset(chunk, path_slot, line);
    }
    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, n_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, path_slot, line);
    }
    paths::emit_directory(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, path_slot, line);
        lget(chunk, n_slot, line);
        push_const(chunk, Value::F64(-1.0), line);
        vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
        lset(chunk, n_slot, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, path_slot, line);
    }
}

/// PHP `file_get_contents($path, ...)` — reads the whole file, ignoring the
/// optional stream/context args.
///
/// PHP answers `false` on failure, so the `null` the spec lowering reports is
/// mapped here. The verb this replaced answered the string
/// `"Error: No such file or directory (os error 2)"` AS THE CONTENTS, which
/// `file_get_contents` can never return — `false` is the documented failure.
pub fn emit_file_get_contents(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    fs_path::emit_read_file(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    let text = alloc_local(chunk);
    lset(chunk, text, line);
    lget(chunk, text, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    lget(chunk, text, line);
    chunk.emit_end(line);
}

/// PHP `file_put_contents($path, $data, $flags = 0, ...)` — writes (or,
/// with `FILE_APPEND`, appends) `$data` to `$path`. The only PHP flags
/// are `FILE_USE_INCLUDE_PATH`(1), `LOCK_EX`(2), `FILE_APPEND`(8), so any
/// value `>= 8` carries the append bit.
pub fn emit_file_put_contents(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        fs_path::emit_write_file(&mut chunks[current], line);
        return;
    }
    // stack: [path, data, flags, (context?)]
    let (flags_slot, data_slot, path_slot) = {
        let chunk = &mut chunks[current];
        // drop optional context beyond the flags arg
        for _ in 3..argc {
            chunk.emit_op(Op::DROP, line);
        }
        (alloc_local(chunk), alloc_local(chunk), alloc_local(chunk))
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, flags_slot, line);
        lset(chunk, data_slot, line);
        lset(chunk, path_slot, line);
        // if flags >= FILE_APPEND(8) -> append, else write
        lget(chunk, flags_slot, line);
        push_const(chunk, Value::F64(8.0), line);
        vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, path_slot, line);
        lget(chunk, data_slot, line);
    }
    fs_path::emit_append_file(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_else(line);
        lget(chunk, path_slot, line);
        lget(chunk, data_slot, line);
    }
    fs_path::emit_write_file(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_end(line);
    }
}

/// PHP `mkdir($path, ...)` — forwards the path and ignores optional mode /
/// recursive / context arguments, matching the prior direct-host binding.
pub fn emit_mkdir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let _ = chunk;
    fs_path::emit_mkdir(&mut chunks[current], line);
}

/// PHP `file_exists($path)` — true iff the filesystem host reports the
/// path exists.
pub fn emit_file_exists(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    fs_path::emit_exists(&mut chunks[current], line);
}

/// PHP `is_file($path)` — direct file check through the runtime fs shim.
pub fn emit_is_file(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    fs_path::emit_is_file(&mut chunks[current], line);
}

/// PHP `is_dir($path)` — direct directory check through the runtime fs shim.
pub fn emit_is_dir(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    fs_path::emit_is_dir(&mut chunks[current], line);
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

/// PHP `filesize($path)` — for an in-memory stream uri, the registered
/// buffer length; otherwise the on-disk size via the fs host.
pub fn emit_filesize(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // stack: [path]
    let (path_slot, reg_slot) = {
        let chunk = &mut chunks[current];
        (alloc_local(chunk), alloc_local(chunk))
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, path_slot, line);
        vybe_compiler::primitives::globals::emit_read(chunk, "__php_stream_registry", line);
        lset(chunk, reg_slot, line);
        // if registry != null && registry.has(path)
        lget(chunk, reg_slot, line);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, reg_slot, line);
        lget(chunk, path_slot, line);
    }
    call_import(chunks, current, "ecma:map", "has", 2, line);
    {
        let chunk = &mut chunks[current];
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        // registry.get(path).__buf.length
        lget(chunk, reg_slot, line);
        lget(chunk, path_slot, line);
    }
    call_import(chunks, current, "ecma:map", "get", 2, line);
    {
        let chunk = &mut chunks[current];
        struct_get_key(chunk, "__buf", line);
        str_len(chunk, line);
        chunk.emit_else(line);
        lget(chunk, path_slot, line);
    }
    fs_path::emit_file_size(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_end(line); // end inner if
        chunk.emit_else(line); // registry was null
        lget(chunk, path_slot, line);
    }
    fs_path::emit_file_size(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_end(line); // end outer if
    }
}

/// PHP `filemtime($path)` — stat(path).modified (ms), divided by 1000
/// to match PHP's epoch-seconds result.
pub fn emit_filemtime(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stat_at(chunks, current, line);
    let chunk = &mut chunks[current];
    let key = chunk.add_constant(Value::String(Arc::from("modified")));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
    push_const(chunk, Value::F64(1000.0), line);
    chunk.emit_op(Op::F64_DIV, line);
}

/// PHP `unlink($path)` — bool success/failure via the runtime fs shim.
pub fn emit_unlink(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        // PHP `unlink($path, $context = null)` — drop optional context arg.
        if argc >= 2 {
            chunk.emit_op(Op::DROP, line);
        }
    }
    fs_path::emit_unlink(&mut chunks[current], line);
    let result_slot = {
        let chunk = &mut chunks[current];
        let slot = alloc_local(chunk);
        lset(chunk, slot, line);
        lget(chunk, slot, line);
        push_const(chunk, Value::Bool(false), line);
        vybe_compiler::primitives::ops::emit_js_strict_eq(chunk, line);
        chunk.emit_if(line);
        push_str(chunk, "unlink failed", line);
        push_const(chunk, Value::F64(512.0), line);
        slot
    };
    super::error_adapter::emit_trigger_error(chunks, current, 2, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);
        lget(chunk, result_slot, line);
    }
}

/// PHP `readlink($path)` — Node-aligned readlink surface. Kept here
/// because the higher-level `wasi:filesystem` shim does not expose
/// symlink targets.
pub fn emit_readlink(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    call_import(chunks, current, "node:fs", "readlinkSync", 1, line);
}

/// PHP `pathinfo($path)` — returns a PHP-shaped record with
/// `dirname`, `basename`, `extension`, and `filename`.
/// `pathinfo($path, $flag)` — returns the single requested component.
/// PATHINFO_DIRNAME(1), PATHINFO_BASENAME(2), PATHINFO_EXTENSION(4),
/// PATHINFO_FILENAME(8).
fn emit_pathinfo_flag(chunks: &mut [Chunk], current: usize, line: u32) {
    // stack: [path, flag]
    let (flag_slot, path_slot) = {
        let chunk = &mut chunks[current];
        (alloc_local(chunk), alloc_local(chunk))
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, flag_slot, line);
        lset(chunk, path_slot, line);
        // if flag == 1 (DIRNAME)
        lget(chunk, flag_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, path_slot, line);
    }
    paths::emit_directory(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_else(line);
        // if flag == 2 (BASENAME)
        lget(chunk, flag_slot, line);
        push_const(chunk, Value::F64(2.0), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, path_slot, line);
    }
    paths::emit_file_name(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_else(line);
        // if flag == 4 (EXTENSION)
        lget(chunk, flag_slot, line);
        push_const(chunk, Value::F64(4.0), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        lget(chunk, path_slot, line);
    }
    paths::emit_extension(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_else(line);
        // else FILENAME (8)
        lget(chunk, path_slot, line);
    }
    paths::emit_file_stem(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }
}

pub fn emit_pathinfo(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        emit_pathinfo_flag(chunks, current, line);
        return;
    }
    let chunk = &mut chunks[current];
    let path_slot = alloc_local(chunk);
    lset(chunk, path_slot, line);

    let dirname_slot = alloc_local(chunk);
    let basename_slot = alloc_local(chunk);
    let extension_slot = alloc_local(chunk);
    let filename_slot = alloc_local(chunk);

    lget(chunk, path_slot, line);
    let _ = chunk;
    paths::emit_directory(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, dirname_slot, line);

    lget(chunk, path_slot, line);
    let _ = chunk;
    paths::emit_file_name(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, basename_slot, line);

    lget(chunk, path_slot, line);
    let _ = chunk;
    paths::emit_extension(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, extension_slot, line);

    lget(chunk, path_slot, line);
    let _ = chunk;
    paths::emit_file_stem(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, filename_slot, line);

    chunk.emit_struct_new(0, 0, line);
    chunk.emit_dup(line);
    lget(chunk, dirname_slot, line);
    let dirname_key = chunk.add_constant(Value::String(Arc::from("dirname")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, dirname_key, line);

    chunk.emit_dup(line);
    lget(chunk, basename_slot, line);
    let basename_key = chunk.add_constant(Value::String(Arc::from("basename")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, basename_key, line);

    chunk.emit_dup(line);
    lget(chunk, extension_slot, line);
    let extension_key = chunk.add_constant(Value::String(Arc::from("extension")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, extension_key, line);

    chunk.emit_dup(line);
    lget(chunk, filename_slot, line);
    let filename_key = chunk.add_constant(Value::String(Arc::from("filename")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, filename_key, line);
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
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
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
    paths::emit_directory(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, dir_slot, line);

    lget(chunk, pattern_slot, line);
    let _ = chunk;
    paths::emit_file_name(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, file_pattern_slot, line);

    lget(chunk, file_pattern_slot, line);
    push_str(chunk, "*", line);
    {
        let idx = chunk.add_import("ecma:string", "includes");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

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
    let pass_slot = alloc_local(chunk);

    lget(chunk, file_pattern_slot, line);
    push_str(chunk, "*", line);
    {
        let idx = chunk.add_import("ecma:string", "split");
        chunk.emit_call(idx, 2, line);
    }
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
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, last_part_idx_slot, line);

    lget(chunk, parts_slot, line);
    lget(chunk, last_part_idx_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, suffix_slot, line);

    lget(chunk, dir_slot, line);
    let _ = chunk;
    fs_path::emit_list_dir(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, entries_slot, line);

    lget(chunk, entries_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, entries_len_slot, line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, index_slot, line);

    let (loop_patch, _) = chunk.emit_loop_s(line);
    lget(chunk, index_slot, line);
    lget(chunk, entries_len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    let skip_entry = chunk.emit_block(line);

    lget(chunk, entries_slot, line);
    lget(chunk, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, entry_slot, line);

    push_const(chunk, Value::Bool(true), line);
    lset(chunk, pass_slot, line);
    lget(chunk, prefix_slot, line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    lget(chunk, entry_slot, line);
    lget(chunk, prefix_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lset(chunk, pass_slot, line);
    chunk.emit_end(line);
    lget(chunk, pass_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);

    push_const(chunk, Value::Bool(true), line);
    lset(chunk, pass_slot, line);
    lget(chunk, suffix_slot, line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    lget(chunk, entry_slot, line);
    lget(chunk, suffix_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "endsWith");
        chunk.emit_call(idx, 2, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    lset(chunk, pass_slot, line);
    chunk.emit_end(line);
    lget(chunk, pass_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);

    lget(chunk, dir_slot, line);
    push_str(chunk, "/", line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lget(chunk, entry_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, full_path_slot, line);

    lget(chunk, full_path_slot, line);
    let _ = chunk;
    fs_path::emit_is_file(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);

    lget(chunk, result_slot, line);
    lget(chunk, full_path_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.patch_block(skip_entry);
    lget(chunk, index_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, index_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);

    chunk.emit_else(line);

    lget(chunk, pattern_slot, line);
    let _ = chunk;
    fs_path::emit_exists(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);
    lget(chunk, result_slot, line);
    lget(chunk, pattern_slot, line);
    let _ = chunk;
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    chunk.emit_else(line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
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
    fs_path::emit_list_dir(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, entries_slot, line);

    // wrapper = STRUCT_NEW; wrapper.__type = "Directory";
    // wrapper.__entries = entries; wrapper.__index = 0
    chunk.emit_struct_new(0, 0, line);
    chunk.emit_dup(line);
    push_str(chunk, "Directory", line);
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);

    chunk.emit_dup(line);
    lget(chunk, entries_slot, line);
    let entries_key = chunk.add_constant(Value::String(Arc::from("__entries")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, entries_key, line);

    chunk.emit_dup(line);
    push_const(chunk, Value::F64(0.0), line);
    let index_key = chunk.add_constant(Value::String(Arc::from("__index")));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, index_key, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, entries_key, line);
    lset(chunk, entries_slot, line);

    lget(chunk, dir_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, index_key, line);
    lset(chunk, index_slot, line);

    lget(chunk, entries_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    lget(chunk, index_slot, line);
    lget(chunk, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);

    lget(chunk, entries_slot, line);
    lget(chunk, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, entry_slot, line);

    lget(chunk, dir_slot, line);
    chunk.emit_dup(line);
    lget(chunk, index_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, index_key, line);
    lget(chunk, entry_slot, line);

    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);
}

/// PHP `Directory->close()` — no-op (stream resource is dropped on
/// gc). Returns null for source-compatibility with `$dir->close();`.
pub fn emit_dir_close(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line); // drop receiver
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

// ── small local helpers ─────────────────────────────────────────────
fn struct_set_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, idx, line);
}
fn struct_get_key(chunk: &mut Chunk, key: &str, line: u32) {
    let idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, idx, line);
}
fn str_len(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(idx, 1, line);
}
fn str_concat(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "concat");
    chunk.emit_call(idx, 2, line);
}

/// PHP `sys_get_temp_dir()` — the host temp directory.
pub fn emit_sys_get_temp_dir(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    paths::emit_temp_path(&mut chunks[current], line);
}

/// PHP `realpath($path)` — canonical absolute path.
pub fn emit_realpath(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    paths::emit_full_path(&mut chunks[current], line);
}

/// PHP `copy($src, $dst, $context = null)` — drops the optional context.
pub fn emit_copy(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        chunks[current].emit_op(Op::DROP, line);
    }
    fs_path::emit_copy(&mut chunks[current], line);
}

/// PHP `rename($old, $new, $context = null)` — drops the optional context.
pub fn emit_rename(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        chunks[current].emit_op(Op::DROP, line);
    }
    fs_path::emit_rename(&mut chunks[current], line);
}

/// PHP `rmdir($dir, $context = null)` — drops the optional context.
pub fn emit_rmdir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        chunks[current].emit_op(Op::DROP, line);
    }
    // `remove-directory-at`, not the kind-dispatching remove: PHP's `rmdir()`
    // fails on a plain file and on a NON-EMPTY directory, exactly as POSIX
    // `rmdir` does. The shim it replaces called `remove_dir_all`, so a
    // non-empty directory was silently deleted RECURSIVELY.
    fs_path::emit_rmdir(&mut chunks[current], line);
}

/// PHP `is_readable($path)` / `is_writable($path)` — approximated by
/// existence through the fs host (the sandbox grants access to what it
/// exposes).
pub fn emit_is_readable(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    fs_path::emit_exists(&mut chunks[current], line);
}

/// PHP `filetype($path)` — `"dir"` for directories, else `"file"`.
pub fn emit_filetype(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stat_at(chunks, current, line);
    let chunk = &mut chunks[current];
    struct_get_key(chunk, "isDir", line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "dir", line);
    chunk.emit_else(line);
    push_str(chunk, "file", line);
    chunk.emit_end(line);
}

/// PHP `scandir($dir, ...)` — `[".", "..", <entries>...]`. Optional
/// sorting-order / context args are ignored.
pub fn emit_scandir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        for _ in 1..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    // stack: [path]
    let (dir_slot, result_slot, entries_slot, len_slot, idx_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, dir_slot, line);
    }

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, result_slot, line);
        lget(chunk, result_slot, line);
        push_str(chunk, ".", line);
    }
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, result_slot, line);
        push_str(chunk, "..", line);
    }
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, dir_slot, line);
    }
    fs_path::emit_list_dir(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, entries_slot, line);
        lget(chunk, entries_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, len_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, idx_slot, line);
    }

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, idx_slot, line);
        lget(chunk, len_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    }
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, result_slot, line);
        lget(chunk, entries_slot, line);
        lget(chunk, idx_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
    }
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
        lget(chunk, idx_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
        lset(chunk, idx_slot, line);
    }
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, result_slot, line);
    }
}

/// PHP `tempnam($dir, $prefix)` — composes `$dir/$prefix<unique>`,
/// creates the (empty) file, and returns the path.
pub fn emit_tempnam(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // stack: [dir, prefix]
    let path_slot = {
        let chunk = &mut chunks[current];
        let prefix_slot = alloc_local(chunk);
        let dir_slot = alloc_local(chunk);
        let path_slot = alloc_local(chunk);
        lset(chunk, prefix_slot, line);
        lset(chunk, dir_slot, line);
        // path = dir + "/" + prefix
        lget(chunk, dir_slot, line);
        push_str(chunk, "/", line);
        str_concat(chunk, line);
        lget(chunk, prefix_slot, line);
        str_concat(chunk, line);
        path_slot
    };
    // append a unique suffix
    crate::emitter::string_adapter::emit_php_uniqid(chunks, current, 0, line);
    {
        let chunk = &mut chunks[current];
        str_concat(chunk, line);
        lset(chunk, path_slot, line);
        // writeFile(path, "")
        lget(chunk, path_slot, line);
        push_str(chunk, "", line);
    }
    fs_path::emit_write_file(&mut chunks[current], line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line); // drop writeFile bool
        lget(chunk, path_slot, line);
    }
}

/// PHP `mime_content_type($path)` — extension-based lookup with a
/// generic fallback.
pub fn emit_mime_content_type(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // stack: [path] -> extension
    paths::emit_extension(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    let ext_slot = alloc_local(chunk);
    lset(chunk, ext_slot, line);

    let emit_case = |chunk: &mut Chunk, ext: &str, mime: &str| {
        lget(chunk, ext_slot, line);
        push_str(chunk, ext, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_str(chunk, mime, line);
        chunk.emit_else(line);
    };
    emit_case(chunk, "txt", "text/plain");
    emit_case(chunk, "html", "text/html");
    emit_case(chunk, "htm", "text/html");
    emit_case(chunk, "css", "text/css");
    emit_case(chunk, "json", "application/json");
    emit_case(chunk, "js", "application/javascript");
    emit_case(chunk, "png", "image/png");
    emit_case(chunk, "jpg", "image/jpeg");
    emit_case(chunk, "jpeg", "image/jpeg");
    emit_case(chunk, "gif", "image/gif");
    emit_case(chunk, "pdf", "application/pdf");
    // default
    push_str(chunk, "application/octet-stream", line);
    // close all the else branches
    for _ in 0..11 {
        chunk.emit_end(line);
    }
}

/// PHP `fileperms($path)` — permission bits including the file-type
/// bits (`S_IFDIR`/`S_IFREG`). Real mode bits are not exposed by the
/// fs host, so the low bits are the common `0755`/`0644` defaults.
pub fn emit_fileperms(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_stat_at(chunks, current, line);
    let chunk = &mut chunks[current];
    struct_get_key(chunk, "isDir", line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // 0o40755 = 16877 (S_IFDIR | 0755)
    push_const(chunk, Value::F64(16877.0), line);
    chunk.emit_else(line);
    // 0o100644 = 33188 (S_IFREG | 0644)
    push_const(chunk, Value::F64(33188.0), line);
    chunk.emit_end(line);
}

/// PHP `disk_free_space($dir)` / `disk_total_space($dir)` — both `false`.
///
/// These used to emit `wasi:filesystem.diskFreeSpace` / `.diskTotalSpace`.
/// Neither was registered ANYWHERE, so both were unresolved imports that trap
/// at the call — and `wasi:filesystem` is the PACKAGE name, not an importable
/// interface, so the module was wrong on top of the verb being invented.
///
/// WASI 0.3.1 has no disk-space interface at all: `filesystem/types` describes
/// descriptors, and nothing in the package reports free or total bytes for a
/// mount. There is therefore nothing honest to call.
///
/// PHP documents `false` as the return "on failure" for both, which is exactly
/// what a host that cannot answer should say — so `false` is the truthful
/// lowering rather than a stub standing in for a missing feature.
fn emit_unsupported_disk_space(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Drop the path argument(s) the caller pushed; nothing consumes them.
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_bool_const(false, line);
}

/// PHP `disk_free_space($dir)` — `false`; WASI 0.3.1 cannot answer it.
pub fn emit_disk_free_space(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_unsupported_disk_space(chunks, current, argc, line);
}

/// PHP `disk_total_space($dir)` — `false`; WASI 0.3.1 cannot answer it.
pub fn emit_disk_total_space(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_unsupported_disk_space(chunks, current, argc, line);
}

// ── php://memory stream resources ────────────────────────────────────
//
// A stream handle is an object
//   { __type:"stream", __buf:"", __pos:0, __uri:<unique>, __mode:<mode> }
// mutated in place by the f* ops. A global `ecma:map` registry keyed by
// __uri lets `filesize($uri)` recover the buffer length.

fn emit_stream_registry(chunks: &mut [Chunk], current: usize, line: u32) {
    // pushes the registry map, creating+storing it on first use.
    vybe_compiler::primitives::globals::emit_read(
        &mut chunks[current],
        "__php_stream_registry",
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_dup(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // was null: drop it, make a fresh map, store, leave map on stack
    chunk.emit_op(Op::DROP, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_dup(line);
    // GLOBAL_SET pops its value, so the dup we kept stays on the stack.
    vybe_compiler::primitives::globals::emit_write(chunk, "__php_stream_registry", line);
    chunk.emit_end(line);
}

/// PHP `fopen($path, $mode, ...)` — creates an in-memory stream handle.
pub fn emit_fopen(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        // drop optional use_include_path / context args
        for _ in 2..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    // stack: [path, mode]
    let (path_slot, mode_slot, uri_slot, stream_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, mode_slot, line);
        lset(chunk, path_slot, line);

        lget(chunk, mode_slot, line);
        push_str(chunk, "r", line);
        let starts_with = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(starts_with, 2, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        lget(chunk, path_slot, line);
        push_str(chunk, "php://", line);
        let starts_with = chunk.add_import("ecma:string", "startsWith");
        chunk.emit_call(starts_with, 2, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_op(Op::I32_AND, line);
        lget(chunk, path_slot, line);
        fs_path::emit_exists(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if(line);
        push_const(chunk, Value::Bool(false), line);
        chunk.emit_else(line);

        // uri = "php://stream/" + uniqid()
        push_str(chunk, "php://stream/", line);
    }
    crate::emitter::string_adapter::emit_php_uniqid(chunks, current, 0, line);
    {
        let chunk = &mut chunks[current];
        str_concat(chunk, line);
        lset(chunk, uri_slot, line);

        // build the stream object
        chunk.emit_struct_new(0, 0, line);
        lset(chunk, stream_slot, line);
        lget(chunk, stream_slot, line);
        push_str(chunk, "stream", line);
        struct_set_key(chunk, "__type", line);
        lget(chunk, stream_slot, line);
        push_str(chunk, "", line);
        struct_set_key(chunk, "__buf", line);
        lget(chunk, stream_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        struct_set_key(chunk, "__pos", line);
        lget(chunk, stream_slot, line);
        lget(chunk, uri_slot, line);
        struct_set_key(chunk, "__uri", line);
        lget(chunk, stream_slot, line);
        lget(chunk, mode_slot, line);
        struct_set_key(chunk, "__mode", line);
        lget(chunk, stream_slot, line);
        push_const(chunk, Value::Bool(true), line);
        struct_set_key(chunk, "__blocked", line);

        // __sink: "stdout" for php://stdout|php://output, "stderr" for
        // php://stderr, else "memory".
        let sink_slot = alloc_local(chunk);
        push_str(chunk, "memory", line);
        lset(chunk, sink_slot, line);
        // stdout schemes
        lget(chunk, path_slot, line);
        push_str(chunk, "php://output", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        lget(chunk, path_slot, line);
        push_str(chunk, "php://stdout", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_OR, line);
        chunk.emit_if(line);
        push_str(chunk, "stdout", line);
        lset(chunk, sink_slot, line);
        chunk.emit_end(line);
        // stderr scheme
        lget(chunk, path_slot, line);
        push_str(chunk, "php://stderr", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        push_str(chunk, "stderr", line);
        lset(chunk, sink_slot, line);
        chunk.emit_end(line);
        lget(chunk, stream_slot, line);
        lget(chunk, sink_slot, line);
        struct_set_key(chunk, "__sink", line);
    }

    // registry.set(uri, stream)
    emit_stream_registry(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, uri_slot, line);
        lget(chunk, stream_slot, line);
    }
    call_import(chunks, current, "ecma:map", "set", 3, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line); // drop map returned by set
        lget(chunk, stream_slot, line);
        chunk.emit_end(line);
    }
}

/// PHP `fwrite($stream, $data, ...)` — for a `stdout`/`stderr` sink writes
/// to the process stream; for a memory stream appends to the buffer.
/// Returns the byte count written.
/// PHP `fputcsv($stream, $fields, $separator = ",", $enclosure = "\"")`.
///
/// Composed, not reimplemented: the row is rendered by the SHARED
/// `primitives/csv.rs::emit_format_row` — the same emitter fortran and any
/// other CSV consumer reach — and then handed to `fwrite`. php terminates the
/// record with `\n`.
pub fn emit_fputcsv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 5 {
        chunk.emit_op(Op::DROP, line); // $escape — doubling is the scanner's job
    }
    let enc_slot = alloc_local(chunk);
    let sep_slot = alloc_local(chunk);
    let rows_slot = alloc_local(chunk);
    let stream_slot = alloc_local(chunk);
    if argc >= 4 {
        lset(chunk, enc_slot, line);
    } else {
        push_str(chunk, "\"", line);
        lset(chunk, enc_slot, line);
    }
    if argc >= 3 {
        lset(chunk, sep_slot, line);
    } else {
        push_str(chunk, ",", line);
        lset(chunk, sep_slot, line);
    }
    lset(chunk, rows_slot, line);
    lset(chunk, stream_slot, line);

    lget(chunk, stream_slot, line);
    lget(chunk, rows_slot, line);
    lget(chunk, sep_slot, line);
    lget(chunk, enc_slot, line);
    vybe_compiler::primitives::csv::emit_format_row(
        chunks,
        current,
        vybe_compiler::primitives::csv::FormatOptions::minimal(),
        line,
    );
    let chunk = &mut chunks[current];
    push_str(chunk, "\n", line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    // stack: [stream, line] — exactly `fwrite`'s contract.
    emit_fwrite(chunks, current, 2, line);
}

/// PHP `fgetcsv($stream, $length = null, $separator = ",", $enclosure = "\"")`.
///
/// `fgets` then the SHARED scanner. The trailing newline `fgets` keeps is
/// stripped first, or it would land in the last field.
pub fn emit_fgetcsv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 5 {
        chunk.emit_op(Op::DROP, line); // $escape
    }
    let enc_slot = alloc_local(chunk);
    let sep_slot = alloc_local(chunk);
    if argc >= 4 {
        lset(chunk, enc_slot, line);
    } else {
        push_str(chunk, "\"", line);
        lset(chunk, enc_slot, line);
    }
    if argc >= 3 {
        lset(chunk, sep_slot, line);
    } else {
        push_str(chunk, ",", line);
        lset(chunk, sep_slot, line);
    }
    if argc >= 2 {
        chunk.emit_op(Op::DROP, line); // $length — `fgets` reads a whole line
    }

    // stack: [stream] → fgets → [record]
    emit_read_record(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, sep_slot, line);
    lget(chunk, enc_slot, line);
    vybe_compiler::primitives::csv::emit_parse_line(chunks, current, line);
}

/// Read one record: `fgets($stream)` with the trailing `\r\n` / `\n` removed.
/// Left on, the terminator would become part of the final field.
fn emit_read_record(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_fgets(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    {
        let idx = chunk.add_import("ecma:string", "trimEnd");
        chunk.emit_call(idx, 1, line);
    }
}

pub fn emit_fwrite(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 3 {
        chunk.emit_op(Op::DROP, line); // drop optional length
    }
    // stack: [stream, data]
    let data_slot = alloc_local(chunk);
    let stream_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let sink_slot = alloc_local(chunk);
    lset(chunk, data_slot, line);
    lset(chunk, stream_slot, line);

    lget(chunk, data_slot, line);
    str_len(chunk, line);
    lset(chunk, len_slot, line);

    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__sink", line);
    lset(chunk, sink_slot, line);

    // if sink == "stdout"
    lget(chunk, sink_slot, line);
    push_str(chunk, "stdout", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // The shared stdout write. `data_slot` already holds the string, which is
    // exactly what `emit_write_stdout_slot` takes — the canon
    // `stream.new`/`stream.write`/`drop` sequence used to be spliced here by
    // hand.
    vybe_compiler::primitives::io::emit_write_stdout_slot(chunk, data_slot, line);
    chunk.emit_else(line);
    // if sink == "stderr" -> discard (no stdout pollution); else memory buffer
    lget(chunk, sink_slot, line);
    push_str(chunk, "stderr", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // stderr: nothing to emit
    chunk.emit_else(line);
    // memory: buf = buf + data ; pos = buf.length
    lget(chunk, stream_slot, line);
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    lget(chunk, data_slot, line);
    str_concat(chunk, line);
    struct_set_key(chunk, "__buf", line);
    lget(chunk, stream_slot, line);
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    str_len(chunk, line);
    struct_set_key(chunk, "__pos", line);
    chunk.emit_end(line); // end stderr/memory if
    chunk.emit_end(line); // end stdout if

    lget(chunk, len_slot, line);
}

/// PHP `fread($stream, $length)` — reads up to `$length` bytes from the
/// current position and advances it.
pub fn emit_fread(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // stack: [stream, length]
    let chunk = &mut chunks[current];
    let length_slot = alloc_local(chunk);
    let stream_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);
    lset(chunk, length_slot, line);
    lset(chunk, stream_slot, line);

    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__pos", line);
    lset(chunk, pos_slot, line);

    // end = pos + length
    lget(chunk, pos_slot, line);
    lget(chunk, length_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, end_slot, line);

    // result = buf.substring(pos, end)
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    lget(chunk, pos_slot, line);
    lget(chunk, end_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    // pos = min(end, buf.length)
    lget(chunk, stream_slot, line);
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    str_len(chunk, line);
    let buflen_slot = alloc_local(chunk);
    lset(chunk, buflen_slot, line);
    lget(chunk, end_slot, line);
    lget(chunk, buflen_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, end_slot, line);
    chunk.emit_else(line);
    lget(chunk, buflen_slot, line);
    chunk.emit_end(line);
    struct_set_key(chunk, "__pos", line);

    lget(chunk, result_slot, line);
}

/// PHP `fgets($stream, ...)` — reads a line (through the next `\n`,
/// inclusive) from the current position.
pub fn emit_fgets(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        if argc >= 2 {
            chunk.emit_op(Op::DROP, line); // drop optional length
        }
    }
    // stack: [stream]
    let chunk = &mut chunks[current];
    let stream_slot = alloc_local(chunk);
    let buf_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    let nl_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);
    lset(chunk, stream_slot, line);

    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    lset(chunk, buf_slot, line);
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__pos", line);
    lset(chunk, pos_slot, line);

    // nl = buf.indexOf("\n", pos)
    lget(chunk, buf_slot, line);
    push_str(chunk, "\n", line);
    lget(chunk, pos_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "indexOf");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, nl_slot, line);

    // end = nl < 0 ? buf.length : nl + 1
    lget(chunk, nl_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, buf_slot, line);
    str_len(chunk, line);
    chunk.emit_else(line);
    lget(chunk, nl_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_end(line);
    lset(chunk, end_slot, line);

    // result = buf.substring(pos, end)
    lget(chunk, buf_slot, line);
    lget(chunk, pos_slot, line);
    lget(chunk, end_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    // pos = end
    lget(chunk, stream_slot, line);
    lget(chunk, end_slot, line);
    struct_set_key(chunk, "__pos", line);

    lget(chunk, result_slot, line);
}

/// PHP `fgetc($stream)` — reads a single character and advances.
pub fn emit_fgetc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let stream_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    lset(chunk, stream_slot, line);

    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__pos", line);
    lset(chunk, pos_slot, line);

    // result = buf.charAt(pos)
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    lget(chunk, pos_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "charAt");
        chunk.emit_call(idx, 2, line);
    }
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    // pos = pos + 1
    lget(chunk, stream_slot, line);
    lget(chunk, pos_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    struct_set_key(chunk, "__pos", line);

    lget(chunk, result_slot, line);
}

/// PHP `feof($stream)` — true when the cursor is at/after buffer end.
pub fn emit_feof(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let stream_slot = alloc_local(chunk);
    lset(chunk, stream_slot, line);

    // pos >= buf.length
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    str_len(chunk, line);
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__pos", line);
    vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

/// PHP `ftell($stream)` — current cursor position.
pub fn emit_ftell(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    struct_get_key(chunk, "__pos", line);
}

/// PHP `fseek($stream, $offset, $whence = SEEK_SET)` — sets the cursor.
/// SEEK_SET(0)=offset, SEEK_CUR(1)=pos+offset, SEEK_END(2)=len+offset.
pub fn emit_fseek(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let whence_slot = alloc_local(chunk);
    let offset_slot = alloc_local(chunk);
    let stream_slot = alloc_local(chunk);
    let target_slot = alloc_local(chunk);
    if argc >= 3 {
        lset(chunk, whence_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, whence_slot, line);
    }
    lset(chunk, offset_slot, line);
    lset(chunk, stream_slot, line);

    // target = offset (SEEK_SET default)
    lget(chunk, offset_slot, line);
    lset(chunk, target_slot, line);

    // if whence == SEEK_CUR: target = pos + offset
    lget(chunk, whence_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__pos", line);
    lget(chunk, offset_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, target_slot, line);
    chunk.emit_end(line);

    // if whence == SEEK_END: target = buf.length + offset
    lget(chunk, whence_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    str_len(chunk, line);
    lget(chunk, offset_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, target_slot, line);
    chunk.emit_end(line);

    // stream.__pos = target
    lget(chunk, stream_slot, line);
    lget(chunk, target_slot, line);
    struct_set_key(chunk, "__pos", line);

    // return 0 (success)
    push_const(chunk, Value::F64(0.0), line);
}

/// PHP `rewind($stream)` — sets the cursor to 0, returns true.
pub fn emit_rewind(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let stream_slot = alloc_local(chunk);
    lset(chunk, stream_slot, line);
    lget(chunk, stream_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    struct_set_key(chunk, "__pos", line);
    push_const(chunk, Value::Bool(true), line);
}

/// PHP `fflush($stream)` — no-op on memory streams, returns true.
pub fn emit_fflush(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(true), line);
}

/// PHP `fclose($stream)` — drops the handle, returns true.
pub fn emit_fclose(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(true), line);
}

/// PHP `stream_get_contents($stream, ...)` — returns the buffer from the
/// current position to the end and advances the cursor.
pub fn emit_stream_get_contents(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        // drop optional maxlength / offset args
        for _ in 1..argc {
            chunk.emit_op(Op::DROP, line);
        }
    }
    let chunk = &mut chunks[current];
    let stream_slot = alloc_local(chunk);
    let pos_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);
    lset(chunk, stream_slot, line);

    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__pos", line);
    lset(chunk, pos_slot, line);
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    str_len(chunk, line);
    lset(chunk, end_slot, line);

    // result = buf.substring(pos, end)
    lget(chunk, stream_slot, line);
    struct_get_key(chunk, "__buf", line);
    lget(chunk, pos_slot, line);
    lget(chunk, end_slot, line);
    {
        let idx = chunk.add_import("ecma:string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    let result_slot = alloc_local(chunk);
    lset(chunk, result_slot, line);

    // pos = end
    lget(chunk, stream_slot, line);
    lget(chunk, end_slot, line);
    struct_set_key(chunk, "__pos", line);

    lget(chunk, result_slot, line);
}

/// PHP `stream_get_meta_data($stream)` — a metadata record. `uri` is the
/// handle's synthetic uri (usable with `filesize`).
pub fn emit_stream_get_meta_data(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // stack: [stream]
    let (stream_slot, meta_slot) = {
        let chunk = &mut chunks[current];
        (alloc_local(chunk), alloc_local(chunk))
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, stream_slot, line);
    }
    // PHP `stream_get_meta_data` returns an *array* — build a Map so
    // `['uri']` subscripting works uniformly (structs miss direct
    // call-result subscript).
    call_import(chunks, current, "ecma:map", "new", 0, line);
    {
        let chunk = &mut chunks[current];
        lset(chunk, meta_slot, line);
    }

    // meta.set("uri", stream.__uri)
    {
        let chunk = &mut chunks[current];
        lget(chunk, meta_slot, line);
        push_str(chunk, "uri", line);
        lget(chunk, stream_slot, line);
        struct_get_key(chunk, "__uri", line);
    }
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    // meta.set("mode", stream.__mode)
    {
        let chunk = &mut chunks[current];
        lget(chunk, meta_slot, line);
        push_str(chunk, "mode", line);
        lget(chunk, stream_slot, line);
        struct_get_key(chunk, "__mode", line);
    }
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    // scalar/bool metadata entries
    for (key, val) in [
        ("stream_type", Value::String(Arc::from("MEMORY"))),
        ("wrapper_type", Value::String(Arc::from("PHP"))),
        ("seekable", Value::Bool(true)),
        ("unread_bytes", Value::F64(0.0)),
        ("timed_out", Value::Bool(false)),
        ("eof", Value::Bool(false)),
    ] {
        {
            let chunk = &mut chunks[current];
            lget(chunk, meta_slot, line);
            push_str(chunk, key, line);
            push_const(chunk, val, line);
        }
        call_import(chunks, current, "ecma:map", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    {
        let chunk = &mut chunks[current];
        lget(chunk, meta_slot, line);
        push_str(chunk, "blocked", line);
        lget(chunk, stream_slot, line);
        struct_get_key(chunk, "__blocked", line);
    }
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, meta_slot, line);
    }
}
