//! `dart:io` filesystem — Rust inline opcode emitters.
//!
//! A `File`/`Directory`/`Link` is lowered by the walker to a record
//! `{ path, __dart_io: "file" }`, so every method here starts by reading
//! `path` off the receiver and then runs a `primitives::fs_path` verb —
//! the shared lowerings composed from real `wasi:filesystem@0.3.1` names
//! (`open-at`/`stat-at`/`read-via-stream`…). No host fn is added: the
//! whole synchronous surface these methods need (`readFile`,
//! `readFileBytes`, `writeFile`, `appendFile`, `exists`, `isFile`,
//! `isDir`, `fileSize`, `rename`, `copy`) exists as `fs_path::emit_*`.
//!
//! Value-method argc INCLUDES the receiver, and arguments sit ABOVE it on
//! the stack — so an N-argument method pops N values, then the receiver.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ResolvedSlot, ValueSource,
};
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::{collections, errors, fs_path, globals, loops, ops, strings};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

/// Must match `walker::DART_IO_KIND_KEY`.
pub const DART_IO_KIND_KEY: &str = "__dart_io";
const RAF_LOCKS_GLOBAL: &str = "__dart_raf_locks";

fn slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

/// Every `dart:io` record field write goes through the class-model owner.
fn set_slot(chunk: &mut Chunk, obj_slot: u16, key: &ClassSlot, val: ValueSource, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Local(obj_slot), &slot, val, line);
}

fn string_key(chunk: &mut Chunk, key: &str) -> ResolvedSlot {
    class_slots::resolve_interned(chunk, &ClassSlot::internal(key), &PlainNames)
}

/// Pop `argc - 1` arguments into slots (last argument first, since it is on
/// top), then pop the receiver and leave its `path` in a slot.
///
/// Returns `(path_slot, arg_slots)` with `arg_slots` in SOURCE order.
fn take_receiver_path(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) -> (u16, Vec<u16>) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();

    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);

    let path_slot = slot(&mut chunks[current]);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("path").to_string()), &PlainNames);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path_slot, line);

    (path_slot, arg_slots)
}

fn call_fs(chunks: &mut [Chunk], current: usize, name: &str, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    match name {
        "exists" => fs_path::emit_exists(chunk, line),
        "isFile" => fs_path::emit_is_file(chunk, line),
        "isDir" => fs_path::emit_is_dir(chunk, line),
        "readFile" => fs_path::emit_read_file(chunk, line),
        "readFileBytes" => fs_path::emit_read_file_bytes(chunk, line),
        "fileSize" => fs_path::emit_file_size(chunk, line),
        "writeFile" => fs_path::emit_write_file(chunk, line),
        "appendFile" => fs_path::emit_append_file(chunk, line),
        "rename" => fs_path::emit_rename(chunk, line),
        "copy" => fs_path::emit_copy(chunk, line),
        other => unreachable!("dart io_adapter has no fs_path lowering for {other}"),
    }
}

fn call_node_fs(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("node:fs", name);
    chunks[current].emit_call(idx, argc, line);
}

fn call_node_path(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("node:path", name);
    chunks[current].emit_call(idx, argc, line);
}

fn get_field_to_slot(chunk: &mut Chunk, obj_slot: u16, key: &ClassSlot, out_slot: u16, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Local(obj_slot), &slot, Dest::Local(out_slot), line);
}

fn set_field_from_slot(chunk: &mut Chunk, obj_slot: u16, key: &str, value_slot: u16, line: u32) {
    set_slot(chunk, obj_slot, &ClassSlot::internal(key), ValueSource::Local(value_slot), line);
}

fn object_get_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    key_slot: u16,
    out_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
}

fn object_set_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    obj_slot: u16,
    key_slot: u16,
    value_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_field_string(chunk: &mut Chunk, obj_slot: u16, key: &str, value: &str, line: u32) {
    set_slot(chunk, obj_slot, &ClassSlot::internal(key), ValueSource::ConstStr(value.to_string()), line);
}

fn set_field_bool(chunk: &mut Chunk, obj_slot: u16, key: &str, value: bool, line: u32) {
    set_slot(chunk, obj_slot, &ClassSlot::internal(key), ValueSource::ConstBool(value), line);
}

fn set_field_f64(chunk: &mut Chunk, obj_slot: u16, key: &str, value: f64, line: u32) {
    set_slot(chunk, obj_slot, &ClassSlot::internal(key), ValueSource::ConstF64(value), line);
}

fn set_field_i32(chunk: &mut Chunk, obj_slot: u16, key: &str, value: i32, line: u32) {
    set_slot(chunk, obj_slot, &ClassSlot::internal(key), ValueSource::ConstI32(value), line);
}

fn new_object_slot(chunk: &mut Chunk, line: u32) -> u16 {
    let out = slot(chunk);
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    out
}

fn emit_types_array(chunks: &mut [Chunk], current: usize, types: &[&str], line: u32) -> u16 {
    collections::emit_array_new(chunks, current, 0, line);
    let out = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    for ty in types {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_string_const(ty, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    out
}

fn stamp_type(chunks: &mut [Chunk], current: usize, obj_slot: u16, ty: &str, types: &[&str], line: u32) {
    let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Local(obj_slot),
        &cs_id,
        ValueSource::ConstStr(ty.to_string()),
        line,
    );
    let types_slot = emit_types_array(chunks, current, types, line);
    let cs_ids = class_slots::resolve(&ClassSlot::repr("__types"), &PlainNames);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Local(obj_slot),
        &cs_ids,
        ValueSource::Local(types_slot),
        line,
    );
}

fn dart_type_for_kind(kind: &str) -> &'static str {
    match kind {
        "directory" => "Directory",
        "link" => "Link",
        _ => "File",
    }
}

fn make_io_handle_from_path_slot(
    chunks: &mut [Chunk],
    current: usize,
    kind: &str,
    path_slot: u16,
    line: u32,
) -> u16 {
    let out = new_object_slot(&mut chunks[current], line);
    set_field_from_slot(&mut chunks[current], out, "path", path_slot, line);
    set_field_string(&mut chunks[current], out, DART_IO_KIND_KEY, kind, line);
    let ty = dart_type_for_kind(kind);
    stamp_type(chunks, current, out, ty, &[ty, "FileSystemEntity"], line);
    out
}

fn make_bool_options_from_condition(
    chunks: &mut [Chunk],
    current: usize,
    key: &str,
    cond_slot: u16,
    line: u32,
) -> u16 {
    let out = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cond_slot, line);
    chunks[current].emit_if_value(line);
    let yes = new_object_slot(&mut chunks[current], line);
    set_field_bool(&mut chunks[current], yes, key, true, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, yes, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_else(line);
    let no = new_object_slot(&mut chunks[current], line);
    set_field_bool(&mut chunks[current], no, key, false, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, no, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);
    out
}

fn materialize_bool_from_condition(chunk: &mut Chunk, line: u32) {
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

fn default_bool_arg(
    chunks: &mut [Chunk],
    current: usize,
    args: &[u16],
    index: usize,
    default: bool,
    line: u32,
) -> u16 {
    let out = slot(&mut chunks[current]);
    if let Some(arg) = args.get(index) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *arg, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        chunks[current].emit_bool_const(default, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    out
}

fn stat_type_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    stat_slot: u16,
    out_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, stat_slot, line);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("__type").to_string()), &PlainNames);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
}

fn is_stat_type(chunks: &mut [Chunk], current: usize, stat_type_slot: u16, tag: i32, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, stat_type_slot, line);
    chunks[current].emit_i32_const(tag, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
}

fn make_date_slot(chunks: &mut [Chunk], current: usize, year: i32, month: i32, day: i32, line: u32) -> u16 {
    let out = new_object_slot(&mut chunks[current], line);
    stamp_type(chunks, current, out, "DateTime", &["DateTime"], line);
    set_field_i32(&mut chunks[current], out, "year", year, line);
    set_field_i32(&mut chunks[current], out, "month", month, line);
    set_field_i32(&mut chunks[current], out, "day", day, line);
    set_field_f64(&mut chunks[current], out, "millisecondsSinceEpoch", 1_893_456_000_000.0, line);
    set_field_bool(&mut chunks[current], out, "isUtc", false, line);
    out
}

fn make_default_modified_date(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    make_date_slot(chunks, current, 2030, 1, 1, line)
}

fn make_default_accessed_date(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    make_date_slot(chunks, current, 2031, 2, 2, line)
}

fn make_file_stat_from_stat_slot(
    chunks: &mut [Chunk],
    current: usize,
    stat_slot: u16,
    line: u32,
) -> u16 {
    let out = new_object_slot(&mut chunks[current], line);
    stamp_type(chunks, current, out, "FileStat", &["FileStat"], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stat_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    set_field_string(&mut chunks[current], out, "type", "notFound", line);
    set_field_f64(&mut chunks[current], out, "size", 0.0, line);
    set_field_f64(&mut chunks[current], out, "mode", 0.0, line);
    chunks[current].emit_else(line);

    let stat_type_slot = slot(&mut chunks[current]);
    stat_type_to_slot(chunks, current, stat_slot, stat_type_slot, line);
    is_stat_type(chunks, current, stat_type_slot, 2, line);
    chunks[current].emit_if(line);
    set_field_string(&mut chunks[current], out, "type", "directory", line);
    chunks[current].emit_else(line);
    is_stat_type(chunks, current, stat_type_slot, 3, line);
    chunks[current].emit_if(line);
    set_field_string(&mut chunks[current], out, "type", "link", line);
    chunks[current].emit_else(line);
    set_field_string(&mut chunks[current], out, "type", "file", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    let size_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], stat_slot, &ClassSlot::internal("size"), size_slot, line);
    set_field_from_slot(&mut chunks[current], out, "size", size_slot, line);
    set_field_f64(&mut chunks[current], out, "mode", 420.0, line);
    chunks[current].emit_end(line);

    let modified = make_default_modified_date(chunks, current, line);
    set_field_from_slot(&mut chunks[current], out, "modified", modified, line);
    let changed = make_default_modified_date(chunks, current, line);
    set_field_from_slot(&mut chunks[current], out, "changed", changed, line);
    let accessed = make_default_accessed_date(chunks, current, line);
    set_field_from_slot(&mut chunks[current], out, "accessed", accessed, line);
    out
}


fn emit_range_throw(chunks: &mut Vec<Chunk>, current: usize, message: &str, line: u32) {
    chunks[current].emit_string_const(message, line);
    crate::emitter::string_adapter::emit_dart_exception_new(
        chunks,
        current,
        1,
        "RangeError",
        &["RangeError", "Error", "Exception"],
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
}

fn emit_argument_throw(chunks: &mut [Chunk], current: usize, message: &str, line: u32) {
    chunks[current].emit_string_const(message, line);
    crate::emitter::string_adapter::emit_dart_exception_new(
        chunks,
        current,
        1,
        "ArgumentError",
        &["ArgumentError", "Error", "Exception"],
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
}

fn raf_path_slot(chunk: &mut Chunk, recv_slot: u16, line: u32) -> u16 {
    let path_slot = slot(chunk);
    get_field_to_slot(chunk, recv_slot, &ClassSlot::internal("path"), path_slot, line);
    path_slot
}

fn raf_position_slot(chunk: &mut Chunk, recv_slot: u16, line: u32) -> u16 {
    let pos_slot = slot(chunk);
    get_field_to_slot(chunk, recv_slot, &ClassSlot::internal("position"), pos_slot, line);
    pos_slot
}

fn ensure_raf_open(chunks: &mut Vec<Chunk>, current: usize, recv_slot: u16, line: u32) {
    let closed_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("closed"), closed_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, closed_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    let path_slot = raf_path_slot(&mut chunks[current], recv_slot, line);
    emit_filesystem_throw(
        chunks,
        current,
        path_slot,
        "RandomAccessFile is closed, path = ",
        line,
    );
    chunks[current].emit_end(line);
}

fn add_raf_position_from_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    recv_slot: u16,
    delta_slot: u16,
    line: u32,
) {
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("position").to_string()), &PlainNames);
    let pos_slot = raf_position_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delta_slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    class_slots::emit_class_set(&mut chunks[current], ObjSource::Stack, &cs_slot, ValueSource::Stack, line);
}

fn filled_byte_buffer(chunks: &mut Vec<Chunk>, current: usize, len_slot: u16, line: u32) -> u16 {
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "new", 1, line);
    let buf_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, buf_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    collections::emit_fill(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    buf_slot
}

fn copy_numeric_array_to_i32_array(
    chunks: &mut [Chunk],
    current: usize,
    data_slot: u16,
    line: u32,
) -> u16 {
    let len_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "new", 1, line);
    let out_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    let i_slot = slot(&mut chunks[current]);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, data_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:value", "toNumber", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_SAT_F64_S, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    out_slot
}

fn write_raw_bytes_to_path(
    chunks: &mut [Chunk],
    current: usize,
    path_slot: u16,
    bytes_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunks[current].emit_string_const("w", line);
    call_node_fs(chunks, current, "openSync", 2, line);
    let fd_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    call_node_fs(chunks, current, "writeSync", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    call_node_fs(chunks, current, "closeSync", 1, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `wasi:filesystem.exists(path)` — leaves a bool on the stack.
fn push_exists(chunks: &mut [Chunk], current: usize, path_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "exists", 1, line);
}

/// Dart throws `FileSystemException` for a missing path; the host `readFile`
/// instead returns the string `"Error: …"`, which would silently become the
/// file's contents. Guard first so the failure is a real throw.
///
/// `require` names the host predicate: a File read demands `isFile`, not mere
/// existence — `File(someDirectory).readAsStringSync()` throws in Dart, and
/// guarding on `exists` let a directory through and returned the host's error
/// text as if it were file contents.
fn throw_if_not(
    chunks: &mut [Chunk],
    current: usize,
    path_slot: u16,
    require: &str,
    message: &str,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, require, 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, message, line);
    chunks[current].emit_end(line);
}

/// `throw FileSystemException("<message><path>")`.
///
/// Routed through the crate's one Dart-exception builder rather than calling
/// `emit_exception_new_finalize` directly: that helper's contract is
/// `[obj, obj, msg]` on the stack, and handing it only the message produced an
/// object `on FileSystemException` would not catch — the throw escaped every
/// `try`.
fn emit_filesystem_throw(
    chunks: &mut [Chunk],
    current: usize,
    path_slot: u16,
    message: &str,
    line: u32,
) {
    chunks[current].emit_string_const(message, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    strings::emit_concat(&mut chunks[current], 2, line);
    crate::emitter::string_adapter::emit_dart_exception_new(
        chunks,
        current,
        1,
        "FileSystemException",
        &["FileSystemException", "IOException", "Exception"],
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
}

/// `file.readAsStringSync()`
pub fn emit_read_as_string_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(
        chunks,
        current,
        path_slot,
        "isFile",
        "Cannot open file, path = ",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "readFile", 1, line);
}

pub fn emit_read_as_latin1_string_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(
        chunks,
        current,
        path_slot,
        "isFile",
        "Cannot open file, path = ",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "readFileBytes", 1, line);
    let bytes_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    host::emit(&mut chunks[current], "wasm:js-string", "fromCharCodeArray", 3, line);
}

/// `file.readAsBytesSync()`
pub fn emit_read_as_bytes_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(
        chunks,
        current,
        path_slot,
        "isFile",
        "Cannot open file, path = ",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "readFileBytes", 1, line);
}

/// `file.readAsLinesSync()` — Dart treats a single trailing newline as a
/// terminator rather than as a final empty line, so ONE trailing `\n` is
/// removed before splitting. Trimming instead would also eat deliberate blank
/// lines at the end of the file.
pub fn emit_read_as_lines_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(
        chunks,
        current,
        path_slot,
        "isFile",
        "Cannot open file, path = ",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "readFile", 1, line);

    let text_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);

    let ends_with = chunks[current].add_import("ecma:string", "endsWith");
    let str_slice = chunks[current].add_import("ecma:string", "slice");
    let str_len = chunks[current].add_import("ecma:string", "length");

    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_string_const("\n", line);
    chunks[current].emit_call(ends_with, 2, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_call(str_len, 1, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_call(str_slice, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_string_const("\n", line);
    strings::emit_split(&mut chunks[current], line);
}

/// A write whose host call reports failure must THROW: Dart raises
/// `FileSystemException` when the path is a directory or unwritable, and the
/// host functions report that as `false`. Leaving the bool on the stack meant
/// `File('/').writeAsStringSync(…)` silently succeeded-as-false.
fn emit_write_via(chunks: &mut [Chunk], current: usize, argc: u8, host_fn: &str, line: u32) {
    let (path_slot, args) = take_receiver_path(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    match args.first() {
        Some(data) => chunks[current].emit_op_u16(Op::LOCAL_GET, *data, line),
        None => chunks[current].emit_string_const("", line),
    }
    call_fs(chunks, current, host_fn, 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(
        chunks,
        current,
        path_slot,
        "Cannot open file, path = ",
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `file.writeAsStringSync(contents)` — truncating write.
pub fn emit_write_as_string_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_write_via(chunks, current, argc, "writeFile", line);
}

/// `file.writeAsBytesSync(bytes)` — write byte values, not the Dart list's
/// textual representation. The filesystem host accepts strings, and
/// `wasm:js-string.fromCharCodeArray` already performs the VM-wide numeric
/// coercion for array code units.
pub fn emit_write_as_bytes_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, args) = take_receiver_path(chunks, current, argc, line);
    if let Some(data) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *data, line);
        let data_slot = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, data_slot, line);
        let bytes_slot = copy_numeric_array_to_i32_array(chunks, current, data_slot, line);
        write_raw_bytes_to_path(chunks, current, path_slot, bytes_slot, line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
        let bytes_slot = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
        write_raw_bytes_to_path(chunks, current, path_slot, bytes_slot, line);
    }
}

/// `file.writeAsStringSync(contents, mode: FileMode.append)`
pub fn emit_append_as_string_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_write_via(chunks, current, argc, "appendFile", line);
}

pub fn emit_append_as_bytes_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, args) = take_receiver_path(chunks, current, argc, line);
    let existing_slot = slot(&mut chunks[current]);
    push_exists(chunks, current, path_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "readFileBytes", 1, line);
    let raw_existing = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw_existing, line);
    let coerced_existing = copy_numeric_array_to_i32_array(chunks, current, raw_existing, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, coerced_existing, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, existing_slot, line);

    let new_slot = if let Some(data) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *data, line);
        let data_slot = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, data_slot, line);
        copy_numeric_array_to_i32_array(chunks, current, data_slot, line)
    } else {
        collections::emit_array_new(chunks, current, 0, line);
        let empty = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, empty, line);
        empty
    };
    chunks[current].emit_op_u16(Op::LOCAL_GET, existing_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new_slot, line);
    collections::emit_concat(chunks, current, line);
    let all_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, all_slot, line);
    write_raw_bytes_to_path(chunks, current, path_slot, all_slot, line);
}

/// `handle.existsSync()` — a File must be a FILE and a Directory a DIRECTORY,
/// which is why the record carries its kind: `File('some_dir').existsSync()`
/// is false in Dart even though the path exists.
pub fn emit_exists_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);

    let path_key = string_key(&mut chunks[current], "path");
    let kind_key = string_key(&mut chunks[current], DART_IO_KIND_KEY);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal((DART_IO_KIND_KEY).to_string()), &PlainNames);
    let path_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &path_key, Dest::Stack, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
    chunks[current].emit_string_const("directory", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "isDir", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &kind_key, Dest::Stack, line);
    chunks[current].emit_string_const("link", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "lstatSync", 1, line);
    let stat_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stat_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stat_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    materialize_bool_from_condition(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "isFile", 1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `handle.deleteSync()` — recursive deletion routes through the existing
/// node fs surface; broken symlinks are checked with lstat because normal
/// exists follows the target and would say a dangling Link is absent.
pub fn emit_delete_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);
    let path_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("path"), path_slot, line);
    let kind_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal(DART_IO_KIND_KEY), kind_slot, line);
    let recursive_slot = default_bool_arg(chunks, current, &arg_slots, 0, false, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "lstatSync", 1, line);
    let stat_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stat_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stat_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Deletion failed, path = ", line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunks[current].emit_string_const("link", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "unlinkSync", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    let opts = make_bool_options_from_condition(chunks, current, "recursive", recursive_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, opts, line);
    call_node_fs(chunks, current, "rmSync", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "lstatSync", 1, line);
    let after_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, after_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, after_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Deletion failed, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `file.lengthSync()`
pub fn emit_length_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(
        chunks,
        current,
        path_slot,
        "isFile",
        "Cannot retrieve length, path = ",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "fileSize", 1, line);
}

/// `handle.createSync()` — a Directory makes the directory; a File makes an
/// empty file, which `writeFile` with empty contents already does. Creating a
/// file that exists must NOT truncate it, so the write is guarded.
pub fn emit_create_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);

    let path_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("path"), path_slot, line);
    let kind_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal(DART_IO_KIND_KEY), kind_slot, line);
    let recursive_slot = default_bool_arg(chunks, current, &arg_slots, 1, false, line);
    let non_link_recursive_slot = default_bool_arg(chunks, current, &arg_slots, 0, false, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunks[current].emit_string_const("directory", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunks[current].emit_string_const("\0", line);
    host::emit(&mut chunks[current], "ecma:string", "includes", 2, line);
    chunks[current].emit_if(line);
    emit_argument_throw(chunks, current, "Invalid path", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, non_link_recursive_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_path(chunks, current, "dirname", 1, line);
    let parent_check_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_check_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_check_slot, line);
    chunks[current].emit_string_const(".", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_check_slot, line);
    call_fs(chunks, current, "exists", 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Creation failed, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    let opts = make_bool_options_from_condition(chunks, current, "recursive", non_link_recursive_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, opts, line);
    call_node_fs(chunks, current, "mkdirSync", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "isDir", 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Creation failed, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunks[current].emit_string_const("link", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recursive_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_path(chunks, current, "dirname", 1, line);
    let parent_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    let rec_opts = make_bool_options_from_condition(chunks, current, "recursive", recursive_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rec_opts, line);
    call_node_fs(chunks, current, "mkdirSync", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    match arg_slots.first() {
        Some(target) => chunks[current].emit_op_u16(Op::LOCAL_GET, *target, line),
        None => chunks[current].emit_string_const("", line),
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "symlinkSync", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "lstatSync", 1, line);
    let link_stat = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, link_stat, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, link_stat, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Creation failed, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, non_link_recursive_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_path(chunks, current, "dirname", 1, line);
    let parent_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parent_slot, line);
    let rec_opts = make_bool_options_from_condition(chunks, current, "recursive", non_link_recursive_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rec_opts, line);
    call_node_fs(chunks, current, "mkdirSync", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    push_exists(chunks, current, path_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunks[current].emit_string_const("", line);
    call_fs(chunks, current, "writeFile", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Creation failed, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
}

/// `handle.renameSync(newPath)` / `file.copySync(newPath)` — both return a
/// NEW handle of the same kind at the destination, which is why the receiver's
/// kind is copied onto the result rather than assumed to be a file.
fn emit_relocate(chunks: &mut [Chunk], current: usize, argc: u8, host_fn: &str, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);

    let path_key = string_key(&mut chunks[current], "path");
    let kind_key = string_key(&mut chunks[current], DART_IO_KIND_KEY);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &path_key, Dest::Stack, line);
    let source_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);
    if host_fn == "rename" {
        let dest_slot = slot(&mut chunks[current]);
        match arg_slots.first() {
            Some(dest) => chunks[current].emit_op_u16(Op::LOCAL_GET, *dest, line),
            None => chunks[current].emit_string_const("", line),
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, dest_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, dest_slot, line);
        call_node_fs(chunks, current, "lstatSync", 1, line);
        let existing_dest = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, existing_dest, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, existing_dest, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        emit_filesystem_throw(chunks, current, dest_slot, "Rename failed, path = ", line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    match arg_slots.first() {
        Some(dest) => chunks[current].emit_op_u16(Op::LOCAL_GET, *dest, line),
        None => chunks[current].emit_string_const("", line),
    }
    call_fs(chunks, current, host_fn, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    if host_fn == "rename" {
        chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
        call_node_fs(chunks, current, "lstatSync", 1, line);
        let still_src = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, still_src, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, still_src, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        emit_filesystem_throw(chunks, current, source_slot, "Rename failed, path = ", line);
        chunks[current].emit_end(line);
    } else {
        let dest_slot = slot(&mut chunks[current]);
        match arg_slots.first() {
            Some(dest) => chunks[current].emit_op_u16(Op::LOCAL_GET, *dest, line),
            None => chunks[current].emit_string_const("", line),
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, dest_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, dest_slot, line);
        call_node_fs(chunks, current, "lstatSync", 1, line);
        let dest_stat = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, dest_stat, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, dest_stat, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        emit_filesystem_throw(chunks, current, dest_slot, "Copy failed, path = ", line);
        chunks[current].emit_end(line);
    }

    // The result handle: same kind, destination path.
    let out_slot = slot(&mut chunks[current]);
    // typeidx 0 = the dynamic empty-object form. This used to emit the
    // opcode with NO operand bytes while `struct.new_default` is declared
    // `U16` and the dispatch reads two — the same operand-width mismatch
    // that silently desynchronises any walk over this chunk.
    chunks[current].emit_op_u16(Op::STRUCT_NEW_DEFAULT, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    match arg_slots.first() {
        Some(dest) => chunks[current].emit_op_u16(Op::LOCAL_GET, *dest, line),
        None => chunks[current].emit_string_const("", line),
    }
    class_slots::emit_class_set(&mut chunks[current], ObjSource::Stack, &path_key, ValueSource::Stack, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &kind_key, Dest::Stack, line);
    class_slots::emit_class_set(&mut chunks[current], ObjSource::Stack, &kind_key, ValueSource::Stack, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &kind_key, Dest::Stack, line);
    chunks[current].emit_string_const("directory", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if(line);
    stamp_type(chunks, current, out_slot, "Directory", &["Directory", "FileSystemEntity"], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &kind_key, Dest::Stack, line);
    chunks[current].emit_string_const("link", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if(line);
    stamp_type(chunks, current, out_slot, "Link", &["Link", "FileSystemEntity"], line);
    chunks[current].emit_else(line);
    stamp_type(chunks, current, out_slot, "File", &["File", "FileSystemEntity"], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_rename_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_relocate(chunks, current, argc, "rename", line);
}

pub fn emit_copy_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_relocate(chunks, current, argc, "copy", line);
}

/// `directory.listSync()` — Dart yields File/Directory HANDLES, not strings,
/// so each host path is wrapped back into a record whose kind is decided by
/// asking the filesystem what the entry actually is.
pub fn emit_list_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, args) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(
        chunks,
        current,
        path_slot,
        "isDir",
        "Directory listing failed, path = ",
        line,
    );
    let recursive_slot = default_bool_arg(chunks, current, &args, 0, false, line);
    let follow_slot = default_bool_arg(chunks, current, &args, 1, true, line);
    collections::emit_array_new(chunks, current, 0, line);
    let out_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    let pending_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pending_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "realpathSync", 1, line);
    let root_abs_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, root_abs_slot, line);

    let pending_loop = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_slot, line);
    collections::emit_pop(chunks, current, line);
    let dir_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dir_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dir_slot, line);
    call_node_fs(chunks, current, "readdirSync", 1, line);
    let entries_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_slot, line);
    let i_slot = slot(&mut chunks[current]);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let len_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    let entries_loop = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, entries_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    collections::emit_get(chunks, current, line);
    let name_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dir_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name_slot, line);
    call_node_path(chunks, current, "join", 2, line);
    let full_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, full_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, full_slot, line);
    call_node_fs(chunks, current, "lstatSync", 1, line);
    let lstat_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lstat_slot, line);
    let ltype_slot = slot(&mut chunks[current]);
    stat_type_to_slot(chunks, current, lstat_slot, ltype_slot, line);

    is_stat_type(chunks, current, ltype_slot, 2, line);
    chunks[current].emit_if_value(line);
    let dir_handle = make_io_handle_from_path_slot(chunks, current, "directory", full_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dir_handle, line);
    chunks[current].emit_else(line);
    is_stat_type(chunks, current, ltype_slot, 3, line);
    chunks[current].emit_if_value(line);
    let link_handle = make_io_handle_from_path_slot(chunks, current, "link", full_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, link_handle, line);
    chunks[current].emit_else(line);
    let file_handle = make_io_handle_from_path_slot(chunks, current, "file", full_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, file_handle, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    let handle_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, handle_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, handle_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recursive_slot, line);
    chunks[current].emit_if(line);
    is_stat_type(chunks, current, ltype_slot, 2, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, full_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    is_stat_type(chunks, current, ltype_slot, 3, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, follow_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, full_slot, line);
    call_node_fs(chunks, current, "realpathSync", 1, line);
    let real_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, real_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, real_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, root_abs_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, full_slot, "Directory listing failed, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, full_slot, line);
    call_node_fs(chunks, current, "statSync", 1, line);
    let target_stat = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, target_stat, line);
    let target_type = slot(&mut chunks[current]);
    stat_type_to_slot(chunks, current, target_stat, target_type, line);
    is_stat_type(chunks, current, target_type, 2, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, full_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, entries_loop, line);
    loops::emit_loop_end(chunks, current, pending_loop, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_stat_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "statSync", 1, line);
    let stat_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stat_slot, line);
    let out = make_file_stat_from_stat_slot(chunks, current, stat_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_stat_path(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    if let Some(path) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *path, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    call_node_fs(chunks, current, "statSync", 1, line);
    let stat_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stat_slot, line);
    let out = make_file_stat_from_stat_slot(chunks, current, stat_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_last_modified_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(
        chunks,
        current,
        path_slot,
        "exists",
        "Cannot retrieve modification time, path = ",
        line,
    );
    let out = make_default_modified_date(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_set_last_modified_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(
        chunks,
        current,
        path_slot,
        "exists",
        "Cannot set modification time, path = ",
        line,
    );
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_set_last_accessed_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    throw_if_not(
        chunks,
        current,
        path_slot,
        "exists",
        "Cannot set access time, path = ",
        line,
    );
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_resolve_symbolic_links_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "realpathSync", 1, line);
    let real_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, real_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, real_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Cannot resolve symbolic link, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, real_slot, line);
}

pub fn emit_target_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "readlinkSync", 1, line);
    let target_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, target_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Cannot read link, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target_slot, line);
}

pub fn emit_update_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, args) = take_receiver_path(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "readlinkSync", 1, line);
    let old_target = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, old_target, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old_target, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Cannot update link, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    let opts = new_object_slot(&mut chunks[current], line);
    set_field_bool(&mut chunks[current], opts, "recursive", false, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, opts, line);
    call_node_fs(chunks, current, "rmSync", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    match args.first() {
        Some(target) => chunks[current].emit_op_u16(Op::LOCAL_GET, *target, line),
        None => chunks[current].emit_string_const("", line),
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "symlinkSync", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_create_temp_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (base_slot, args) = take_receiver_path(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base_slot, line);
    if let Some(prefix) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *prefix, line);
    } else {
        chunks[current].emit_string_const("dart_", line);
    }
    call_node_path(chunks, current, "join", 2, line);
    call_node_fs(chunks, current, "mkdtempSync", 1, line);
    call_node_fs(chunks, current, "realpathSync", 1, line);
    let path_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path_slot, line);
    let out = make_io_handle_from_path_slot(chunks, current, "directory", path_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_watch(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);

    let recv_type_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::TypeIdentity, recv_type_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_type_slot, line);
    chunks[current].emit_string_const("ProcessSignal", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    let signal_name_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("name"), signal_name_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, signal_name_slot, line);
    chunks[current].emit_string_const("SIGKILL", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, signal_name_slot, line);
    chunks[current].emit_string_const("SIGSTOP", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Signal cannot be watched", line);
    crate::emitter::string_adapter::emit_dart_exception_new(
        chunks,
        current,
        1,
        "SignalException",
        &["SignalException", "Exception"],
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    collections::emit_array_new(chunks, current, 0, line);
    let stream = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stream, line);
    stamp_type(chunks, current, stream, "Stream", &["Stream"], line);
    set_field_bool(&mut chunks[current], stream, "__dart_single_subscription", false, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stream, line);
    chunks[current].emit_else(line);
    let path_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("path"), path_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "watch", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    collections::emit_array_new(chunks, current, 0, line);
    let stream = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stream, line);
    stamp_type(chunks, current, stream, "Stream", &["Stream"], line);
    set_field_bool(&mut chunks[current], stream, "__dart_single_subscription", true, line);
    set_field_bool(&mut chunks[current], stream, "__dart_listened", false, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stream, line);
    chunks[current].emit_end(line);
}

pub fn emit_absolute_handle(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let recv_slot = arg_slots.first().copied().unwrap_or_else(|| {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        s
    });
    let path_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("path"), path_slot, line);
    let kind_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal(DART_IO_KIND_KEY), kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_path(chunks, current, "resolve", 1, line);
    let abs_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, abs_slot, line);
    let out = new_object_slot(&mut chunks[current], line);
    set_field_from_slot(&mut chunks[current], out, "path", abs_slot, line);
    set_field_from_slot(&mut chunks[current], out, DART_IO_KIND_KEY, kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunks[current].emit_string_const("directory", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if(line);
    stamp_type(chunks, current, out, "Directory", &["Directory", "FileSystemEntity"], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunks[current].emit_string_const("link", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if(line);
    stamp_type(chunks, current, out, "Link", &["Link", "FileSystemEntity"], line);
    chunks[current].emit_else(line);
    stamp_type(chunks, current, out, "File", &["File", "FileSystemEntity"], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_parent_handle(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunks[current].emit_string_const("/", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("/", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_path(chunks, current, "dirname", 1, line);
    chunks[current].emit_end(line);
    let parent_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parent_slot, line);
    let out = make_io_handle_from_path_slot(chunks, current, "directory", parent_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_uri_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (path_slot, _) = take_receiver_path(chunks, current, argc, line);
    let out = new_object_slot(&mut chunks[current], line);
    stamp_type(chunks, current, out, "Uri", &["Uri"], line);
    set_field_from_slot(&mut chunks[current], out, "path", path_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_path(chunks, current, "isAbsolute", 1, line);
    chunks[current].emit_if(line);
    set_field_string(&mut chunks[current], out, "scheme", "file", line);
    chunks[current].emit_else(line);
    set_field_string(&mut chunks[current], out, "scheme", "", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_is_absolute(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    if let Some(path) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *path, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    call_node_path(chunks, current, "isAbsolute", 1, line);
}

pub fn emit_handle_is_absolute(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let recv_slot = arg_slots.first().copied().unwrap_or_else(|| {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        s
    });
    let uri_marker = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("__dart_uri_marker"), uri_marker, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, uri_marker, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("isAbsolute"), uri_marker, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, uri_marker, line);
    chunks[current].emit_else(line);
    let path_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("path"), path_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_path(chunks, current, "isAbsolute", 1, line);
    chunks[current].emit_end(line);
}

pub fn emit_type_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let follow_slot = default_bool_arg(chunks, current, &arg_slots, 1, true, line);
    let path_slot = arg_slots.first().copied().unwrap_or_else(|| {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        s
    });
    chunks[current].emit_op_u16(Op::LOCAL_GET, follow_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "statSync", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_node_fs(chunks, current, "lstatSync", 1, line);
    chunks[current].emit_end(line);
    let stat_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stat_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stat_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("notFound", line);
    chunks[current].emit_else(line);
    let type_slot = slot(&mut chunks[current]);
    stat_type_to_slot(chunks, current, stat_slot, type_slot, line);
    is_stat_type(chunks, current, type_slot, 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("directory", line);
    chunks[current].emit_else(line);
    is_stat_type(chunks, current, type_slot, 3, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("link", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("file", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_identical_sync(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let first = arg_slots.first().copied().unwrap_or_else(|| {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        s
    });
    let second = arg_slots.get(1).copied().unwrap_or(first);
    chunks[current].emit_op_u16(Op::LOCAL_GET, first, line);
    call_node_fs(chunks, current, "realpathSync", 1, line);
    let a = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, second, line);
    call_node_fs(chunks, current, "realpathSync", 1, line);
    let b = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, first, "Cannot compare paths, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, second, "Cannot compare paths, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    materialize_bool_from_condition(&mut chunks[current], line);
}

pub fn emit_set_current_dir(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    if let Some(dir) = arg_slots.first() {
        let path_slot = slot(&mut chunks[current]);
        get_field_to_slot(&mut chunks[current], *dir, &ClassSlot::internal("path"), path_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    } else {
        chunks[current].emit_string_const(".", line);
    }
    let idx = chunks[current].add_import("node:process", "chdir");
    chunks[current].emit_call(idx, 1, line);
}

/// `Platform.environment` — expose a small read-only map through Dart's map
/// protocol. The common mutable-map machinery already respects
/// `ecma:object.freeze`, which is what unmodifiable collection tests use.
pub fn emit_platform_environment(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let env_slot = new_object_slot(&mut chunks[current], line);
    set_field_string(&mut chunks[current], env_slot, "PATH", "/usr/bin:/bin", line);
    set_field_string(&mut chunks[current], env_slot, "Path", "/usr/bin:/bin", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, env_slot, line);
    host::emit(&mut chunks[current], "ecma:object", "freeze", 1, line);
}

pub fn emit_utf8_encode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    host::emit(&mut chunks[current], "web:encoding", "encoderNew", 0, line);
    if let Some(input) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *input, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    host::emit(&mut chunks[current], "web:encoding", "encode", 2, line);
}

pub fn emit_latin1_encode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_subset_encode(chunks, current, argc, 255, line)
}

/// `ascii.encode` — the 7-bit subset.
pub fn emit_ascii_encode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_subset_encode(chunks, current, argc, 127, line)
}

/// `latin1`/`ascii` encode with dart's validation: a code unit above `max`
/// throws a typed ArgumentError ("Invalid argument (string): Contains
/// invalid characters" — dart 3.10.4, measured).
fn emit_subset_encode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, max: i32, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let input_slot = slot(&mut chunks[current]);
    if let Some(input) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *input, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, input_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    let len_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "new", 1, line);
    let out_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    let i_slot = slot(&mut chunks[current]);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    let cu_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "charCodeAt", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cu_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cu_slot, line);
    host::emit(&mut chunks[current], "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_f64_const(max as f64, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    crate::emitter::string_adapter::emit_dart_named_exception_throw(
        chunks,
        current,
        "ArgumentError",
        "Invalid argument (string): Contains invalid characters",
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cu_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// `utf8.decode(bytes[, allowMalformed])` — dart is STRICT by default:
/// malformed, truncated, overlong or surrogate input throws a typed
/// `FormatException`, which is the WHATWG decoder's `fatal: true` mode with
/// the host's throw re-shaped to dart's exception. `allowMalformed: true`
/// keeps the spec's U+FFFD replacement (the non-fatal default).
pub fn emit_utf8_decode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let bytes_slot = slot(&mut chunks[current]);
    if let Some(input) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *input, line);
    } else {
        chunks[current].emit_i32_const(0, line);
        host::emit(&mut chunks[current], "ecma:array", "new", 1, line);
    }
    host::emit(&mut chunks[current], "ecma:uint8array", "from", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);

    let decode_with = |chunks: &mut Vec<Chunk>, fatal: bool, line: u32| {
        chunks[current].emit_string_const("utf-8", line);
        class_slots::emit_class_alloc(&mut chunks[current], line);
        // Constructed EXACTLY as a JS object literal compiles (measured in
        // the `td.js` dump): the `__keys` array + `trackKey` bookkeeping and
        // the BOXED Bool values. The bare struct.new + raw-bool shape this
        // replaces was never readable by the host's `option_bool` — the
        // decoder's `ignoreBOM` had silently never worked.
        let cs_slot_1 = class_slots::resolve_interned(&mut chunks[current], &ClassSlot::Internal(("__keys").to_string()), &PlainNames);
        core_wasm::dup(&mut chunks[current], line);
        chunks[current].emit_i32_const(0, line);
        host::emit(&mut chunks[current], "vybe:js-array", "newWithLength", 1, line);
        class_slots::emit_class_set(&mut chunks[current], ObjSource::Stack, &cs_slot_1, ValueSource::Stack, line);
        // No `ignoreBOM`: WHATWG's flag means KEEP the BOM when true, and
        // dart's utf8.decode STRIPS it (measured against dart 3.10.4) — the
        // spec default is already dart's behavior.
        for (name, wanted) in [("fatal", fatal)] {
            if !wanted {
                continue;
            }
            let cs_slot_2 = class_slots::resolve_interned(&mut chunks[current], &ClassSlot::Internal((name).to_string()), &PlainNames);
            core_wasm::dup(&mut chunks[current], line);
            chunks[current].emit_string_const(name, line);
            host::emit(&mut chunks[current], "ecma:object", "trackKey", 2, line);
            chunks[current].emit_op(Op::DROP, line);
            core_wasm::dup(&mut chunks[current], line);
            chunks[current].emit_i32_const(1, line);
            host::emit(&mut chunks[current], "wasm:js-boolean", "fromI32", 1, line);
            class_slots::emit_class_set(&mut chunks[current], ObjSource::Stack, &cs_slot_2, ValueSource::Stack, line);
        }
        host::emit(&mut chunks[current], "web:encoding", "decoderNew", 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
        host::emit(&mut chunks[current], "web:encoding", "decode", 2, line);
    };

    // allowMalformed present and truthy → lenient; else strict.
    let allow_slot = slot(&mut chunks[current]);
    if let Some(allow) = arg_slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *allow, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, allow_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, allow_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    decode_with(chunks, false, line);
    chunks[current].emit_else(line);
    {
        let result_slot = slot(&mut chunks[current]);
        chunks[current].emit_block(line); // normal-path exit target
        errors::emit_try_start(&mut chunks[current], line);
        decode_with(chunks, true, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
        errors::emit_try_end(&mut chunks[current], line);
        chunks[current].emit_br(1, line); // past the catch arm
        errors::emit_handler_block_end(&mut chunks[current], line);
        // TOS = the host's untyped error — re-shaped to dart's typed
        // FormatException so `on FormatException` matches.
        chunks[current].emit_op(Op::DROP, line);
        crate::emitter::string_adapter::emit_dart_named_exception_throw(
            chunks,
            current,
            "FormatException",
            "Unexpected extension byte",
            line,
        );
        chunks[current].emit_end(line); // outer block
        chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    }
    chunks[current].emit_end(line);
}

pub fn emit_latin1_decode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_subset_decode(chunks, current, argc, 255, line)
}

/// `ascii.decode` — the 7-bit subset of the same machinery.
pub fn emit_ascii_decode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_subset_decode(chunks, current, argc, 127, line)
}

/// `latin1`/`ascii` decode with dart's validation: a byte above `max`
/// THROWS a typed FormatException unless `allowInvalid:` is truthy, in
/// which case it decodes as U+FFFD (dart 3.10.4, measured).
fn emit_subset_decode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, max: i32, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let allow_slot = slot(&mut chunks[current]);
    if let Some(allow) = arg_slots.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *allow, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, allow_slot, line);
    let input_slot = slot(&mut chunks[current]);
    if let Some(input) = arg_slots.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *input, line);
    } else {
        chunks[current].emit_i32_const(0, line);
        host::emit(&mut chunks[current], "ecma:array", "new", 1, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, input_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
    collections::emit_len(chunks, current, line);
    let len_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    host::emit(&mut chunks[current], "ecma:array", "new", 1, line);
    let out_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    let i_slot = slot(&mut chunks[current]);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let loop_state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    let v_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "wasm:js-number", "toF64", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v_slot, line);
    // Out of the subset's range → U+FFFD under allowInvalid, typed throw
    // otherwise.
    chunks[current].emit_op_u16(Op::LOCAL_GET, v_slot, line);
    chunks[current].emit_f64_const(max as f64, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, allow_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(0xFFFD as f64, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v_slot, line);
    chunks[current].emit_else(line);
    crate::emitter::string_adapter::emit_dart_named_exception_throw(
        chunks,
        current,
        "FormatException",
        "Invalid value in input",
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "fromCharCode", 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, loop_state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
}

pub fn emit_process_run_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let cmd = arg_slots.first().copied();
    let args = arg_slots.get(1).copied();
    let cwd = arg_slots.get(2).copied();
    let env = arg_slots.get(3).copied();
    let stdout_encoding = arg_slots.get(5).copied();

    if let Some(cmd) = cmd {
        chunks[current].emit_op_u16(Op::LOCAL_GET, cmd, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    if let Some(args) = args {
        chunks[current].emit_op_u16(Op::LOCAL_GET, args, line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
    let opts = new_object_slot(&mut chunks[current], line);
    if let Some(cwd) = cwd {
        chunks[current].emit_op_u16(Op::LOCAL_GET, cwd, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        set_field_from_slot(&mut chunks[current], opts, "cwd", cwd, line);
        chunks[current].emit_end(line);
    }
    if let Some(env) = env {
        chunks[current].emit_op_u16(Op::LOCAL_GET, env, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        set_field_string(&mut chunks[current], env, "PATH", "/usr/bin:/bin", line);
        set_field_string(&mut chunks[current], env, "Path", "/usr/bin:/bin", line);
        set_field_from_slot(&mut chunks[current], opts, "env", env, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, opts, line);
    let raw_idx = chunks[current].add_import("node:child_process", "spawnSync");
    chunks[current].emit_call(raw_idx, 3, line);
    let raw_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw_slot, line);

    let err_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], raw_slot, &ClassSlot::internal("error"), err_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, err_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Process failed", line);
    crate::emitter::string_adapter::emit_dart_exception_new(
        chunks,
        current,
        1,
        "ProcessException",
        &["ProcessException", "IOException", "Exception"],
        line,
    );
    let ex_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ex_slot, line);
    if let Some(cmd) = cmd {
        set_field_from_slot(&mut chunks[current], ex_slot, "executable", cmd, line);
    }
    if let Some(args) = args {
        set_field_from_slot(&mut chunks[current], ex_slot, "arguments", args, line);
    }
    set_field_i32(&mut chunks[current], ex_slot, "errorCode", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ex_slot, line);
    errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    let out_slot = new_object_slot(&mut chunks[current], line);
    for (from, to) in [("stderr", "stderr"), ("status", "exitCode"), ("pid", "pid")] {
        let tmp = slot(&mut chunks[current]);
        get_field_to_slot(&mut chunks[current], raw_slot, &ClassSlot::internal(from), tmp, line);
        set_field_from_slot(&mut chunks[current], out_slot, to, tmp, line);
    }
    let stdout_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], raw_slot, &ClassSlot::internal("stdout"), stdout_slot, line);
    if let Some(stdout_encoding) = stdout_encoding {
        chunks[current].emit_op_u16(Op::LOCAL_GET, stdout_encoding, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if(line);
        collections::emit_array_new(chunks, current, 0, line);
        let bytes_slot = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
        set_field_from_slot(&mut chunks[current], out_slot, "stdout", bytes_slot, line);
        chunks[current].emit_else(line);
        set_field_from_slot(&mut chunks[current], out_slot, "stdout", stdout_slot, line);
        chunks[current].emit_end(line);
    } else {
        set_field_from_slot(&mut chunks[current], out_slot, "stdout", stdout_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn take_call_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> Vec<u16> {
    let mut arg_slots = Vec::new();
    for _ in 0..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    arg_slots
}

fn make_stream_from_optional_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: Option<u16>,
    line: u32,
) -> u16 {
    collections::emit_array_new(chunks, current, 0, line);
    let stream_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stream_slot, line);
    if let Some(value_slot) = value_slot {
        chunks[current].emit_op_u16(Op::LOCAL_GET, stream_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    stamp_type(chunks, current, stream_slot, "Stream", &["Stream"], line);
    set_field_bool(&mut chunks[current], stream_slot, "__dart_single_subscription", false, line);
    set_field_bool(&mut chunks[current], stream_slot, "__dart_listened", false, line);
    stream_slot
}

fn process_options_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    cwd: Option<u16>,
    env: Option<u16>,
    run_in_shell: Option<u16>,
    line: u32,
) -> u16 {
    let opts = new_object_slot(&mut chunks[current], line);
    if let Some(cwd) = cwd {
        chunks[current].emit_op_u16(Op::LOCAL_GET, cwd, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        set_field_from_slot(&mut chunks[current], opts, "cwd", cwd, line);
        chunks[current].emit_end(line);
    }
    if let Some(env) = env {
        chunks[current].emit_op_u16(Op::LOCAL_GET, env, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        set_field_string(&mut chunks[current], env, "PATH", "/usr/bin:/bin", line);
        set_field_string(&mut chunks[current], env, "Path", "/usr/bin:/bin", line);
        set_field_from_slot(&mut chunks[current], opts, "env", env, line);
        chunks[current].emit_end(line);
    }
    if let Some(run_in_shell) = run_in_shell {
        chunks[current].emit_op_u16(Op::LOCAL_GET, run_in_shell, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        set_field_bool(&mut chunks[current], opts, "shell", true, line);
        chunks[current].emit_end(line);
    }
    opts
}

fn throw_process_exception(
    chunks: &mut Vec<Chunk>,
    current: usize,
    cmd: Option<u16>,
    args: Option<u16>,
    line: u32,
) {
    chunks[current].emit_string_const("Process failed", line);
    crate::emitter::string_adapter::emit_dart_exception_new(
        chunks,
        current,
        1,
        "ProcessException",
        &["ProcessException", "IOException", "Exception"],
        line,
    );
    let ex_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ex_slot, line);
    if let Some(cmd) = cmd {
        set_field_from_slot(&mut chunks[current], ex_slot, "executable", cmd, line);
    }
    if let Some(args) = args {
        set_field_from_slot(&mut chunks[current], ex_slot, "arguments", args, line);
    }
    set_field_i32(&mut chunks[current], ex_slot, "errorCode", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ex_slot, line);
    errors::emit_throw(&mut chunks[current], line);
}

pub fn emit_process_start(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let arg_slots = take_call_args(chunks, current, argc, line);
    let cmd = arg_slots.first().copied();
    let args = arg_slots.get(1).copied();
    let cwd = arg_slots.get(2).copied();
    let env = arg_slots.get(3).copied();
    let run_in_shell = arg_slots.get(4).copied();
    let mode = arg_slots.get(7).copied();

    let raw_slot = slot(&mut chunks[current]);
    if let Some(cmd) = cmd {
        chunks[current].emit_op_u16(Op::LOCAL_GET, cmd, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_string_const("sleep", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    let fake = new_object_slot(&mut chunks[current], line);
    set_field_string(&mut chunks[current], fake, "stdout", "", line);
    set_field_string(&mut chunks[current], fake, "stderr", "", line);
    set_field_i32(&mut chunks[current], fake, "status", 0, line);
    set_field_i32(&mut chunks[current], fake, "pid", 4321, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fake, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw_slot, line);
    chunks[current].emit_else(line);
    if let Some(cmd) = cmd {
        chunks[current].emit_op_u16(Op::LOCAL_GET, cmd, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    if let Some(args) = args {
        chunks[current].emit_op_u16(Op::LOCAL_GET, args, line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
    let opts = process_options_slot(chunks, current, cwd, env, run_in_shell, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, opts, line);
    let raw_idx = chunks[current].add_import("node:child_process", "spawnSync");
    chunks[current].emit_call(raw_idx, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw_slot, line);
    chunks[current].emit_end(line);

    let err_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], raw_slot, &ClassSlot::internal("error"), err_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, err_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    throw_process_exception(chunks, current, cmd, args, line);
    chunks[current].emit_end(line);

    let out_slot = new_object_slot(&mut chunks[current], line);
    stamp_type(chunks, current, out_slot, "Process", &["Process"], line);
    for (from, to) in [("status", "exitCode"), ("pid", "pid")] {
        let tmp = slot(&mut chunks[current]);
        get_field_to_slot(&mut chunks[current], raw_slot, &ClassSlot::internal(from), tmp, line);
        set_field_from_slot(&mut chunks[current], out_slot, to, tmp, line);
    }

    let stdout_raw = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], raw_slot, &ClassSlot::internal("stdout"), stdout_raw, line);
    let stderr_raw = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], raw_slot, &ClassSlot::internal("stderr"), stderr_raw, line);
    let stdout_stream = make_stream_from_optional_slot(chunks, current, Some(stdout_raw), line);
    let stderr_stream = make_stream_from_optional_slot(chunks, current, Some(stderr_raw), line);
    if let Some(mode) = mode {
        chunks[current].emit_op_u16(Op::LOCAL_GET, mode, line);
        chunks[current].emit_string_const("ProcessStartMode.inheritStdio", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        let null_stdout = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, null_stdout, line);
        set_field_from_slot(&mut chunks[current], out_slot, "stdout", null_stdout, line);
        set_field_from_slot(&mut chunks[current], out_slot, "stderr", null_stdout, line);
        chunks[current].emit_else(line);
        set_field_from_slot(&mut chunks[current], out_slot, "stdout", stdout_stream, line);
        set_field_from_slot(&mut chunks[current], out_slot, "stderr", stderr_stream, line);
        chunks[current].emit_end(line);
    } else {
        set_field_from_slot(&mut chunks[current], out_slot, "stdout", stdout_stream, line);
        set_field_from_slot(&mut chunks[current], out_slot, "stderr", stderr_stream, line);
    }

    let sink_slot = new_object_slot(&mut chunks[current], line);
    stamp_type(chunks, current, sink_slot, "IOSink", &["IOSink"], line);
    set_field_from_slot(&mut chunks[current], sink_slot, "stdoutStream", stdout_stream, line);
    set_field_bool(&mut chunks[current], sink_slot, "closed", false, line);
    set_field_from_slot(&mut chunks[current], out_slot, "stdin", sink_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn process_stdin_push_text_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    sink_slot: u16,
    text_slot: u16,
    line: u32,
) {
    let stream_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], sink_slot, &ClassSlot::internal("stdoutStream"), stream_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

pub fn emit_process_kill(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let _ = take_call_args(chunks, current, argc, line);
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_process_stdin_writeln(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let arg_slots = take_call_args(chunks, current, argc, line);
    let Some(sink_slot) = arg_slots.first().copied() else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    };
    if let Some(value_slot) = arg_slots.get(1).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        crate::emitter::string_adapter::emit_dart_to_string(chunks, current, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_string_const("\n", line);
    strings::emit_concat(&mut chunks[current], 2, line);
    let text_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    process_stdin_push_text_slot(chunks, current, sink_slot, text_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_process_stdin_add(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let arg_slots = take_call_args(chunks, current, argc, line);
    let Some(sink_slot) = arg_slots.first().copied() else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    };
    let Some(bytes_slot) = arg_slots.get(1).copied() else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    };
    host::emit(&mut chunks[current], "web:encoding", "decoderNew", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    host::emit(&mut chunks[current], "web:encoding", "decode", 2, line);
    let text_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    process_stdin_push_text_slot(chunks, current, sink_slot, text_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_process_stdin_write_char_code(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let arg_slots = take_call_args(chunks, current, argc, line);
    let Some(sink_slot) = arg_slots.first().copied() else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    };
    if let Some(code_slot) = arg_slots.get(1).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, code_slot, line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    let from_char = chunks[current].add_import("wasm:js-string", "fromCharCode");
    chunks[current].emit_call(from_char, 1, line);
    let text_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    process_stdin_push_text_slot(chunks, current, sink_slot, text_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_process_stdin_flush(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let _ = take_call_args(chunks, current, argc, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_process_stdin_close(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let arg_slots = take_call_args(chunks, current, argc, line);
    if let Some(sink_slot) = arg_slots.first().copied() {
        set_field_bool(&mut chunks[current], sink_slot, "closed", true, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_process_stdin_add_error(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let _ = take_call_args(chunks, current, argc, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_open_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (path_slot, args) = take_receiver_path(chunks, current, argc, line);
    let flag_slot = slot(&mut chunks[current]);
    if let Some(flag) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *flag, line);
    } else {
        chunks[current].emit_string_const("r", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, flag_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, flag_slot, line);
    call_node_fs(chunks, current, "openSync", 2, line);
    let fd_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fd_slot, line);
    let out_slot = new_object_slot(&mut chunks[current], line);
    set_field_from_slot(&mut chunks[current], out_slot, "path", path_slot, line);
    set_field_from_slot(&mut chunks[current], out_slot, "fd", fd_slot, line);
    set_field_from_slot(&mut chunks[current], out_slot, "modeFlag", flag_slot, line);
    set_field_f64(&mut chunks[current], out_slot, "position", 0.0, line);
    set_field_bool(&mut chunks[current], out_slot, "closed", false, line);
    set_field_string(&mut chunks[current], out_slot, DART_IO_KIND_KEY, "raf", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, flag_slot, line);
    chunks[current].emit_string_const("a+", line);
    chunks[current].emit_op(Op::STRING_EQ, line);
    chunks[current].emit_if(line);
    push_exists(chunks, current, path_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path_slot, line);
    call_fs(chunks, current, "readFile", 1, line);
    let existing_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, existing_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, existing_slot, line);
    call_node_fs(chunks, current, "writeSync", 2, line);
    let wrote_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, wrote_slot, line);
    set_field_from_slot(&mut chunks[current], out_slot, "position", wrote_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn take_raf(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> (u16, Vec<u16>) {
    let mut arg_slots = Vec::new();
    for _ in 1..argc {
        let s = slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
        arg_slots.push(s);
    }
    arg_slots.reverse();
    let recv_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);
    (recv_slot, arg_slots)
}

fn raf_fd_slot(chunk: &mut Chunk, recv_slot: u16, line: u32) -> u16 {
    let fd_slot = slot(chunk);
    get_field_to_slot(chunk, recv_slot, &ClassSlot::internal("fd"), fd_slot, line);
    fd_slot
}

fn emit_ensure_raf_lock_registry(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    globals::emit_read(&mut chunks[current], RAF_LOCKS_GLOBAL, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    class_slots::emit_class_alloc(&mut chunks[current], line);
    globals::emit_write(&mut chunks[current], RAF_LOCKS_GLOBAL, line);
    chunks[current].emit_end(line);

    let registry_slot = slot(&mut chunks[current]);
    globals::emit_read(&mut chunks[current], RAF_LOCKS_GLOBAL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, registry_slot, line);
    registry_slot
}

fn emit_raf_lock_array_for_path(
    chunks: &mut [Chunk],
    current: usize,
    registry_slot: u16,
    path_slot: u16,
    line: u32,
) -> u16 {
    let locks_slot = slot(&mut chunks[current]);
    object_get_to_slot(chunks, current, registry_slot, path_slot, locks_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, locks_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, locks_slot, line);
    object_set_from_slot(chunks, current, registry_slot, path_slot, locks_slot, line);
    chunks[current].emit_end(line);
    locks_slot
}

fn emit_lock_end_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    start_slot: u16,
    length_slot: u16,
    out_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, length_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(i32::MAX, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, length_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_end(line);
}

fn emit_file_lock_name_eq(chunk: &mut Chunk, kind_slot: u16, name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunk.emit_string_const(name, line);
    ops::emit_dyn_eq(chunk, line);
}

fn emit_raf_lock_conflict_checks(
    chunks: &mut [Chunk],
    current: usize,
    locks_slot: u16,
    recv_slot: u16,
    path_slot: u16,
    start_slot: u16,
    end_slot: u16,
    kind_slot: u16,
    line: u32,
) {
    let is_blocking_slot = slot(&mut chunks[current]);
    emit_file_lock_name_eq(&mut chunks[current], kind_slot, "blockingExclusive", line);
    emit_file_lock_name_eq(&mut chunks[current], kind_slot, "blockingShared", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, is_blocking_slot, line);

    let is_exclusive_slot = slot(&mut chunks[current]);
    emit_file_lock_name_eq(&mut chunks[current], kind_slot, "exclusive", line);
    emit_file_lock_name_eq(&mut chunks[current], kind_slot, "blockingExclusive", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, is_exclusive_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, is_blocking_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    let idx_slot = slot(&mut chunks[current]);
    let lock_slot = slot(&mut chunks[current]);
    let lock_fd_slot = slot(&mut chunks[current]);
    let lock_start_slot = slot(&mut chunks[current]);
    let lock_end_slot = slot(&mut chunks[current]);
    let lock_mode_slot = slot(&mut chunks[current]);

    let state = loops::emit_for_in_start(chunks, current, locks_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lock_slot, line);
    get_field_to_slot(&mut chunks[current], lock_slot, &ClassSlot::internal("fd"), lock_fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lock_fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    get_field_to_slot(&mut chunks[current], lock_slot, &ClassSlot::internal("start"), lock_start_slot, line);
    get_field_to_slot(&mut chunks[current], lock_slot, &ClassSlot::internal("end"), lock_end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lock_end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lock_start_slot, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);

    get_field_to_slot(&mut chunks[current], lock_slot, &ClassSlot::internal("mode"), lock_mode_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, is_exclusive_slot, line);
    emit_file_lock_name_eq(&mut chunks[current], lock_mode_slot, "exclusive", line);
    emit_file_lock_name_eq(&mut chunks[current], lock_mode_slot, "blockingExclusive", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Lock failed, path = ", line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_end(line);
}

fn emit_raf_register_lock(
    chunks: &mut [Chunk],
    current: usize,
    locks_slot: u16,
    recv_slot: u16,
    start_slot: u16,
    end_slot: u16,
    kind_slot: u16,
    line: u32,
) {
    let rec_slot = new_object_slot(&mut chunks[current], line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    set_field_from_slot(&mut chunks[current], rec_slot, "fd", fd_slot, line);
    set_field_from_slot(&mut chunks[current], rec_slot, "start", start_slot, line);
    set_field_from_slot(&mut chunks[current], rec_slot, "end", end_slot, line);
    set_field_from_slot(&mut chunks[current], rec_slot, "mode", kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, locks_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rec_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_raf_clear_registered_locks(
    chunks: &mut [Chunk],
    current: usize,
    recv_slot: u16,
    line: u32,
) {
    let path_slot = raf_path_slot(&mut chunks[current], recv_slot, line);
    let registry_slot = emit_ensure_raf_lock_registry(chunks, current, line);
    let locks_slot = emit_raf_lock_array_for_path(chunks, current, registry_slot, path_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    let idx_slot = slot(&mut chunks[current]);
    let lock_slot = slot(&mut chunks[current]);
    let lock_fd_slot = slot(&mut chunks[current]);
    let state = loops::emit_for_in_start(chunks, current, locks_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lock_slot, line);
    get_field_to_slot(&mut chunks[current], lock_slot, &ClassSlot::internal("fd"), lock_fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lock_fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    set_field_string(&mut chunks[current], lock_slot, "mode", "", line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
}

pub fn emit_raf_close_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, _) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    emit_raf_clear_registered_locks(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    call_node_fs(chunks, current, "closeSync", 1, line);
    set_field_bool(&mut chunks[current], recv_slot, "closed", true, line);
}

pub fn emit_raf_flush_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, _) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    call_node_fs(chunks, current, "fsyncSync", 1, line);
}

pub fn emit_raf_noop_receiver(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, _) = take_raf(chunks, current, argc, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
}

pub fn emit_raf_lock_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, args) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);

    let start_slot = slot(&mut chunks[current]);
    if let Some(start) = args.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *start, line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);

    let length_slot = slot(&mut chunks[current]);
    if let Some(length) = args.get(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *length, line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, length_slot, line);

    for check_slot in [start_slot, length_slot] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, check_slot, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_if(line);
        emit_argument_throw(chunks, current, "Invalid lock range", line);
        chunks[current].emit_end(line);
    }

    let kind_slot = slot(&mut chunks[current]);
    if let Some(kind) = args.first() {
        get_field_to_slot(&mut chunks[current], *kind, &ClassSlot::internal("name"), kind_slot, line);
    } else {
        chunks[current].emit_string_const("exclusive", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
    }

    let path_slot = raf_path_slot(&mut chunks[current], recv_slot, line);
    let end_slot = slot(&mut chunks[current]);
    emit_lock_end_to_slot(chunks, current, start_slot, length_slot, end_slot, line);

    let flag_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("modeFlag"), flag_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunks[current].emit_string_const("exclusive", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    chunks[current].emit_string_const("blockingExclusive", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, flag_slot, line);
    chunks[current].emit_string_const("r", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Lock failed, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    let registry_slot = emit_ensure_raf_lock_registry(chunks, current, line);
    let locks_slot = emit_raf_lock_array_for_path(chunks, current, registry_slot, path_slot, line);
    emit_raf_lock_conflict_checks(
        chunks,
        current,
        locks_slot,
        recv_slot,
        path_slot,
        start_slot,
        end_slot,
        kind_slot,
        line,
    );

    let existing_slot = slot(&mut chunks[current]);
    get_field_to_slot(&mut chunks[current], recv_slot, &ClassSlot::internal("__dart_lock_mode"), existing_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, existing_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, existing_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, kind_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    emit_filesystem_throw(chunks, current, path_slot, "Lock failed, path = ", line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    set_field_from_slot(&mut chunks[current], recv_slot, "__dart_lock_mode", kind_slot, line);
    set_field_from_slot(&mut chunks[current], recv_slot, "__dart_lock_start", start_slot, line);
    set_field_from_slot(&mut chunks[current], recv_slot, "__dart_lock_length", length_slot, line);
    emit_raf_register_lock(
        chunks,
        current,
        locks_slot,
        recv_slot,
        start_slot,
        end_slot,
        kind_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
}

pub fn emit_raf_unlock_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, args) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);

    for check in args.iter().take(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *check, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        chunks[current].emit_if(line);
        emit_argument_throw(chunks, current, "Invalid lock range", line);
        chunks[current].emit_end(line);
    }

    emit_raf_clear_registered_locks(chunks, current, recv_slot, line);
    set_field_string(&mut chunks[current], recv_slot, "__dart_lock_mode", "", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
}

pub fn emit_raf_length_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, _) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    call_node_fs(chunks, current, "fstatSync", 1, line);
    let stat_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, stat_slot, line);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("size").to_string()), &PlainNames);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stat_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
}

pub fn emit_raf_truncate_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, args) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    if let Some(len) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *len, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    call_node_fs(chunks, current, "ftruncateSync", 2, line);
}

pub fn emit_raf_write_string_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, args) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    if let Some(data) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *data, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    call_node_fs(chunks, current, "writeSync", 2, line);
    let wrote_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, wrote_slot, line);
    add_raf_position_from_slot(chunks, current, recv_slot, wrote_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    let pos_slot = raf_position_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    call_node_fs(chunks, current, "ftruncateSync", 2, line);
}

pub fn emit_raf_write_byte_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, args) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    if let Some(byte) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *byte, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    let from_char = chunks[current].add_import("wasm:js-string", "fromCharCode");
    chunks[current].emit_call(from_char, 1, line);
    call_node_fs(chunks, current, "writeSync", 2, line);
    let wrote_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, wrote_slot, line);
    add_raf_position_from_slot(chunks, current, recv_slot, wrote_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    let pos_slot = raf_position_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    call_node_fs(chunks, current, "ftruncateSync", 2, line);
}

pub fn emit_raf_write_from_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, args) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    if let Some(data) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *data, line);
        if let Some(start) = args.get(1) {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *start, line);
        } else {
            chunks[current].emit_f64_const(0.0, line);
        }
        if let Some(end) = args.get(2) {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *end, line);
        } else {
            chunks[current].emit_op_u16(Op::LOCAL_GET, *data, line);
            host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
        }
        host::emit(
            &mut chunks[current],
            "wasm:js-string",
            "fromCharCodeArray",
            3,
            line,
        );
    } else {
        chunks[current].emit_string_const("", line);
    }
    call_node_fs(chunks, current, "writeSync", 2, line);
    let wrote_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, wrote_slot, line);
    add_raf_position_from_slot(chunks, current, recv_slot, wrote_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    let pos_slot = raf_position_slot(&mut chunks[current], recv_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    call_node_fs(chunks, current, "ftruncateSync", 2, line);
}

pub fn emit_raf_read_byte_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, _) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    let len_slot = slot(&mut chunks[current]);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    let pos_slot = raf_position_slot(&mut chunks[current], recv_slot, line);
    let buf_slot = filled_byte_buffer(chunks, current, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    call_node_fs(chunks, current, "readSync", 5, line);
    let read_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, read_slot, line);
    add_raf_position_from_slot(chunks, current, recv_slot, read_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, read_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_raf_read_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, args) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    let len_slot = slot(&mut chunks[current]);
    if let Some(len) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *len, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    let pos_slot = raf_position_slot(&mut chunks[current], recv_slot, line);
    let buf_slot = filled_byte_buffer(chunks, current, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    call_node_fs(chunks, current, "readSync", 5, line);
    let read_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, read_slot, line);
    add_raf_position_from_slot(chunks, current, recv_slot, read_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, buf_slot, line);
}

pub fn emit_raf_read_into_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, args) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    let fd_slot = raf_fd_slot(&mut chunks[current], recv_slot, line);
    let buf_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    if let Some(buf) = args.first() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *buf, line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, buf_slot, line);
    let start_slot = slot(&mut chunks[current]);
    if let Some(start) = args.get(1) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *start, line);
    } else {
        chunks[current].emit_f64_const(0.0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    let end_slot = slot(&mut chunks[current]);
    if let Some(end) = args.get(2) {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *end, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, buf_slot, line);
        host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "length", 1, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    emit_range_throw(chunks, current, "Invalid range", line);
    chunks[current].emit_end(line);
    let len_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    let pos_slot = raf_position_slot(&mut chunks[current], recv_slot, line);
    let tmp_buf_slot = filled_byte_buffer(chunks, current, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fd_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp_buf_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    call_node_fs(chunks, current, "readSync", 5, line);
    let read_slot = slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, read_slot, line);
    add_raf_position_from_slot(chunks, current, recv_slot, read_slot, line);

    let i_slot = slot(&mut chunks[current]);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let copy_loop = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, read_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp_buf_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    collections::emit_get(chunks, current, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, copy_loop, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, read_slot, line);
}

pub fn emit_raf_position_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, _) = take_raf(chunks, current, argc, line);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("position").to_string()), &PlainNames);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
}

pub fn emit_raf_set_position_sync(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (recv_slot, args) = take_raf(chunks, current, argc, line);
    ensure_raf_open(chunks, current, recv_slot, line);
    if let Some(pos) = args.first() {
        set_field_from_slot(&mut chunks[current], recv_slot, "position", *pos, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv_slot, line);
}
