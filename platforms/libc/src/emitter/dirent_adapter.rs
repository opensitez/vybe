//! POSIX directory streams over the shared WASI filesystem path operations.

use vybe_compiler::primitives::fs_path;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::{Op, heaptype::HT_EXTERN};

pub fn emit(name: &str, chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> bool {
    match name {
        "libc.dirent.opendir" if argc == 1 => emit_opendir(&mut chunks[current], line),
        "libc.dirent.readdir" if argc == 1 => emit_readdir(&mut chunks[current], line),
        "libc.dirent.closedir" if argc == 1 => emit_closedir(&mut chunks[current], line),
        _ => return false,
    }
    true
}

fn object_set(chunk: &mut Chunk, object: u16, field: &str, value: u16, line: u32) {
    let set = chunk.add_import("ecma:object", "set");
    chunk.emit_op_u16(Op::LOCAL_GET, object, line);
    chunk.emit_string_const(field, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
}

fn object_get(chunk: &mut Chunk, object: u16, field: &str, line: u32) {
    let get = chunk.add_import("ecma:object", "get");
    chunk.emit_op_u16(Op::LOCAL_GET, object, line);
    chunk.emit_string_const(field, line);
    chunk.emit_call(get, 2, line);
}

fn emit_opendir(chunk: &mut Chunk, line: u32) {
    let path = chunk.alloc_scratch(1);
    let entries = chunk.alloc_scratch(1);
    let dir = chunk.alloc_scratch(1);
    let index = chunk.alloc_scratch(1);
    let new = chunk.add_import("ecma:object", "new");
    let unshift = chunk.add_import("ecma:array", "unshift");

    chunk.emit_op_u16(Op::LOCAL_SET, path, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_is_dir(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_list_dir(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entries, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entries, line);
    chunk.emit_string_const("..", line);
    chunk.emit_call(unshift, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entries, line);
    chunk.emit_string_const(".", line);
    chunk.emit_call(unshift, 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_call(new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dir, line);
    object_set(chunk, dir, "entries", entries, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    object_set(chunk, dir, "index", index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, dir, line);
    chunk.emit_else(line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_end(line);
}

fn emit_readdir(chunk: &mut Chunk, line: u32) {
    let dir = chunk.alloc_scratch(1);
    let entries = chunk.alloc_scratch(1);
    let index = chunk.alloc_scratch(1);
    let next = chunk.alloc_scratch(1);
    let entry = chunk.alloc_scratch(1);
    let name = chunk.alloc_scratch(1);
    let length = chunk.add_import("ecma:array", "length");
    let at = chunk.add_import("ecma:array", "at");
    let new = chunk.add_import("ecma:object", "new");

    chunk.emit_op_u16(Op::LOCAL_SET, dir, line);
    chunk.emit_op_u16(Op::LOCAL_GET, dir, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_else(line);
    object_get(chunk, dir, "entries", line);
    chunk.emit_op_u16(Op::LOCAL_SET, entries, line);
    object_get(chunk, dir, "index", line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entries, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_if_value(line);
    chunk.emit_ref_null(HT_EXTERN, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, next, line);
    object_set(chunk, dir, "index", next, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entries, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_call(at, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name, line);
    chunk.emit_call(new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entry, line);
    object_set(chunk, entry, "d_name", name, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entry, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_closedir(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
}
