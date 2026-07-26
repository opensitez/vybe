use std::sync::Arc;
use vybe_emitter::instructions::host;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use vybe_emitter::{collections, loops};

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

pub fn emit_file_read_all_lines(chunks: &mut [Chunk], current: usize, line: u32) {
    let read_idx = chunks[current].add_import("node:fs", "readFileSync");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("utf8")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, read_idx, line);
    chunk.emit(2, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    host::emit(chunk, "ecma:string", "split", 2, line);
}

pub fn emit_path_get_file_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("\\")), line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    host::emit(chunk, "ecma:string", "replaceAll", 3, line);
    host::emit(chunk, "wasi:filesystem", "pathGetFileName", 1, line);
}

fn emit_normalized_path(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::String(Arc::from("\\")), line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    host::emit(chunk, "ecma:string", "replaceAll", 3, line);
}

pub fn emit_path_get_directory_name(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    let dirname = chunks[current].add_import("node:path", "dirname");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, dirname, line);
    chunk.emit(1, line);
}

pub fn emit_path_get_file_name_without_extension(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    let ext_slot = reserve_slot(&mut chunks[current]);
    let path_slot = reserve_slot(&mut chunks[current]);
    let extname = chunks[current].add_import("node:path", "extname");
    let basename = chunks[current].add_import("node:path", "basename");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, extname, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ext_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ext_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, basename, line);
    chunk.emit(2, line);
}

pub fn emit_path_combine(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let combine = chunks[current].add_import("wasi:filesystem", "pathCombine");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, combine, line);
    chunk.emit(argc, line);
}

pub fn emit_path_change_extension(chunks: &mut [Chunk], current: usize, line: u32) {
    let change = chunks[current].add_import("wasi:filesystem", "pathChangeExtension");
    let chunk = &mut chunks[current];
    let ext_slot = reserve_slot(chunk);
    let path_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, ext_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ext_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ext_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::CALL_IMPORT, change, line);
    chunk.emit(2, line);
}

pub fn emit_path_get_full_path(chunks: &mut [Chunk], current: usize, line: u32) {
    let resolve = chunks[current].add_import("node:path", "resolve");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, resolve, line);
    chunk.emit(1, line);
}

pub fn emit_path_get_path_root(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    let parse = chunks[current].add_import("node:path", "parse");
    let chunk = &mut chunks[current];
    let root_key = chunk.add_constant(Value::String(Arc::from("root")));
    chunk.emit_op_u16(Op::CALL_IMPORT, parse, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::STRUCT_GET, root_key, line);
}

pub fn emit_path_get_temp_file_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let temp = chunks[current].add_import("wasi:filesystem", "pathGetTempPath");
    let write = chunks[current].add_import("node:fs", "writeFileSync");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::CALL_IMPORT, temp, line);
    chunk.emit(0, line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    push_const(chunk, Value::String(Arc::from("vybe-dotnet-")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    push_const(chunk, Value::F64(1000000.0), line);
    host::emit(chunk, "ecma:math", "random", 0, line);
    chunk.emit_op(Op::F64_MUL, line);
    host::emit(chunk, "ecma:math", "floor", 1, line);
    host::emit(chunk, "ecma:string", "toString", 1, line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    push_const(chunk, Value::String(Arc::from(".tmp")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, write, line);
    chunk.emit(2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
}

pub fn emit_path_get_random_file_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    push_const(chunk, Value::F64(1000000000000.0), line);
    host::emit(chunk, "ecma:math", "random", 0, line);
    chunk.emit_op(Op::F64_MUL, line);
    host::emit(chunk, "ecma:math", "floor", 1, line);
    host::emit(chunk, "ecma:string", "toString", 1, line);
    push_const(chunk, Value::String(Arc::from(".tmp")), line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
}

pub fn emit_path_get_invalid_file_name_chars(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    for ch in ["\"", "<", ">", "|", "\0", ":", "*", "?", "\\", "/"] {
        push_const(chunk, Value::String(Arc::from(ch)), line);
    }
    collections::emit_array_new(chunks, current, 10, line);
}

pub fn emit_path_get_invalid_path_chars(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    for ch in ["\"", "<", ">", "|", "\0"] {
        push_const(chunk, Value::String(Arc::from(ch)), line);
    }
    collections::emit_array_new(chunks, current, 5, line);
}

pub fn emit_path_has_extension(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "node:path", "extname", 1, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op(Op::NE, line);
    vybe_emitter::ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_path_is_path_rooted(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    host::emit(&mut chunks[current], "node:path", "isAbsolute", 1, line);
}

pub fn emit_path_get_relative_path(chunks: &mut [Chunk], current: usize, line: u32) {
    let relative = chunks[current].add_import("node:path", "relative");
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, relative, line);
    chunk.emit(2, line);
}

pub fn emit_path_trim_ending_directory_separator(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("/")), line);
    host::emit(chunk, "ecma:string", "endsWith", 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("\\")), line);
    host::emit(chunk, "ecma:string", "endsWith", 2, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    host::emit(chunk, "ecma:string", "length", 1, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    host::emit(chunk, "ecma:string", "slice", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_directory_entries(chunks: &mut [Chunk], current: usize, line: u32, want_directories: bool) {
    let list_idx = chunks[current].add_import("wasi:filesystem", "listDir");
    let is_dir_idx = chunks[current].add_import("wasi:filesystem", "isDir");
    let resolve_idx = chunks[current].add_import("node:path", "resolve");

    let chunk = &mut chunks[current];
    let root_slot = reserve_slot(chunk);
    let entries_slot = reserve_slot(chunk);
    let idx_slot = reserve_slot(chunk);
    let entry_slot = reserve_slot(chunk);
    let full_path_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, root_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, root_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, list_idx, line);
    chunk.emit(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entries_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, entries_slot, idx_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, entry_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, root_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, resolve_idx, line);
    chunk.emit(2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, full_path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, full_path_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, is_dir_idx, line);
    chunk.emit(1, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    if !want_directories {
        vybe_emitter::ops::emit_dyn_not(chunk, line);
    }

    let skip_push = chunk.emit_block(line);
    vybe_emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, full_path_slot, line);
    collections::emit_push(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.patch_block(skip_push);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_directory_get_files(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_directory_entries(chunks, current, line, false);
}

pub fn emit_directory_get_directories(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_directory_entries(chunks, current, line, true);
}

pub fn emit_directory_delete(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let remove_idx = chunks[current].add_import("wasi:filesystem", "remove");
    let chunk = &mut chunks[current];

    if argc > 1 {
        chunk.emit_op(Op::DROP, line);
    }

    chunk.emit_op_u16(Op::CALL_IMPORT, remove_idx, line);
    chunk.emit(1, line);
}
