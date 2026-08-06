use std::sync::Arc;
use vybe_compiler::primitives::instructions::{core_wasm, host};

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::{collections, loops};

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

fn set_field_from_slot(chunk: &mut Chunk, obj_slot: u16, name: &str, value_slot: u16, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

fn set_field_const(chunk: &mut Chunk, obj_slot: u16, name: &str, val: Value, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    push_const(chunk, val, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

fn set_field_with_stack_value(chunk: &mut Chunk, obj_slot: u16, name: &str, line: u32) {
    let value_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    set_field_from_slot(chunk, obj_slot, name, value_slot, line);
}

fn emit_file_stream_object(
    chunks: &mut [Chunk],
    current: usize,
    path_slot: u16,
    content_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let obj_slot = reserve_slot(chunk);

    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    set_field_const(
        chunk,
        obj_slot,
        "__type",
        Value::String(Arc::from("FileStream")),
        line,
    );
    set_field_from_slot(chunk, obj_slot, "__path", path_slot, line);
    set_field_from_slot(chunk, obj_slot, "__content", content_slot, line);
    set_field_from_slot(chunk, obj_slot, "__buf", content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    set_field_with_stack_value(chunk, obj_slot, "Length", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_file_write_all_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let write = chunks[current].add_import("node:fs", "writeFileSync");
    let chunk = &mut chunks[current];
    let bytes_slot = reserve_slot(chunk);
    let path_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let text_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    host::emit(&mut chunks[current], "wasm:js-string", "fromCharCode", 1, line);
    host::emit(&mut chunks[current], "ecma:string", "concat", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, state, line);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_call(write, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_file_read_all_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let read = chunks[current].add_import("node:fs", "readFileSync");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);
    let content_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("utf8")), line);
    chunk.emit_call(read, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    chunks[current].emit_op(Op::ARRAY_SET, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_file_create(chunks: &mut [Chunk], current: usize, line: u32) {
    let write = chunks[current].add_import("node:fs", "writeFileSync");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);
    let content_slot = reserve_slot(chunk);
    let obj_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_call(write, 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    set_field_const(
        chunk,
        obj_slot,
        "__type",
        Value::String(Arc::from("FileStream")),
        line,
    );
    set_field_from_slot(chunk, obj_slot, "__path", path_slot, line);
    set_field_from_slot(chunk, obj_slot, "__content", content_slot, line);
    set_field_from_slot(chunk, obj_slot, "__buf", content_slot, line);
    set_field_const(chunk, obj_slot, "Length", Value::I32(0), line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_file_open_read(chunks: &mut [Chunk], current: usize, line: u32) {
    let read = chunks[current].add_import("node:fs", "readFileSync");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);
    let content_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("utf8")), line);
    chunk.emit_call(read, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);
    emit_file_stream_object(chunks, current, path_slot, content_slot, line);
}

pub fn emit_file_stream_write_byte(chunks: &mut [Chunk], current: usize, line: u32) {
    let write = chunks[current].add_import("node:fs", "writeFileSync");
    let chunk = &mut chunks[current];
    let buf_key = chunk.add_constant(Value::String(Arc::from("__buf")));
    let path_key = chunk.add_constant(Value::String(Arc::from("__path")));
    let length_key = chunk.add_constant(Value::String(Arc::from("Length")));
    let stream_slot = reserve_slot(chunk);
    let byte_slot = reserve_slot(chunk);
    let buf_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, byte_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, stream_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buf_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, buf_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, byte_slot, line);
    host::emit(chunk, "wasm:js-string", "fromCharCode", 1, line);
    host::emit(chunk, "ecma:string", "concat", 2, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, buf_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, length_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, path_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    chunk.emit_call(write, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_file_info_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let extname = chunks[current].add_import("node:path", "extname");
    let basename = chunks[current].add_import("node:path", "basename");
    let exists = chunks[current].add_import("wasi:filesystem", "exists");
    let size = chunks[current].add_import("wasi:filesystem", "fileSize");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);
    let obj_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);
    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    set_field_const(
        chunk,
        obj_slot,
        "__type",
        Value::String(Arc::from("FileInfo")),
        line,
    );
    set_field_from_slot(chunk, obj_slot, "FullName", path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_call(basename, 1, line);
    set_field_with_stack_value(chunk, obj_slot, "Name", line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_call(extname, 1, line);
    set_field_with_stack_value(chunk, obj_slot, "Extension", line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_call(exists, 1, line);
    set_field_with_stack_value(chunk, obj_slot, "Exists", line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_call(size, 1, line);
    set_field_with_stack_value(chunk, obj_slot, "Length", line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_file_read_all_lines(chunks: &mut [Chunk], current: usize, line: u32) {
    let read_idx = chunks[current].add_import("node:fs", "readFileSync");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("utf8")), line);
    chunk.emit_call(read_idx, 2, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    host::emit(chunk, "ecma:string", "split", 2, line);
}

pub fn emit_file_write_all_lines(chunks: &mut [Chunk], current: usize, line: u32) {
    let write_idx = chunks[current].add_import("node:fs", "writeFileSync");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);
    let lines_slot = reserve_slot(chunk);
    let text_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, lines_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, lines_slot, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    host::emit(chunk, "ecma:array", "join", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunk.emit_call(write_idx, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    chunk.emit_call(dirname, 1, line);
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
    chunk.emit_call(extname, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ext_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ext_slot, line);
    chunk.emit_call(basename, 2, line);
}

pub fn emit_path_combine(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let combine = chunks[current].add_import("wasi:filesystem", "pathCombine");
    let chunk = &mut chunks[current];
    chunk.emit_call(combine, argc, line);
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

    chunk.emit_call(change, 2, line);
}

pub fn emit_path_get_full_path(chunks: &mut [Chunk], current: usize, line: u32) {
    let resolve = chunks[current].add_import("node:path", "resolve");
    let chunk = &mut chunks[current];
    chunk.emit_call(resolve, 1, line);
}

pub fn emit_path_get_path_root(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    let parse = chunks[current].add_import("node:path", "parse");
    let chunk = &mut chunks[current];
    let root_key = chunk.add_constant(Value::String(Arc::from("root")));
    chunk.emit_call(parse, 1, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, root_key, line);
}

pub fn emit_path_get_temp_file_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let temp = chunks[current].add_import("wasi:filesystem", "pathGetTempPath");
    let write = chunks[current].add_import("node:fs", "writeFileSync");
    let chunk = &mut chunks[current];
    let path_slot = reserve_slot(chunk);

    chunk.emit_call(temp, 0, line);
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
    chunk.emit_call(write, 2, line);
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
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_path_is_path_rooted(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_normalized_path(chunks, current, line);
    host::emit(&mut chunks[current], "node:path", "isAbsolute", 1, line);
}

pub fn emit_path_get_relative_path(chunks: &mut [Chunk], current: usize, line: u32) {
    let relative = chunks[current].add_import("node:path", "relative");
    let chunk = &mut chunks[current];
    chunk.emit_call(relative, 2, line);
}

pub fn emit_path_trim_ending_directory_separator(chunks: &mut [Chunk], current: usize, line: u32) {
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

fn emit_directory_entries(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
    want_directories: bool,
) {
    let list_idx = chunks[current].add_import("wasi:filesystem", "listDir");
    let is_dir_idx = chunks[current].add_import("wasi:filesystem", "isDir");
    let resolve_idx = chunks[current].add_import("node:path", "resolve");

    let chunk = &mut chunks[current];
    let root_slot = reserve_slot(chunk);
    let entries_slot = reserve_slot(chunk);
    let idx_slot = reserve_slot(chunk);
    let entry_slot = reserve_slot(chunk);
    let full_path_slot = reserve_slot(chunk);
    let pattern_slot = reserve_slot(chunk);
    let allowed_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    if argc > 1 {
        chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    } else {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunk.emit_op_u16(Op::LOCAL_SET, pattern_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, root_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, root_slot, line);
    chunk.emit_call(list_idx, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entries_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, entries_slot, idx_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, entry_slot, line);
    chunk.emit_bool_const(true, line);
    chunk.emit_op_u16(Op::LOCAL_SET, allowed_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, root_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
    chunk.emit_call(resolve_idx, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, full_path_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, full_path_slot, line);
    chunk.emit_call(is_dir_idx, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    if !want_directories {
        vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    }

    let skip_push = chunk.emit_block(line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    if !want_directories {
        chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
        chunk.emit_else(line);

        chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
        push_const(chunk, Value::String(Arc::from("*")), line);
        host::emit(chunk, "ecma:string", "startsWith", 2, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        host::emit(chunk, "ecma:string", "slice", 2, line);
        host::emit(chunk, "ecma:string", "endsWith", 2, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, entry_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, pattern_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_SET, allowed_slot, line);
        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, allowed_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
        chunk.emit_br_if(0, line);
    }

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

pub fn emit_directory_get_files(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_directory_entries(chunks, current, argc, line, false);
}

pub fn emit_directory_get_directories(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_directory_entries(chunks, current, 1, line, true);
}

pub fn emit_directory_delete(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let remove_idx = chunks[current].add_import("wasi:filesystem", "remove");
    let chunk = &mut chunks[current];

    if argc > 1 {
        chunk.emit_op(Op::DROP, line);
    }

    chunk.emit_call(remove_idx, 1, line);
}
