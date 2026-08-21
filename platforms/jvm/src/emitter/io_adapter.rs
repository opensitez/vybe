//! JVM `java.io` in-memory stream/reader/writer adapters.
//!
//! These are platform classes. They model the Java surface needed by JVM
//! frontends without putting `java.io.*` behaviour in Kotlin or Java walkers.

use vybe_compiler::primitives::{
    callable, collections, globals, instructions::host, loops, ops, strings,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const DATA: &str = "__java_io_data";
const POS: &str = "__java_io_pos";
const MARK: &str = "__java_io_mark";
const TARGET: &str = "__java_io_target";
const LINE: &str = "__java_io_line";
const IS_DIR: &str = "__java_io_is_dir";
const JVM_FILE_STORE: &str = "__jvm_java_io_file_store";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(vybe_runtime::Value::String(name.into()))
}

fn field_get(chunk: &mut Chunk, obj: u16, name: &str, line: u32) {
    get(chunk, obj, line);
    let k = key(chunk, name);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

fn field_set_from_stack(chunk: &mut Chunk, obj: u16, name: &str, line: u32) {
    let value = chunk.alloc_scratch(1);
    set(chunk, value, line);
    get(chunk, obj, line);
    get(chunk, value, line);
    let k = key(chunk, name);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

fn object_get(chunks: &mut [Chunk], current: usize, obj: u16, name: &str, line: u32) {
    get(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(name, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn object_set_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    name: &str,
    value: u16,
    line: u32,
) {
    get(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(name, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn object_set_bool(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    name: &str,
    value: bool,
    line: u32,
) {
    get(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(name, line);
    chunks[current].emit_bool_const(value, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

#[allow(dead_code)]
fn object_delete(chunks: &mut [Chunk], current: usize, obj: u16, name: &str, line: u32) {
    get(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(name, line);
    host::emit(&mut chunks[current], "ecma:object", "delete", 2, line);
}

use crate::emitter::exceptions::emit_jvm_exception_throw;

fn emit_file_store(chunks: &mut [Chunk], current: usize, line: u32) {
    globals::emit_read(&mut chunks[current], JVM_FILE_STORE, line);
    let store = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], store, line);
    get(&mut chunks[current], store, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], store, line);
    chunks[current].emit_else(line);
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
    let fresh = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fresh, line);
    get(&mut chunks[current], fresh, line);
    globals::emit_write(&mut chunks[current], JVM_FILE_STORE, line);
    get(&mut chunks[current], fresh, line);
    chunks[current].emit_end(line);
}

fn emit_file_path_from_slot(chunks: &mut [Chunk], current: usize, file: u16, line: u32) {
    field_get(&mut chunks[current], file, DATA, line);
}

fn emit_file_name_from_path_slot(chunks: &mut [Chunk], current: usize, path: u16, line: u32) {
    get(&mut chunks[current], path, line);
    chunks[current].emit_string_const("/", line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    let parts = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], parts, line);
    get(&mut chunks[current], parts, line);
    get(&mut chunks[current], parts, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
}

fn emit_file_object_from_path_record(
    chunks: &mut [Chunk],
    current: usize,
    path: u16,
    rec: u16,
    line: u32,
) {
    let obj = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    set(&mut chunks[current], obj, line);

    get(&mut chunks[current], path, line);
    field_set_from_stack(&mut chunks[current], obj, DATA, line);
    get(&mut chunks[current], path, line);
    field_set_from_stack(&mut chunks[current], obj, "path", line);
    get(&mut chunks[current], path, line);
    field_set_from_stack(&mut chunks[current], obj, "absolutePath", line);
    emit_file_name_from_path_slot(chunks, current, path, line);
    field_set_from_stack(&mut chunks[current], obj, "name", line);

    object_get(chunks, current, rec, IS_DIR, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    let is_dir = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], is_dir, line);
    get(&mut chunks[current], is_dir, line);
    field_set_from_stack(&mut chunks[current], obj, "isDirectory", line);
    get(&mut chunks[current], is_dir, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    field_set_from_stack(&mut chunks[current], obj, "isFile", line);
    chunks[current].emit_i32_const(0, line);
    field_set_from_stack(&mut chunks[current], obj, "__java_io_depth", line);

    get(&mut chunks[current], obj, line);
}

fn emit_file_object_from_path_record_with_depth(
    chunks: &mut [Chunk],
    current: usize,
    path: u16,
    rec: u16,
    depth: u16,
    line: u32,
) {
    emit_file_object_from_path_record(chunks, current, path, rec, line);
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    get(&mut chunks[current], depth, line);
    field_set_from_stack(&mut chunks[current], obj, "__java_io_depth", line);
    get(&mut chunks[current], obj, line);
}

fn emit_new_file_record(
    chunks: &mut [Chunk],
    current: usize,
    content: u16,
    is_dir: bool,
    line: u32,
) {
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);
    object_set_from_slot(chunks, current, rec, DATA, content, line);
    object_set_bool(chunks, current, rec, IS_DIR, is_dir, line);
    get(&mut chunks[current], rec, line);
}

fn emit_file_record_for_path(chunks: &mut [Chunk], current: usize, path: u16, line: u32) {
    emit_file_store(chunks, current, line);
    let store = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], store, line);
    get(&mut chunks[current], store, line);
    get(&mut chunks[current], path, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn emit_store_file_record(
    chunks: &mut [Chunk],
    current: usize,
    path: u16,
    content: u16,
    is_dir: bool,
    line: u32,
) {
    emit_file_store(chunks, current, line);
    let store = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], store, line);
    emit_new_file_record(chunks, current, content, is_dir, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);
    get(&mut chunks[current], store, line);
    get(&mut chunks[current], path, line);
    get(&mut chunks[current], rec, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

pub fn emit_file_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let path = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        let child = chunks[current].alloc_scratch(1);
        let parent = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], child, line);
        set(&mut chunks[current], parent, line);
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        get(&mut chunks[current], parent, line);
        host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
        chunks[current].emit_string_const("object", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        field_get(&mut chunks[current], parent, DATA, line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], parent, line);
        chunks[current].emit_end(line);
        chunks[current].emit_string_const("/", line);
        strings::emit_str_concat_coercing(&mut chunks[current], line);
        get(&mut chunks[current], child, line);
        strings::emit_str_concat_coercing(&mut chunks[current], line);
    } else if argc == 1 {
        set(&mut chunks[current], path, line);
        get(&mut chunks[current], path, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    set(&mut chunks[current], path, line);
    let obj = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    set(&mut chunks[current], obj, line);
    get(&mut chunks[current], path, line);
    field_set_from_stack(&mut chunks[current], obj, DATA, line);
    get(&mut chunks[current], obj, line);
}

fn new_object_with_data(chunks: &mut [Chunk], current: usize, data_slot: u16, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    set(&mut chunks[current], obj, line);
    get(&mut chunks[current], data_slot, line);
    field_set_from_stack(&mut chunks[current], obj, DATA, line);
    chunks[current].emit_i32_const(0, line);
    field_set_from_stack(&mut chunks[current], obj, POS, line);
    chunks[current].emit_i32_const(0, line);
    field_set_from_stack(&mut chunks[current], obj, MARK, line);
    chunks[current].emit_i32_const(0, line);
    field_set_from_stack(&mut chunks[current], obj, LINE, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    field_set_from_stack(&mut chunks[current], obj, TARGET, line);
    get(&mut chunks[current], obj, line);
}

pub fn emit_byte_array_output_stream_new(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    collections::emit_array_new(chunks, current, 0, line);
    let data = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], data, line);
    new_object_with_data(chunks, current, data, line);
}

pub fn emit_byte_array_input_stream_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let data = chunks[current].alloc_scratch(1);
    if argc == 0 {
        collections::emit_array_new(chunks, current, 0, line);
    } else {
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        set(&mut chunks[current], data, line);
        get(&mut chunks[current], data, line);
    }
    set(&mut chunks[current], data, line);
    new_object_with_data(chunks, current, data, line);
}

pub fn emit_string_reader_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let data = chunks[current].alloc_scratch(1);
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    } else {
        set(&mut chunks[current], data, line);
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        get(&mut chunks[current], data, line);
    }
    set(&mut chunks[current], data, line);
    new_object_with_data(chunks, current, data, line);
}

pub fn emit_passthrough_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        emit_byte_array_output_stream_new(chunks, current, 0, line);
        return;
    }
    let first = chunks[current].alloc_scratch(1);
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    set(&mut chunks[current], first, line);
    get(&mut chunks[current], first, line);
}

pub fn emit_sequence_input_stream_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        emit_passthrough_new(chunks, current, argc, line);
        return;
    }
    let second = chunks[current].alloc_scratch(1);
    let first = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], second, line);
    set(&mut chunks[current], first, line);
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    field_get(&mut chunks[current], first, DATA, line);
    collections::emit_clone(chunks, current, line);
    let data = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], data, line);
    get(&mut chunks[current], data, line);
    get(&mut chunks[current], data, line);
    collections::emit_len(chunks, current, line);
    field_get(&mut chunks[current], second, DATA, line);
    collections::emit_insert_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    new_object_with_data(chunks, current, data, line);
}

pub fn emit_string_writer_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_string_const("", line);
    let data = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], data, line);
    new_object_with_data(chunks, current, data, line);
}

pub fn emit_print_writer_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let target = chunks[current].alloc_scratch(1);
    if argc == 0 {
        emit_string_writer_new(chunks, current, 0, line);
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    set(&mut chunks[current], target, line);
    chunks[current].emit_struct_new(0, 0, line);
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    get(&mut chunks[current], target, line);
    field_set_from_stack(&mut chunks[current], obj, TARGET, line);
    chunks[current].emit_string_const("", line);
    field_set_from_stack(&mut chunks[current], obj, DATA, line);
    get(&mut chunks[current], obj, line);
}

pub fn emit_file_write_text(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    append: bool,
    line: u32,
) {
    let text = chunks[current].alloc_scratch(1);
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], text, line);
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    set(&mut chunks[current], file, line);
    emit_file_path_from_slot(chunks, current, file, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    if append {
        emit_file_record_for_path(chunks, current, path, line);
        let rec = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], rec, line);
        get(&mut chunks[current], rec, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        object_get(chunks, current, rec, DATA, line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_end(line);
        get(&mut chunks[current], text, line);
        strings::emit_str_concat_coercing(&mut chunks[current], line);
        set(&mut chunks[current], text, line);
    }
    emit_store_file_record(chunks, current, path, text, false, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_file_read_text(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], file, line);
    emit_file_path_from_slot(chunks, current, file, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    emit_file_record_for_path(chunks, current, path, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);
    get(&mut chunks[current], rec, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    object_get(chunks, current, rec, DATA, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
}

pub fn emit_file_read_lines(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_file_read_text(chunks, current, argc, line);
    chunks[current].emit_string_const("\n", line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    let lines = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], lines, line);

    get(&mut chunks[current], lines, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], lines, line);
    get(&mut chunks[current], lines, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_string_const("", line);
    let equals = chunks[current].add_import("wasm:js-string", "equals");
    chunks[current].emit_call(equals, 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], lines, line);
    collections::emit_pop(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], lines, line);
}

pub fn emit_file_list_files(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let predicate = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        set(&mut chunks[current], predicate, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        set(&mut chunks[current], predicate, line);
    }
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], file, line);
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }

    emit_file_path_from_slot(chunks, current, file, line);
    chunks[current].emit_string_const("/", line);
    strings::emit_str_concat(&mut chunks[current], line);
    let prefix = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], prefix, line);

    emit_file_store(chunks, current, line);
    let store = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], store, line);
    get(&mut chunks[current], store, line);
    host::emit(&mut chunks[current], "ecma:object", "keys", 1, line);
    let keys = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], keys, line);
    collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);

    let idx = chunks[current].alloc_scratch(1);
    let state = loops::emit_for_in_start(chunks, current, keys, idx, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);

    get(&mut chunks[current], path, line);
    get(&mut chunks[current], prefix, line);
    host::emit(&mut chunks[current], "ecma:string", "startsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], path, line);
    get(&mut chunks[current], prefix, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    host::emit(&mut chunks[current], "ecma:string", "substring", 2, line);
    let tail = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tail, line);
    get(&mut chunks[current], tail, line);
    chunks[current].emit_string_const("", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    get(&mut chunks[current], tail, line);
    chunks[current].emit_string_const("/", line);
    host::emit(&mut chunks[current], "ecma:string", "indexOf", 2, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);

    get(&mut chunks[current], store, line);
    get(&mut chunks[current], path, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);
    emit_file_object_from_path_record(chunks, current, path, rec, line);
    let child = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], child, line);

    get(&mut chunks[current], predicate, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op(Op::REF_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], predicate, line);
    get(&mut chunks[current], child, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], child, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx, state, line);

    get(&mut chunks[current], out, line);
}

pub fn emit_file_walk(chunks: &mut [Chunk], current: usize, line: u32) {
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], file, line);
    emit_file_path_from_slot(chunks, current, file, line);
    let root_path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], root_path, line);
    get(&mut chunks[current], root_path, line);
    chunks[current].emit_string_const("/", line);
    strings::emit_str_concat(&mut chunks[current], line);
    let prefix = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], prefix, line);

    emit_file_store(chunks, current, line);
    let store = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], store, line);
    collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);

    get(&mut chunks[current], store, line);
    get(&mut chunks[current], root_path, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    let root_rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], root_rec, line);
    get(&mut chunks[current], root_rec, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    let root_depth = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], root_depth, line);
    emit_file_object_from_path_record_with_depth(
        chunks, current, root_path, root_rec, root_depth, line,
    );
    let root_obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], root_obj, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], root_obj, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], store, line);
    host::emit(&mut chunks[current], "ecma:object", "keys", 1, line);
    let keys = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], keys, line);
    let idx = chunks[current].alloc_scratch(1);
    let state = loops::emit_for_in_start(chunks, current, keys, idx, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    get(&mut chunks[current], path, line);
    get(&mut chunks[current], prefix, line);
    host::emit(&mut chunks[current], "ecma:string", "startsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], path, line);
    get(&mut chunks[current], prefix, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    host::emit(&mut chunks[current], "ecma:string", "substring", 2, line);
    let tail = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], tail, line);
    get(&mut chunks[current], tail, line);
    chunks[current].emit_string_const("/", line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    collections::emit_len(chunks, current, line);
    let depth = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], depth, line);
    get(&mut chunks[current], store, line);
    get(&mut chunks[current], path, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);
    emit_file_object_from_path_record_with_depth(chunks, current, path, rec, depth, line);
    let child = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], child, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], child, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx, state, line);

    get(&mut chunks[current], out, line);
}

pub fn emit_file_exists(chunks: &mut [Chunk], current: usize, line: u32) {
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], file, line);
    emit_file_path_from_slot(chunks, current, file, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    emit_file_record_for_path(chunks, current, path, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_file_delete(chunks: &mut [Chunk], current: usize, line: u32) {
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], file, line);
    emit_file_path_from_slot(chunks, current, file, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    emit_file_store(chunks, current, line);
    let store = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], store, line);

    chunks[current].emit_bool_const(true, line);
    let can_delete = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], can_delete, line);
    emit_file_record_for_path(chunks, current, path, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);
    get(&mut chunks[current], rec, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    object_get(chunks, current, rec, IS_DIR, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], path, line);
    chunks[current].emit_string_const("/", line);
    strings::emit_str_concat(&mut chunks[current], line);
    let prefix = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], prefix, line);
    get(&mut chunks[current], store, line);
    host::emit(&mut chunks[current], "ecma:object", "keys", 1, line);
    let keys = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], keys, line);
    let state = loops::emit_for_in_start(chunks, current, keys, idx, line);
    let child = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], child, line);
    get(&mut chunks[current], child, line);
    get(&mut chunks[current], prefix, line);
    host::emit(&mut chunks[current], "ecma:string", "startsWith", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], can_delete, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx, state, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], can_delete, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], store, line);
    get(&mut chunks[current], path, line);
    host::emit(&mut chunks[current], "ecma:object", "delete", 2, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_file_mkdirs(chunks: &mut [Chunk], current: usize, line: u32) {
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], file, line);
    emit_file_path_from_slot(chunks, current, file, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    get(&mut chunks[current], path, line);
    chunks[current].emit_string_const("/", line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    let parts = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], parts, line);
    chunks[current].emit_string_const("", line);
    let current_path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], current_path, line);
    let idx = chunks[current].alloc_scratch(1);
    let state = loops::emit_for_in_start(chunks, current, parts, idx, line);
    let part = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], part, line);
    get(&mut chunks[current], part, line);
    chunks[current].emit_string_const("", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], current_path, line);
    chunks[current].emit_string_const("", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("/", line);
    get(&mut chunks[current], part, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], current_path, line);
    chunks[current].emit_string_const("/", line);
    strings::emit_str_concat(&mut chunks[current], line);
    get(&mut chunks[current], part, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], current_path, line);
    chunks[current].emit_string_const("", line);
    let content = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], content, line);
    emit_store_file_record(chunks, current, current_path, content, true, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx, state, line);
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_file_is_directory(chunks: &mut [Chunk], current: usize, want_dir: bool, line: u32) {
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], file, line);
    emit_file_path_from_slot(chunks, current, file, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    emit_file_record_for_path(chunks, current, path, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);
    get(&mut chunks[current], rec, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    object_get(chunks, current, rec, IS_DIR, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    if !want_dir {
        chunks[current].emit_op(Op::I32_EQZ, line);
    }
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_file_get_path(chunks: &mut [Chunk], current: usize, line: u32) {
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], file, line);
    emit_file_path_from_slot(chunks, current, file, line);
}

pub fn emit_file_get_name(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_file_get_path(chunks, current, line);
    chunks[current].emit_string_const("/", line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    let parts = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], parts, line);
    get(&mut chunks[current], parts, line);
    get(&mut chunks[current], parts, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_file_extension(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_file_get_name(chunks, current, line);
    let name = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], name, line);
    get(&mut chunks[current], name, line);
    chunks[current].emit_string_const(".", line);
    host::emit(&mut chunks[current], "ecma:string", "lastIndexOf", 2, line);
    let idx = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], idx, line);
    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], name, line);
    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    host::emit(&mut chunks[current], "ecma:string", "substring", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
}

pub fn emit_file_name_without_extension(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_file_get_name(chunks, current, line);
    let name = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], name, line);
    get(&mut chunks[current], name, line);
    chunks[current].emit_string_const(".", line);
    host::emit(&mut chunks[current], "ecma:string", "lastIndexOf", 2, line);
    let idx = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], idx, line);
    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], name, line);
    chunks[current].emit_i32_const(0, line);
    get(&mut chunks[current], idx, line);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], name, line);
    chunks[current].emit_end(line);
}

pub fn emit_file_parent(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_file_get_path(chunks, current, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    get(&mut chunks[current], path, line);
    chunks[current].emit_string_const("/", line);
    host::emit(&mut chunks[current], "ecma:string", "lastIndexOf", 2, line);
    let idx = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], idx, line);
    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], path, line);
    chunks[current].emit_i32_const(0, line);
    get(&mut chunks[current], idx, line);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

pub fn emit_file_parent_file(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_file_parent(chunks, current, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    chunks[current].emit_string_const("", line);
    let content = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], content, line);
    emit_store_file_record(chunks, current, path, content, true, line);
    get(&mut chunks[current], path, line);
    emit_file_new(chunks, current, 1, line);
}

pub fn emit_file_create_temp(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let suffix = chunks[current].alloc_scratch(1);
    let prefix = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        set(&mut chunks[current], suffix, line);
        set(&mut chunks[current], prefix, line);
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else {
        chunks[current].emit_string_const(".tmp", line);
        set(&mut chunks[current], suffix, line);
        if argc == 1 {
            set(&mut chunks[current], prefix, line);
        } else {
            chunks[current].emit_string_const("vybe", line);
            set(&mut chunks[current], prefix, line);
        }
    }
    chunks[current].emit_string_const("/tmp/", line);
    get(&mut chunks[current], prefix, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    host::emit(&mut chunks[current], "ecma:date", "now", 0, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    get(&mut chunks[current], suffix, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    chunks[current].emit_string_const("", line);
    let content = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], content, line);
    emit_store_file_record(chunks, current, path, content, false, line);
    get(&mut chunks[current], path, line);
    emit_file_new(chunks, current, 1, line);
}

pub fn emit_file_rename_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let dst = chunks[current].alloc_scratch(1);
    let src = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], dst, line);
    set(&mut chunks[current], src, line);
    emit_file_path_from_slot(chunks, current, src, line);
    let src_path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], src_path, line);
    emit_file_path_from_slot(chunks, current, dst, line);
    let dst_path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], dst_path, line);
    emit_file_store(chunks, current, line);
    let store = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], store, line);
    get(&mut chunks[current], store, line);
    get(&mut chunks[current], src_path, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);
    get(&mut chunks[current], rec, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], store, line);
    get(&mut chunks[current], dst_path, line);
    get(&mut chunks[current], rec, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], store, line);
    get(&mut chunks[current], src_path, line);
    host::emit(&mut chunks[current], "ecma:object", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_file_copy_to(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let overwrite = chunks[current].alloc_scratch(1);
    let dst = chunks[current].alloc_scratch(1);
    let src = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        set(&mut chunks[current], overwrite, line);
    } else {
        chunks[current].emit_bool_const(false, line);
        set(&mut chunks[current], overwrite, line);
    }
    set(&mut chunks[current], dst, line);
    set(&mut chunks[current], src, line);

    emit_file_path_from_slot(chunks, current, src, line);
    let src_path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], src_path, line);
    emit_file_path_from_slot(chunks, current, dst, line);
    let dst_path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], dst_path, line);

    emit_file_store(chunks, current, line);
    let store = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], store, line);
    get(&mut chunks[current], store, line);
    get(&mut chunks[current], src_path, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);

    get(&mut chunks[current], store, line);
    get(&mut chunks[current], dst_path, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    get(&mut chunks[current], overwrite, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    emit_jvm_exception_throw(chunks, current, "FileAlreadyExistsException", line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], rec, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], store, line);
    get(&mut chunks[current], dst_path, line);
    get(&mut chunks[current], rec, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], dst, line);
}

pub fn emit_file_input_stream(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_file_read_text(chunks, current, 1, line);
    let data = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], data, line);
    new_object_with_data(chunks, current, data, line);
}

pub fn emit_file_output_stream(chunks: &mut [Chunk], current: usize, append: bool, line: u32) {
    let file = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], file, line);
    emit_file_path_from_slot(chunks, current, file, line);
    let path = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], path, line);
    let obj = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    set(&mut chunks[current], obj, line);
    get(&mut chunks[current], path, line);
    field_set_from_stack(&mut chunks[current], obj, TARGET, line);
    if append {
        emit_file_record_for_path(chunks, current, path, line);
        let rec = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], rec, line);
        get(&mut chunks[current], rec, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        object_get(chunks, current, rec, DATA, line);
        chunks[current].emit_else(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    field_set_from_stack(&mut chunks[current], obj, DATA, line);
    get(&mut chunks[current], obj, line);
}

fn data_array_or_string_len(chunks: &mut [Chunk], current: usize, data: u16, line: u32) {
    get(&mut chunks[current], data, line);
    collections::emit_len(chunks, current, line);
}

pub fn emit_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, DATA, line);
    let data = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], data, line);
    data_array_or_string_len(chunks, current, data, line);
}

pub fn emit_reset_buffer(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    collections::emit_array_new(chunks, current, 0, line);
    field_set_from_stack(&mut chunks[current], obj, DATA, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn append_byte_to_array(chunks: &mut [Chunk], current: usize, obj: u16, value: u16, line: u32) {
    field_get(&mut chunks[current], obj, DATA, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

pub fn emit_output_write(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    if argc > 2 {
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    set(&mut chunks[current], obj, line);

    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    field_get(&mut chunks[current], obj, DATA, line);
    chunks[current].emit_i32_const(0, line);
    get(&mut chunks[current], value, line);
    collections::emit_insert_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    append_byte_to_array(chunks, current, obj, value, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_to_byte_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, DATA, line);
    collections::emit_clone(chunks, current, line);
}

fn append_text_to_target(chunks: &mut [Chunk], current: usize, target: u16, text: u16, line: u32) {
    get(&mut chunks[current], target, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    let path = target;
    emit_file_record_for_path(chunks, current, path, line);
    let rec = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], rec, line);
    get(&mut chunks[current], rec, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    object_get(chunks, current, rec, DATA, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], text, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    let content = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], content, line);
    emit_store_file_record(chunks, current, path, content, false, line);
    chunks[current].emit_else(line);
    let data = chunks[current].alloc_scratch(1);
    field_get(&mut chunks[current], target, DATA, line);
    set(&mut chunks[current], data, line);
    get(&mut chunks[current], data, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    let len = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], text, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    set(&mut chunks[current], len, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], idx, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], idx, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], data, line);
    get(&mut chunks[current], text, line);
    get(&mut chunks[current], idx, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_else(line);
    get(&mut chunks[current], data, line);
    get(&mut chunks[current], text, line);
    strings::emit_str_concat(&mut chunks[current], line);
    field_set_from_stack(&mut chunks[current], target, DATA, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn append_value_to_text_target(
    chunks: &mut [Chunk],
    current: usize,
    target: u16,
    value: u16,
    newline: bool,
    line: u32,
) {
    chunks[current].emit_string_const("", line);
    get(&mut chunks[current], value, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    if newline {
        chunks[current].emit_string_const("\n", line);
        strings::emit_str_concat(&mut chunks[current], line);
    }
    let text = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], text, line);
    append_text_to_target(chunks, current, target, text, line);
}

fn build_text_from_value_range(
    chunks: &mut [Chunk],
    current: usize,
    value: u16,
    off: u16,
    len: u16,
    out: u16,
    line: u32,
) {
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], off, line);
    get(&mut chunks[current], off, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    host::emit(&mut chunks[current], "ecma:string", "substring", 3, line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("", line);
    set(&mut chunks[current], out, line);
    let idx = chunks[current].alloc_scratch(1);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], idx, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], idx, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], off, line);
    get(&mut chunks[current], idx, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    collections::emit_get(chunks, current, line);
    let elem = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], elem, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], elem, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], elem, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], elem, line);
    host::emit(&mut chunks[current], "ecma:string", "fromCharCode", 1, line);
    chunks[current].emit_end(line);
    strings::emit_str_concat(&mut chunks[current], line);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_end(line);
}

fn append_range_to_text_target(
    chunks: &mut [Chunk],
    current: usize,
    target: u16,
    value: u16,
    off: u16,
    len: u16,
    line: u32,
) {
    let text = chunks[current].alloc_scratch(1);
    build_text_from_value_range(chunks, current, value, off, len, text, line);
    append_text_to_target(chunks, current, target, text, line);
}

fn append_char_code_to_text_target(
    chunks: &mut [Chunk],
    current: usize,
    target: u16,
    value: u16,
    line: u32,
) {
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    append_value_to_text_target(chunks, current, target, value, false, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:string", "fromCharCode", 1, line);
    let text = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], text, line);
    append_text_to_target(chunks, current, target, text, line);
    chunks[current].emit_end(line);
}

pub fn emit_writer_write(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    newline: bool,
    return_self: bool,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    let off = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let has_range = argc > 3;
    if has_range {
        set(&mut chunks[current], len, line);
        set(&mut chunks[current], off, line);
        set(&mut chunks[current], value, line);
        for _ in 3..argc - 1 {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else if argc > 1 {
        set(&mut chunks[current], value, line);
        if argc > 2 {
            for _ in 2..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
        }
    } else {
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], value, line);
    }
    set(&mut chunks[current], obj, line);

    field_get(&mut chunks[current], obj, TARGET, line);
    let target = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], target, line);
    get(&mut chunks[current], target, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if has_range {
        append_range_to_text_target(chunks, current, obj, value, off, len, line);
    } else if newline || return_self {
        append_value_to_text_target(chunks, current, obj, value, newline, line);
    } else {
        append_char_code_to_text_target(chunks, current, obj, value, line);
    }
    chunks[current].emit_else(line);
    if has_range {
        append_range_to_text_target(chunks, current, target, value, off, len, line);
    } else if newline || return_self {
        append_value_to_text_target(chunks, current, target, value, newline, line);
    } else {
        append_char_code_to_text_target(chunks, current, target, value, line);
    }
    chunks[current].emit_end(line);
    if return_self {
        get(&mut chunks[current], obj, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
}

pub fn emit_writer_print(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    newline: bool,
    return_self: bool,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    if argc > 1 {
        set(&mut chunks[current], value, line);
        for _ in 1..argc - 1 {
            chunks[current].emit_op(Op::DROP, line);
        }
    } else {
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], value, line);
    }
    set(&mut chunks[current], obj, line);

    field_get(&mut chunks[current], obj, TARGET, line);
    let target = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], target, line);
    get(&mut chunks[current], target, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    append_value_to_text_target(chunks, current, obj, value, newline, line);
    chunks[current].emit_else(line);
    append_value_to_text_target(chunks, current, target, value, newline, line);
    chunks[current].emit_end(line);
    if return_self {
        get(&mut chunks[current], obj, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
}

pub fn emit_writer_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, DATA, line);
}

pub fn emit_output_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, DATA, line);
    set(&mut chunks[current], data, line);
    get(&mut chunks[current], data, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], idx, line);
    chunks[current].emit_string_const("", line);
    set(&mut chunks[current], out, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], idx, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], data, line);
    get(&mut chunks[current], idx, line);
    collections::emit_get(chunks, current, line);
    let elem = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], elem, line);
    get(&mut chunks[current], elem, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], elem, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], elem, line);
    host::emit(&mut chunks[current], "ecma:string", "fromCharCode", 1, line);
    chunks[current].emit_end(line);
    strings::emit_str_concat(&mut chunks[current], line);
    set(&mut chunks[current], out, line);

    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    get(&mut chunks[current], out, line);
}

pub fn emit_read(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        emit_read_into_buffer(chunks, current, line);
        return;
    }
    let obj = chunks[current].alloc_scratch(1);
    let pos = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, DATA, line);
    set(&mut chunks[current], data, line);
    field_get(&mut chunks[current], obj, POS, line);
    set(&mut chunks[current], pos, line);
    get(&mut chunks[current], pos, line);
    get(&mut chunks[current], data, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], data, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], data, line);
    get(&mut chunks[current], pos, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    chunks[current].emit_else(line);
    get(&mut chunks[current], data, line);
    get(&mut chunks[current], pos, line);
    collections::emit_get(chunks, current, line);
    let element = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], element, line);
    get(&mut chunks[current], element, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], element, line);
    chunks[current].emit_i32_const(0, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    chunks[current].emit_else(line);
    get(&mut chunks[current], element, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], pos, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    field_set_from_stack(&mut chunks[current], obj, POS, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_end(line);
}

pub fn emit_read_text(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, DATA, line);
    let data = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], data, line);
    get(&mut chunks[current], data, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], data, line);
    emit_chars_to_string(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], data, line);
    chunks[current].emit_end(line);
}

pub fn emit_use(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let block = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        set(&mut chunks[current], block, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        set(&mut chunks[current], block, line);
    }
    set(&mut chunks[current], obj, line);
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }

    get(&mut chunks[current], block, line);
    get(&mut chunks[current], obj, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1, line);
}

pub fn emit_stream_copy_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let dst = chunks[current].alloc_scratch(1);
    let src = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], dst, line);
    set(&mut chunks[current], src, line);

    field_get(&mut chunks[current], src, DATA, line);
    let data = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], data, line);
    get(&mut chunks[current], data, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], data, line);
    emit_chars_to_string(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], data, line);
    chunks[current].emit_end(line);
    let text = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], text, line);

    field_get(&mut chunks[current], dst, TARGET, line);
    let target = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], target, line);
    get(&mut chunks[current], target, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    append_text_to_target(chunks, current, target, text, line);
    chunks[current].emit_else(line);
    append_text_to_target(chunks, current, dst, text, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], text, line);
    collections::emit_len(chunks, current, line);
}

fn emit_read_into_buffer(chunks: &mut [Chunk], current: usize, line: u32) {
    let buf = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    let count = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], buf, line);
    set(&mut chunks[current], obj, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], count, line);
    get(&mut chunks[current], buf, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], count, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], obj, line);
    emit_read(chunks, current, 1, line);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_i32_const(-1, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], buf, line);
    get(&mut chunks[current], count, line);
    get(&mut chunks[current], value, line);
    field_get(&mut chunks[current], obj, DATA, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCharCode",
        1,
        line,
    );
    chunks[current].emit_end(line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], count, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], count, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    get(&mut chunks[current], count, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], count, line);
    chunks[current].emit_end(line);
}

pub fn emit_available(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, DATA, line);
    collections::emit_len(chunks, current, line);
    field_get(&mut chunks[current], obj, POS, line);
    chunks[current].emit_op(Op::I32_SUB, line);
}

pub fn emit_mark(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        chunks[current].emit_op(Op::DROP, line);
    }
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, POS, line);
    field_set_from_stack(&mut chunks[current], obj, MARK, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_reset_pos(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, MARK, line);
    field_set_from_stack(&mut chunks[current], obj, POS, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_mark_supported(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_skip(chunks: &mut [Chunk], current: usize, line: u32) {
    let count = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], count, line);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, POS, line);
    get(&mut chunks[current], count, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    field_set_from_stack(&mut chunks[current], obj, POS, line);
    get(&mut chunks[current], count, line);
}

pub fn emit_flush_close(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_ready(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, POS, line);
    field_get(&mut chunks[current], obj, DATA, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_unread(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, POS, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    let pos = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], pos, line);
    get(&mut chunks[current], pos, line);
    field_set_from_stack(&mut chunks[current], obj, POS, line);
    field_get(&mut chunks[current], obj, DATA, line);
    get(&mut chunks[current], pos, line);
    get(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_writer_to_char_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, DATA, line);
    emit_string_to_char_array(chunks, current, line);
}

pub fn emit_chars_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    // A byte/char array whose elements are NUMBERS holds CODES —
    // `String(s.toByteArray())` round-trips through `fromCharCode`, not a
    // join of the digits.
    let arr = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], arr, line);
    get(&mut chunks[current], arr, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    let idx = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(idx, 1, line);
    chunks[current].emit_string_const("number", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    {
        // out = "" ; for code in arr: out += fromCharCode(code)
        let out = chunks[current].alloc_scratch(1);
        let i = chunks[current].alloc_scratch(1);
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], out, line);
        let state =
            vybe_compiler::primitives::loops::emit_for_in_start(chunks, current, arr, i, line);
        let fcc = chunks[current].add_import("ecma:string", "fromCharCode");
        chunks[current].emit_call(fcc, 1, line);
        let piece = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], piece, line);
        get(&mut chunks[current], out, line);
        get(&mut chunks[current], piece, line);
        strings::emit_str_concat(&mut chunks[current], line);
        set(&mut chunks[current], out, line);
        vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, i, state, line);
        get(&mut chunks[current], out, line);
    }
    chunks[current].emit_else(line);
    get(&mut chunks[current], arr, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_read_utf(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    let pos = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, POS, line);
    set(&mut chunks[current], pos, line);
    field_get(&mut chunks[current], obj, DATA, line);
    get(&mut chunks[current], pos, line);
    collections::emit_get(chunks, current, line);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], pos, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    field_set_from_stack(&mut chunks[current], obj, POS, line);
    get(&mut chunks[current], value, line);
}

pub fn emit_read_line(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let saw_any = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    chunks[current].emit_string_const("", line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], saw_any, line);
    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], obj, line);
    emit_read(chunks, current, 1, line);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_i32_const(-1, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_i32_const(10, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], value, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCharCode",
        1,
        line,
    );
    strings::emit_str_concat(&mut chunks[current], line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], saw_any, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    get(&mut chunks[current], saw_any, line);
    chunks[current].emit_if_value(line);
    field_get(&mut chunks[current], obj, LINE, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    field_set_from_stack(&mut chunks[current], obj, LINE, line);
    get(&mut chunks[current], out, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

pub fn emit_get_line_number(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, LINE, line);
}

pub fn emit_string_to_char_array(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
}

pub fn emit_string_to_byte_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let text = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    set(&mut chunks[current], text, line);
    get(&mut chunks[current], text, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    set(&mut chunks[current], len, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], idx, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], idx, line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], text, line);
    get(&mut chunks[current], idx, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    get(&mut chunks[current], out, line);
}

pub fn emit_int_to_char(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:string", "fromCharCode", 1, line);
}

pub fn emit_char_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_i32_const(0, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

pub fn emit_new_filled_array(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    fill_char: bool,
    line: u32,
) {
    let len = chunks[current].alloc_scratch(1);
    if argc == 0 {
        chunks[current].emit_i32_const(0, line);
    } else if argc > 1 {
        chunks[current].emit_op(Op::DROP, line);
        set(&mut chunks[current], len, line);
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        get(&mut chunks[current], len, line);
    } else {
        set(&mut chunks[current], len, line);
        get(&mut chunks[current], len, line);
    }
    set(&mut chunks[current], len, line);
    get(&mut chunks[current], len, line);
    collections::emit_new_with_length(chunks, current, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], out, line);
    if fill_char {
        chunks[current].emit_string_const("\0", line);
    } else {
        chunks[current].emit_i32_const(0, line);
    }
    chunks[current].emit_i32_const(0, line);
    get(&mut chunks[current], len, line);
    collections::emit_fill(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], out, line);
}
