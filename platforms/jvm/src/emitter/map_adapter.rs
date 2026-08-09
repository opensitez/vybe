//! JVM `java.util.Map` adapters.

use vybe_compiler::primitives::{
    collections, errors,
    instructions::{core_wasm, host},
    ops, sorted_collection,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const CONCURRENT_MAP_KEY: &str = "__java_concurrent_map";
const IDENTITY_MAP_KEY: &str = "__java_identity_map";
const LINKED_MAP_KEY: &str = "__java_linked_map";
const IMMUTABLE_MAP_KEY: &str = "__java_immutable_map";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn drop_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
}

fn mark_bool(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], map, line);
}

pub fn emit_mark_immutable_map(chunks: &mut [Chunk], current: usize, line: u32) {
    mark_bool(chunks, current, IMMUTABLE_MAP_KEY, line);
}

fn emit_jvm_exception_throw(chunks: &mut [Chunk], current: usize, name: &str, line: u32) {
    chunks[current].emit_string_const(name, line);
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
    errors::emit_throw(&mut chunks[current], line);
}

fn emit_throw_if_immutable_map(chunks: &mut [Chunk], current: usize, map: u16, line: u32) {
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(IMMUTABLE_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_jvm_exception_throw(chunks, current, "UnsupportedOperationException", line);
    chunks[current].emit_end(line);
}

fn map_entries(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:map", "entries", 1, line);
}

pub fn emit_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    drop_args(chunks, current, argc, line);
    collections::emit_map_new(chunks, current, line);
}

pub fn emit_concurrent_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_hash_map_new(chunks, current, argc, line);
    mark_bool(chunks, current, CONCURRENT_MAP_KEY, line);
}

pub fn emit_identity_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_hash_map_new(chunks, current, argc, line);
    mark_bool(chunks, current, IDENTITY_MAP_KEY, line);
}

pub fn emit_linked_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_hash_map_new(chunks, current, argc, line);
    mark_bool(chunks, current, LINKED_MAP_KEY, line);
}

pub fn emit_sorted_map_values(chunks: &mut [Chunk], current: usize, line: u32) {
    sorted_collection::emit_sorted_map_values(chunks, current, line);
}

pub fn emit_map_entry(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_pair(chunks, current, line);
}

pub fn emit_map_of_entries(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    collections::emit_array_new(chunks, current, argc as u16, line);
    host::emit(&mut chunks[current], "ecma:map", "new", 1, line);
    emit_mark_immutable_map(chunks, current, line);
}

pub fn emit_map_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        collections::emit_map_new(chunks, current, line);
        emit_mark_immutable_map(chunks, current, line);
        return;
    }

    let argc = argc as u16;
    let base = chunks[current].alloc_scratch(argc);
    for i in (0..argc).rev() {
        set(&mut chunks[current], base + i, line);
    }

    collections::emit_array_new(chunks, current, 0, line);
    let pairs = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], pairs, line);
    let mut i = 0;
    while i + 1 < argc {
        get(&mut chunks[current], base + i, line);
        get(&mut chunks[current], base + i + 1, line);
        collections::emit_array_pair(chunks, current, line);
        let pair = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], pair, line);
        get(&mut chunks[current], pairs, line);
        get(&mut chunks[current], pair, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        i += 2;
    }

    get(&mut chunks[current], pairs, line);
    host::emit(&mut chunks[current], "ecma:map", "new", 1, line);
    emit_mark_immutable_map(chunks, current, line);
}

pub fn emit_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let numeric_bool_key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], result, line);

    get(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "wasm:js-boolean", "test", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "wasm:js-boolean", "cast", 1, line);
    set(&mut chunks[current], numeric_bool_key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], numeric_bool_key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], numeric_bool_key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], result, line);
}

pub fn emit_put(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let previous = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_throw_if_immutable_map(chunks, current, map, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], previous, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], previous, line);
}

pub fn emit_size(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:map", "size", 1, line);
}

pub fn emit_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_size(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:map", "clear", 1, line);
}

pub fn emit_key_set(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:map", "keys", 1, line);
}

pub fn emit_values(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:map", "values", 1, line);
}

pub fn emit_entry_set(chunks: &mut [Chunk], current: usize, line: u32) {
    map_entries(chunks, current, line);
}

fn emit_sorted_keys(chunks: &mut [Chunk], current: usize, line: u32) {
    sorted_collection::emit_sorted_map_key_set(chunks, current, line);
}

pub fn emit_sorted_map_key(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    emit_sorted_keys(chunks, current, line);
    let keys = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], keys, line);
    get(&mut chunks[current], keys, line);
    if last {
        get(&mut chunks[current], keys, line);
        collections::emit_len(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_SUB, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    collections::emit_get(chunks, current, line);
}

pub fn emit_sorted_map_key_set(chunks: &mut [Chunk], current: usize, line: u32) {
    sorted_collection::emit_sorted_map_key_set(chunks, current, line);
}

pub fn emit_sorted_map_bound_key(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
    let bound = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    let candidate = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], bound, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    emit_sorted_keys(chunks, current, line);
    set(&mut chunks[current], keys, line);
    get(&mut chunks[current], keys, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], found, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set(&mut chunks[current], candidate, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], keys, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], key, line);
    sorted_collection::emit_bound_condition(chunks, current, key, bound, mode, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if matches!(mode, 0 | 2) {
        get(&mut chunks[current], found, line);
        ops::emit_dyn_not(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], key, line);
        set(&mut chunks[current], candidate, line);
        chunks[current].emit_bool_const(true, line);
        set(&mut chunks[current], found, line);
        chunks[current].emit_end(line);
    } else {
        get(&mut chunks[current], key, line);
        set(&mut chunks[current], candidate, line);
        chunks[current].emit_bool_const(true, line);
        set(&mut chunks[current], found, line);
    }
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], candidate, line);
}
