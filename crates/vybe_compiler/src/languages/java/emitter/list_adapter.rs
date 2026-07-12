//! Java collection overloads composed from the shared ECMA array surface.

use crate::emitter::{
    collections,
    instructions::{core_wasm, host},
};
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

const COMPARATOR_KEY: &str = "__java_comparator";
const IMMUTABLE_MAP_KEY: &str = "__java_immutable_map";
const DESCENDING_MAP_KEY: &str = "__java_descending_map";
const ACCESS_ORDER_MAP_KEY: &str = "__java_access_order_map";
const IDENTITY_MAP_KEY: &str = "__java_identity_map";

fn emit_comparator(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    get(&mut chunks[current], value, line);
    chunks[current].emit_string_const(COMPARATOR_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn emit_sort_if_ordered(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    emit_comparator(chunks, current, value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    emit_comparator(chunks, current, value, line);
    collections::emit_runtime_helper_call(chunks, current, "__vybe_sort_with_comparator", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn emit_java_exception_throw(chunks: &mut [Chunk], current: usize, name: &str, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    crate::emitter::errors::emit_exception_new_finalize(&mut chunks[current], name, line);
    crate::emitter::errors::emit_throw(&mut chunks[current], line);
}

fn emit_throw_if_null(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_java_exception_throw(chunks, current, "NullPointerException", line);
    chunks[current].emit_end(line);
}

fn emit_mark_immutable_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(IMMUTABLE_MAP_KEY, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], map, line);
}

fn emit_mark_access_order_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(ACCESS_ORDER_MAP_KEY, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], map, line);
}

fn emit_mark_identity_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(IDENTITY_MAP_KEY, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], map, line);
}

fn emit_throw_if_immutable_map(chunks: &mut [Chunk], current: usize, map: u16, line: u32) {
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(IMMUTABLE_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_java_exception_throw(chunks, current, "UnsupportedOperationException", line);
    chunks[current].emit_end(line);
}

fn emit_canonicalize_map_key(chunks: &mut [Chunk], current: usize, map: u16, key: u16, line: u32) {
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(IDENTITY_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], key, line);
    chunks[current].emit_string_const("valueOf", line);
    host::emit(&mut chunks[current], "ecma:value", "invokeMethod", 2, line);
    set(&mut chunks[current], key, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_touch_access_order_map(
    chunks: &mut [Chunk],
    current: usize,
    map: u16,
    key: u16,
    line: u32,
) {
    let value = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(ACCESS_ORDER_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_sorted_collection_new(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    map: bool,
    line: u32,
) {
    let comparator = chunks[current].alloc_scratch(1);
    if argc == 0 {
        chunks[current].emit_op(Op::NULL, line);
    }
    set(&mut chunks[current], comparator, line);
    if map {
        collections::emit_map_new(chunks, current, line);
    } else {
        collections::emit_array_new(chunks, current, 0, line);
    }
    let collection = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], collection, line);
    get(&mut chunks[current], collection, line);
    chunks[current].emit_string_const(COMPARATOR_KEY, line);
    get(&mut chunks[current], comparator, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], collection, line);
}

pub fn emit_arrays_as_list(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc != 1 {
        collections::emit_array_new(chunks, current, argc as u16, line);
        return;
    }

    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    let length = chunks[current].add_import("ecma:array", "length");
    chunks[current].emit_call(length, 1, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

pub fn emit_n_copies(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    collections::emit_new_with_length(chunks, current, line);
    get(&mut chunks[current], value, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_i32_const(i32::MAX, line);
    collections::emit_fill(chunks, current, line);
}

pub fn emit_double_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    host::emit(&mut chunks[current], "ecma:number", "toFixed", 2, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:number", "toString", 1, line);
    chunks[current].emit_end(line);
}

pub fn emit_map_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], result, line);
    emit_touch_access_order_map(chunks, current, map, key, line);
    get(&mut chunks[current], result, line);
}

pub fn emit_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    collections::emit_map_new(chunks, current, line);
}

pub fn emit_identity_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    collections::emit_map_new(chunks, current, line);
    emit_mark_identity_map(chunks, current, line);
}

pub fn emit_linked_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        collections::emit_map_new(chunks, current, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc as u16).rev() {
        set(&mut chunks[current], base + i, line);
    }

    collections::emit_map_new(chunks, current, line);
    if argc >= 3 {
        let map = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], map, line);
        get(&mut chunks[current], base + 2, line);
        crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], map, line);
        emit_mark_access_order_map(chunks, current, line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], map, line);
        chunks[current].emit_end(line);
    }
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
    let argc = argc as u16;
    if argc == 0 {
        collections::emit_map_new(chunks, current, line);
        emit_mark_immutable_map(chunks, current, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc);
    for i in (0..argc).rev() {
        set(&mut chunks[current], base + i, line);
    }

    collections::emit_array_new(chunks, current, 0, line);
    let pairs = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], pairs, line);

    let mut i = 0;
    while i + 1 < argc {
        emit_throw_if_null(chunks, current, base + i, line);
        emit_throw_if_null(chunks, current, base + i + 1, line);
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

pub fn emit_map_put(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let previous = chunks[current].alloc_scratch(1);
    let had_key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    emit_throw_if_immutable_map(chunks, current, map, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    set(&mut chunks[current], had_key, line);
    get(&mut chunks[current], had_key, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], previous, line);
    get(&mut chunks[current], had_key, line);
    chunks[current].emit_if_value(line);
    emit_touch_access_order_map(chunks, current, map, key, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], previous, line);
}

pub fn emit_map_put_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    set(&mut chunks[current], target, line);
    emit_throw_if_immutable_map(chunks, current, target, line);

    get(&mut chunks[current], source, line);
    host::emit(&mut chunks[current], "ecma:map", "entries", 1, line);
    set(&mut chunks[current], entries, line);
    get(&mut chunks[current], entries, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    let key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], key, line);
    emit_canonicalize_map_key(chunks, current, target, key, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], target, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], pair, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_map_get_or_default(chunks: &mut [Chunk], current: usize, line: u32) {
    let default = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let had_key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], default, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    set(&mut chunks[current], had_key, line);
    get(&mut chunks[current], had_key, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], default, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], result, line);
    get(&mut chunks[current], had_key, line);
    chunks[current].emit_if_value(line);
    emit_touch_access_order_map(chunks, current, map, key, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], result, line);
}

pub fn emit_map_contains_key(chunks: &mut [Chunk], current: usize, line: u32) {
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
}

pub fn emit_map_put_if_absent(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

pub fn emit_map_compute_if_absent(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fn_slot, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], fn_slot, line);
    host::emit(
        &mut chunks[current],
        "ecma:map",
        "getOrInsertComputed",
        3,
        line,
    );
}

pub fn emit_map_remove(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 3 {
        let expected = chunks[current].alloc_scratch(1);
        let key = chunks[current].alloc_scratch(1);
        let map = chunks[current].alloc_scratch(1);
        let current_value = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], expected, line);
        set(&mut chunks[current], key, line);
        set(&mut chunks[current], map, line);
        emit_canonicalize_map_key(chunks, current, map, key, line);
        get(&mut chunks[current], map, line);
        get(&mut chunks[current], key, line);
        host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
        crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], map, line);
        get(&mut chunks[current], key, line);
        host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
        set(&mut chunks[current], current_value, line);
        get(&mut chunks[current], current_value, line);
        get(&mut chunks[current], expected, line);
        crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
        crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], map, line);
        get(&mut chunks[current], key, line);
        host::emit(&mut chunks[current], "ecma:map", "delete", 2, line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
        return;
    }

    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let previous = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], previous, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], previous, line);
}

fn emit_null_check(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    get(&mut chunks[current], slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
}

pub fn emit_map_replace(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 4 {
        let new_value = chunks[current].alloc_scratch(1);
        let old_value = chunks[current].alloc_scratch(1);
        let key = chunks[current].alloc_scratch(1);
        let map = chunks[current].alloc_scratch(1);
        let current_value = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], new_value, line);
        set(&mut chunks[current], old_value, line);
        set(&mut chunks[current], key, line);
        set(&mut chunks[current], map, line);
        emit_canonicalize_map_key(chunks, current, map, key, line);
        get(&mut chunks[current], map, line);
        get(&mut chunks[current], key, line);
        host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], map, line);
        get(&mut chunks[current], key, line);
        host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
        set(&mut chunks[current], current_value, line);
        get(&mut chunks[current], current_value, line);
        get(&mut chunks[current], old_value, line);
        crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
        crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], map, line);
        get(&mut chunks[current], key, line);
        get(&mut chunks[current], new_value, line);
        host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
        return;
    }

    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let previous = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    set(&mut chunks[current], previous, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], previous, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

pub fn emit_map_compute(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let old_value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fn_slot, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);

    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], old_value, line);

    get(&mut chunks[current], fn_slot, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], old_value, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    set(&mut chunks[current], result, line);

    emit_null_check(chunks, current, result, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    emit_touch_access_order_map(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], result, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], result, line);
}

pub fn emit_map_compute_if_present(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let old_value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fn_slot, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);

    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    set(&mut chunks[current], old_value, line);
    get(&mut chunks[current], fn_slot, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], old_value, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    set(&mut chunks[current], result, line);
    emit_null_check(chunks, current, result, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    emit_touch_access_order_map(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], result, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], result, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

pub fn emit_map_merge(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let old_value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let had_key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fn_slot, line);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);

    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    set(&mut chunks[current], had_key, line);
    get(&mut chunks[current], had_key, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    set(&mut chunks[current], old_value, line);
    get(&mut chunks[current], fn_slot, line);
    get(&mut chunks[current], old_value, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], result, line);

    emit_null_check(chunks, current, result, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], had_key, line);
    chunks[current].emit_if_value(line);
    emit_touch_access_order_map(chunks, current, map, key, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], result, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], result, line);
}

pub fn emit_map_replace_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fn_slot, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "entries", 1, line);
    set(&mut chunks[current], entries, line);
    get(&mut chunks[current], entries, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], key, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);

    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], fn_slot, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_map_for_each(chunks: &mut [Chunk], current: usize, line: u32) {
    let callback = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], callback, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "entries", 1, line);
    set(&mut chunks[current], entries, line);
    get(&mut chunks[current], entries, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], callback, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_map_clone(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let cloned = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    get(&mut chunks[current], source, line);
    host::emit(&mut chunks[current], "ecma:map", "entries", 1, line);
    host::emit(&mut chunks[current], "ecma:map", "new", 1, line);
    set(&mut chunks[current], cloned, line);

    get(&mut chunks[current], source, line);
    chunks[current].emit_string_const(IDENTITY_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], cloned, line);
    emit_mark_identity_map(chunks, current, line);
    set(&mut chunks[current], cloned, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], source, line);
    chunks[current].emit_string_const(ACCESS_ORDER_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], cloned, line);
    emit_mark_access_order_map(chunks, current, line);
    set(&mut chunks[current], cloned, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], cloned, line);
}

pub fn emit_map_key_set_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "delete", 2, line);
}

pub fn emit_map_entry_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "entries", 1, line);
    set(&mut chunks[current], entries, line);
    get(&mut chunks[current], entries, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], pair, line);
    get(&mut chunks[current], map, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], entries, line);
}

pub fn emit_entry_set_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let new_value = chunks[current].alloc_scratch(1);
    let entry = chunks[current].alloc_scratch(1);
    let old_value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], new_value, line);
    set(&mut chunks[current], entry, line);

    get(&mut chunks[current], entry, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], key, line);
    get(&mut chunks[current], entry, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], old_value, line);
    get(&mut chunks[current], entry, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], map, line);

    get(&mut chunks[current], entry, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    get(&mut chunks[current], new_value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], map, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], new_value, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], old_value, line);
}

pub fn emit_iterator_next(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("next", line);
    host::emit(&mut chunks[current], "ecma:value", "invokeMethod", 2, line);
    chunks[current].emit_string_const("value", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

pub fn emit_iterator_has_next(chunks: &mut [Chunk], current: usize, line: u32) {
    let iterator = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get(&mut chunks[current], iterator, line);
    chunks[current].emit_string_const("__index", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], iterator, line);
    collections::emit_len(chunks, current, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_map_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], other, line);
    set(&mut chunks[current], map, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], result, line);

    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "size", 1, line);
    get(&mut chunks[current], other, line);
    host::emit(&mut chunks[current], "ecma:map", "size", 1, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "entries", 1, line);
    set(&mut chunks[current], entries, line);
    get(&mut chunks[current], entries, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], result, line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], key, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);

    get(&mut chunks[current], other, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], other, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    get(&mut chunks[current], value, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_end(line);
    get(&mut chunks[current], result, line);
}

pub fn emit_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    if argc == 2 {
        let list = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], list, line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], value, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
    } else {
        let index = chunks[current].alloc_scratch(1);
        let list = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], index, line);
        set(&mut chunks[current], list, line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], index, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        get(&mut chunks[current], value, line);
        collections::emit_insert(chunks, current, line);
    }
}

pub fn emit_sorted_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_sort_if_ordered(chunks, current, list, line);
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_queue_poll(chunks: &mut [Chunk], current: usize, line: u32) {
    let queue = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], queue, line);
    emit_sort_if_ordered(chunks, current, queue, line);
    get(&mut chunks[current], queue, line);
    collections::emit_shift(chunks, current, line);
}

pub fn emit_sorted_end(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let collection = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], collection, line);
    emit_sort_if_ordered(chunks, current, collection, line);
    get(&mut chunks[current], collection, line);
    if last {
        get(&mut chunks[current], collection, line);
        collections::emit_len(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_SUB, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    collections::emit_get(chunks, current, line);
}

pub fn emit_sorted_map_key(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let use_last = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    chunks[current].emit_bool_const(last, line);
    set(&mut chunks[current], use_last, line);
    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "keys", 1, line);
    set(&mut chunks[current], keys, line);
    emit_comparator(chunks, current, map, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], keys, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], keys, line);
    emit_comparator(chunks, current, map, line);
    collections::emit_runtime_helper_call(chunks, current, "__vybe_sort_with_comparator", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], keys, line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(DESCENDING_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], use_last, line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    set(&mut chunks[current], use_last, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], use_last, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], keys, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_sorted_map_key_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "keys", 1, line);
    set(&mut chunks[current], keys, line);
    emit_comparator(chunks, current, map, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], keys, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], keys, line);
    emit_comparator(chunks, current, map, line);
    collections::emit_runtime_helper_call(chunks, current, "__vybe_sort_with_comparator", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(DESCENDING_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], keys, line);
    collections::emit_reverse(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], keys, line);
}

pub fn emit_sorted_map_descending_key_set(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_sorted_map_key_set(chunks, current, line);
    collections::emit_reverse(chunks, current, line);
}

pub fn emit_sorted_map_descending_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(DESCENDING_MAP_KEY, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], map, line);
}

pub fn emit_sorted_map_end_entry(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    emit_sorted_map_key(chunks, current, last, line);
    set(&mut chunks[current], key, line);
    emit_map_entry_from_key(chunks, current, map, key, line);
}

pub fn emit_sorted_map_bound_entry(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
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
    host::emit(&mut chunks[current], "ecma:map", "keys", 1, line);
    set(&mut chunks[current], keys, line);
    get(&mut chunks[current], keys, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], keys, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], found, line);
    chunks[current].emit_op(Op::NULL, line);
    set(&mut chunks[current], candidate, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], keys, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], key, line);

    emit_bound_condition(chunks, current, key, bound, mode, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if matches!(mode, 0 | 2) {
        get(&mut chunks[current], found, line);
        crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
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
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    emit_map_entry_from_key(chunks, current, map, candidate, line);
}

pub fn emit_sorted_map_bound_key(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
    emit_sorted_map_bound_entry(chunks, current, mode, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
}

pub fn emit_sorted_map_poll_entry(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let entry = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);

    get(&mut chunks[current], map, line);
    emit_sorted_map_key(chunks, current, last, line);
    set(&mut chunks[current], key, line);

    emit_map_entry_from_key(chunks, current, map, key, line);
    set(&mut chunks[current], entry, line);

    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], entry, line);
}

fn emit_bound_condition(
    chunks: &mut [Chunk],
    current: usize,
    key: u16,
    bound: u16,
    mode: u8,
    line: u32,
) {
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], bound, line);
    match mode {
        0 => {
            crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
            crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
        }
        1 => {
            crate::emitter::ops::emit_dyn_gt(&mut chunks[current], line);
            crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
        }
        2 => crate::emitter::ops::emit_dyn_gt(&mut chunks[current], line),
        _ => crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line),
    }
}

fn emit_map_entry_from_key(chunks: &mut [Chunk], current: usize, map: u16, key: u16, line: u32) {
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    get(&mut chunks[current], map, line);
    collections::emit_array_new(chunks, current, 3, line);
}

pub fn emit_map_range_view(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
    let upper = if mode == 0 || mode == 1 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let lower = if mode == 0 || mode == 2 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let map = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    collections::emit_map_new(chunks, current, line);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "entries", 1, line);
    set(&mut chunks[current], entries, line);
    get(&mut chunks[current], entries, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], key, line);
    emit_range_condition(chunks, current, key, lower, upper, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], out, line);
}

fn emit_range_condition(
    chunks: &mut [Chunk],
    current: usize,
    key: u16,
    lower: Option<u16>,
    upper: Option<u16>,
    line: u32,
) {
    if let Some(lower) = lower {
        get(&mut chunks[current], key, line);
        get(&mut chunks[current], lower, line);
        crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
        crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
        crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        emit_upper_condition(chunks, current, key, upper, line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
    } else {
        emit_upper_condition(chunks, current, key, upper, line);
    }
}

fn emit_upper_condition(
    chunks: &mut [Chunk],
    current: usize,
    key: u16,
    upper: Option<u16>,
    line: u32,
) {
    if let Some(upper) = upper {
        get(&mut chunks[current], key, line);
        get(&mut chunks[current], upper, line);
        crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
        crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_bool_const(true, line);
    }
}

pub fn emit_add_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    let index = if argc == 3 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let list = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);

    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    set(&mut chunks[current], changed, line);

    get(&mut chunks[current], list, line);
    if let Some(index) = index {
        get(&mut chunks[current], index, line);
    } else {
        get(&mut chunks[current], list, line);
        collections::emit_len(chunks, current, line);
    }
    get(&mut chunks[current], source, line);
    collections::emit_insert_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], changed, line);
}

pub fn emit_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let previous = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], index, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], previous, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], previous, line);
}

pub fn emit_sort(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 1 {
        collections::emit_sort(chunks, current, line);
        return;
    }
    let comparator = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], comparator, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], comparator, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], comparator, line);
    collections::emit_runtime_helper_call(chunks, current, "__vybe_sort_with_comparator", 2, line);
    chunks[current].emit_end(line);
}

pub fn emit_remove_all(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_filter_members(chunks, current, line, false);
}

pub fn emit_retain_all(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_filter_members(chunks, current, line, true);
}

fn emit_filter_members(chunks: &mut [Chunk], current: usize, line: u32, retain: bool) {
    let members = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let snapshot = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let length = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], members, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    collections::emit_clone(chunks, current, line);
    set(&mut chunks[current], snapshot, line);
    get(&mut chunks[current], snapshot, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], length, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);

    let outer = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], length, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], snapshot, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], members, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    if retain {
        chunks[current].emit_op(Op::I32_EQZ, line);
    }
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_remove_value(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], changed, line);
}
