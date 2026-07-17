//! Java collection overloads composed from the shared ECMA array surface.

use std::sync::Arc;
use vybe_bytecode::Chunk;
use vybe_bytecode::Value;
use vybe_bytecode::opcode::Op;
use vybe_emitter::{
    collections,
    functions::create_function_chunk,
    heap,
    instructions::{core_wasm, host},
};

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn set_object_prop_from_local(
    chunks: &mut [Chunk],
    current: usize,
    object: u16,
    key: &str,
    value: u16,
    line: u32,
) {
    get(&mut chunks[current], object, line);
    chunks[current].emit_string_const(key, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_object_prop_i32(
    chunks: &mut [Chunk],
    current: usize,
    object: u16,
    key: &str,
    value: i32,
    line: u32,
) {
    get(&mut chunks[current], object, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_i32_const(value, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn get_object_prop(chunks: &mut [Chunk], current: usize, object: u16, key: &str, line: u32) {
    get(&mut chunks[current], object, line);
    chunks[current].emit_string_const(key, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn get_iterator_list(chunks: &mut [Chunk], current: usize, iterator: u16, line: u32) {
    get_object_prop(chunks, current, iterator, "__list", line);
}

fn emit_is_sublist(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(SUBLIST_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn emit_sublist_parent(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get_object_prop(chunks, current, list, SUBLIST_PARENT_KEY, line);
}

fn emit_sublist_from(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get_object_prop(chunks, current, list, SUBLIST_FROM_KEY, line);
}

fn emit_sublist_to(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get_object_prop(chunks, current, list, SUBLIST_TO_KEY, line);
}

fn emit_sublist_absolute_index(
    chunks: &mut [Chunk],
    current: usize,
    list: u16,
    index: u16,
    line: u32,
) {
    emit_sublist_from(chunks, current, list, line);
    get(&mut chunks[current], index, line);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
}

fn emit_sublist_increment_to(
    chunks: &mut [Chunk],
    current: usize,
    list: u16,
    delta: i32,
    line: u32,
) {
    let next = chunks[current].alloc_scratch(1);
    emit_sublist_to(chunks, current, list, line);
    core_wasm::i32_const(&mut chunks[current], line, delta);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], next, line);
    set_object_prop_from_local(chunks, current, list, SUBLIST_TO_KEY, next, line);
}

const COMPARATOR_KEY: &str = "__java_comparator";
const IMMUTABLE_MAP_KEY: &str = "__java_immutable_map";
const IMMUTABLE_LIST_KEY: &str = "__java_immutable_list";
const SET_COLLECTION_KEY: &str = "__java_set_collection";
const SUBLIST_KEY: &str = "__java_sublist";
const SUBLIST_PARENT_KEY: &str = "__java_sublist_parent";
const SUBLIST_FROM_KEY: &str = "__java_sublist_from";
const SUBLIST_TO_KEY: &str = "__java_sublist_to";
const DESCENDING_MAP_KEY: &str = "__java_descending_map";
const DESCENDING_SET_KEY: &str = "__java_descending_set";
const ACCESS_ORDER_MAP_KEY: &str = "__java_access_order_map";
const IDENTITY_MAP_KEY: &str = "__java_identity_map";
const CONCURRENT_MAP_KEY: &str = "__java_concurrent_map";
const SEMAPHORE_PERMITS_KEY: &str = "__java_semaphore_permits";
const SEMAPHORE_FAIR_KEY: &str = "__java_semaphore_fair";
const SEMAPHORE_QUEUED_KEY: &str = "__java_semaphore_queued";
const SEMAPHORE_CELLS_KEY: &str = "__java_semaphore_cells";
const JAVA_THREAD_HANDLE_KEY: &str = "__java_thread_handle";
const JAVA_THREAD_ALIVE_KEY: &str = "alive";
const IDENTITY_KEY_PROP: &str = "__java_identity_key";
const IDENTITY_NEXT_KEY: &str = "__java_identity_next";
const VECTOR_CAPACITY_KEY: &str = "__java_vector_capacity";
const ENUM_ITEMS_KEY: &str = "__java_enum_items";
const ENUM_INDEX_KEY: &str = "__java_enum_index";

fn emit_comparator(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    get(&mut chunks[current], value, line);
    chunks[current].emit_string_const(COMPARATOR_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn emit_sort_if_ordered(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    let comparator = chunks[current].alloc_scratch(1);
    emit_comparator(chunks, current, value, line);
    set(&mut chunks[current], comparator, line);
    get(&mut chunks[current], comparator, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], comparator, line);
    collections::emit_sort_with_comparator(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    collections::emit_sort(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_string_const(DESCENDING_SET_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    collections::emit_reverse(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn emit_java_exception_throw(chunks: &mut [Chunk], current: usize, name: &str, line: u32) {
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    vybe_emitter::errors::emit_exception_new_finalize(&mut chunks[current], name, line);
    vybe_emitter::errors::emit_throw(&mut chunks[current], line);
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

pub fn emit_mark_immutable_list(chunks: &mut [Chunk], current: usize, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(IMMUTABLE_LIST_KEY, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], list, line);
}

pub fn emit_mark_set_collection(chunks: &mut [Chunk], current: usize, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(SET_COLLECTION_KEY, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], list, line);
}

pub fn emit_set_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    collections::emit_array_new(chunks, current, argc as u16, line);
    let list = chunks[current].alloc_scratch(1);
    let outer_index = chunks[current].alloc_scratch(1);
    let inner_index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let outer_value = chunks[current].alloc_scratch(1);
    let inner_value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], outer_index, line);

    let outer = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], outer_index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], outer_index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], outer_value, line);
    get(&mut chunks[current], outer_index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], inner_index, line);

    let inner = chunks[current].emit_block(line);
    let (inner_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], inner_index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], inner_index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], inner_value, line);
    get(&mut chunks[current], outer_value, line);
    get(&mut chunks[current], inner_value, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_java_exception_throw(chunks, current, "IllegalArgumentException", line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], inner_index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], inner_index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inner_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner);

    get(&mut chunks[current], outer_index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], outer_index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], list, line);
    emit_mark_set_collection(chunks, current, line);
    emit_mark_immutable_list(chunks, current, line);
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
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_java_exception_throw(chunks, current, "UnsupportedOperationException", line);
    chunks[current].emit_end(line);
}

fn emit_throw_if_immutable_list(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(IMMUTABLE_LIST_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_java_exception_throw(chunks, current, "UnsupportedOperationException", line);
    chunks[current].emit_end(line);
}

fn emit_is_set_collection(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(SET_COLLECTION_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn emit_canonicalize_map_key(chunks: &mut [Chunk], current: usize, map: u16, key: u16, line: u32) {
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(IDENTITY_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    get(&mut chunks[current], key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    let identity = chunks[current].alloc_scratch(1);
    get_object_prop(chunks, current, key, IDENTITY_KEY_PROP, line);
    set(&mut chunks[current], identity, line);
    get(&mut chunks[current], identity, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("undefined", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get_object_prop(chunks, current, map, IDENTITY_NEXT_KEY, line);
    set(&mut chunks[current], identity, line);
    get(&mut chunks[current], identity, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("undefined", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    set(&mut chunks[current], identity, line);
    chunks[current].emit_end(line);
    set_object_prop_from_local(chunks, current, key, IDENTITY_KEY_PROP, identity, line);
    get(&mut chunks[current], identity, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    let next = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], next, line);
    set_object_prop_from_local(chunks, current, map, IDENTITY_NEXT_KEY, next, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], identity, line);
    set(&mut chunks[current], key, line);
    chunks[current].emit_end(line);
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

fn emit_identity_map_flag(chunks: &mut [Chunk], current: usize, map: u16, line: u32) {
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(IDENTITY_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn emit_pack_identity_map_value(
    chunks: &mut [Chunk],
    current: usize,
    map: u16,
    original_key: u16,
    value: u16,
    line: u32,
) {
    emit_identity_map_flag(chunks, current, map, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], original_key, line);
    get(&mut chunks[current], value, line);
    collections::emit_array_pair(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

fn emit_unwrap_identity_map_value(
    chunks: &mut [Chunk],
    current: usize,
    map: u16,
    value: u16,
    line: u32,
) {
    emit_identity_map_flag(chunks, current, map, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

fn emit_identity_entry_key(chunks: &mut [Chunk], current: usize, map: u16, pair: u16, line: u32) {
    emit_identity_map_flag(chunks, current, map, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_identity_entry_value(chunks: &mut [Chunk], current: usize, map: u16, pair: u16, line: u32) {
    emit_identity_map_flag(chunks, current, map, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    collections::emit_get(chunks, current, line);
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
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    if !map {
        get(&mut chunks[current], collection, line);
        chunks[current].emit_string_const(SET_COLLECTION_KEY, line);
        chunks[current].emit_bool_const(true, line);
        host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    get(&mut chunks[current], collection, line);
}

pub fn emit_priority_queue_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let comparator = chunks[current].alloc_scratch(1);
    if argc == 0 {
        chunks[current].emit_op(Op::NULL, line);
    }
    set(&mut chunks[current], comparator, line);
    collections::emit_array_new(chunks, current, 0, line);
    let collection = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], collection, line);
    get(&mut chunks[current], collection, line);
    chunks[current].emit_string_const(COMPARATOR_KEY, line);
    get(&mut chunks[current], comparator, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], collection, line);
}

pub fn emit_vector_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let capacity = chunks[current].alloc_scratch(1);
    if argc == 0 {
        core_wasm::i32_const(&mut chunks[current], line, 10);
    }
    set(&mut chunks[current], capacity, line);
    collections::emit_array_new(chunks, current, 0, line);
    let vector = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], vector, line);
    set_object_prop_from_local(chunks, current, vector, VECTOR_CAPACITY_KEY, capacity, line);
    get(&mut chunks[current], vector, line);
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
    emit_mark_immutable_list(chunks, current, line);
}

pub fn emit_sub_list(chunks: &mut [Chunk], current: usize, line: u32) {
    let to = chunks[current].alloc_scratch(1);
    let from = chunks[current].alloc_scratch(1);
    let parent = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], to, line);
    set(&mut chunks[current], from, line);
    set(&mut chunks[current], parent, line);

    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    let view = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], view, line);
    get(&mut chunks[current], view, line);
    chunks[current].emit_string_const(SUBLIST_KEY, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    set_object_prop_from_local(chunks, current, view, SUBLIST_PARENT_KEY, parent, line);
    set_object_prop_from_local(chunks, current, view, SUBLIST_FROM_KEY, from, line);
    set_object_prop_from_local(chunks, current, view, SUBLIST_TO_KEY, to, line);
    get(&mut chunks[current], view, line);
}

pub fn emit_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let index = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], index, line);
    set(&mut chunks[current], list, line);
    emit_is_sublist(chunks, current, list, line);
    chunks[current].emit_if_value(line);
    emit_sublist_parent(chunks, current, list, line);
    emit_sublist_absolute_index(chunks, current, list, index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    emit_is_sublist(chunks, current, list, line);
    chunks[current].emit_if_value(line);
    emit_sublist_to(chunks, current, list, line);
    emit_sublist_from(chunks, current, list, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_collection_frequency(chunks: &mut [Chunk], current: usize, line: u32) {
    let target = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let count = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], target, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], count, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], target, line);
    vybe_emitter::object::emit_equals(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], count, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], count, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], count, line);
}

pub fn emit_collection_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let dest = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    set(&mut chunks[current], dest, line);
    emit_throw_if_immutable_list(chunks, current, dest, line);
    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], source, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], dest, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_collection_disjoint(chunks: &mut [Chunk], current: usize, line: u32) {
    let right = chunks[current].alloc_scratch(1);
    let left = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], right, line);
    set(&mut chunks[current], left, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], result, line);
    get(&mut chunks[current], left, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], result, line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], left, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], right, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], result, line);
}

pub fn emit_collection_extreme(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    min: bool,
    line: u32,
) {
    let comparator = chunks[current].alloc_scratch(1);
    if argc == 2 {
        set(&mut chunks[current], comparator, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        set(&mut chunks[current], comparator, line);
    }
    let list = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let best = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);

    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    get(&mut chunks[current], list, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], best, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], comparator, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], best, line);
    if min {
        vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    } else {
        vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    }
    chunks[current].emit_else(line);
    get(&mut chunks[current], comparator, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], best, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    if min {
        vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    } else {
        vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    }
    chunks[current].emit_end(line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    set(&mut chunks[current], best, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], best, line);
}

pub fn emit_double_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    let numeric_bool_key = chunks[current].alloc_scratch(1);
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
    get(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "wasm:js-boolean", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    get(&mut chunks[current], numeric_bool_key, line);
    set(&mut chunks[current], key, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    emit_touch_access_order_map(chunks, current, map, key, line);
    get(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    emit_unwrap_identity_map_value(chunks, current, map, result, line);
    chunks[current].emit_end(line);
}

pub fn emit_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    collections::emit_map_new(chunks, current, line);
}

pub fn emit_concurrent_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_hash_map_new(chunks, current, argc, line);
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(CONCURRENT_MAP_KEY, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], map, line);
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
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    let original_key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    set(&mut chunks[current], original_key, line);
    get_object_prop(chunks, current, map, CONCURRENT_MAP_KEY, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_throw_if_null(chunks, current, key, line);
    chunks[current].emit_end(line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    emit_throw_if_immutable_map(chunks, current, map, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    emit_pack_identity_map_value(chunks, current, map, original_key, value, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], previous, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    emit_unwrap_identity_map_value(chunks, current, map, previous, line);
    chunks[current].emit_end(line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    get(&mut chunks[current], had_key, line);
    chunks[current].emit_if_value(line);
    emit_unwrap_identity_map_value(chunks, current, map, result, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], result, line);
    chunks[current].emit_end(line);
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

pub fn emit_map_contains_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let expected = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], expected, line);
    set(&mut chunks[current], map, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], found, line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], found, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_br(1, line);
    chunks[current].emit_end(line);
    emit_identity_entry_value(chunks, current, map, pair, line);
    get(&mut chunks[current], expected, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], found, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], found, line);
}

pub fn emit_map_put_if_absent(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let original_key = chunks[current].alloc_scratch(1);
    let previous = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    set(&mut chunks[current], original_key, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    set(&mut chunks[current], previous, line);
    emit_unwrap_identity_map_value(chunks, current, map, previous, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    emit_pack_identity_map_value(chunks, current, map, original_key, value, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

pub fn emit_map_compute_if_absent(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let original_key = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fn_slot, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    set(&mut chunks[current], original_key, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    set(&mut chunks[current], result, line);
    emit_unwrap_identity_map_value(chunks, current, map, result, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], fn_slot, line);
    get(&mut chunks[current], original_key, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    set(&mut chunks[current], result, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    emit_pack_identity_map_value(chunks, current, map, original_key, result, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], result, line);
    chunks[current].emit_end(line);
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
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], map, line);
        get(&mut chunks[current], key, line);
        host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
        set(&mut chunks[current], current_value, line);
        emit_unwrap_identity_map_value(chunks, current, map, current_value, line);
        get(&mut chunks[current], expected, line);
        vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    emit_unwrap_identity_map_value(chunks, current, map, previous, line);
    chunks[current].emit_end(line);
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
        emit_unwrap_identity_map_value(chunks, current, map, current_value, line);
        get(&mut chunks[current], old_value, line);
        vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], map, line);
        get(&mut chunks[current], key, line);
        emit_pack_identity_map_value(chunks, current, map, key, new_value, line);
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
    emit_pack_identity_map_value(chunks, current, map, key, value, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_unwrap_identity_map_value(chunks, current, map, previous, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_end(line);
}

pub fn emit_map_compute(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let original_key = chunks[current].alloc_scratch(1);
    let old_value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fn_slot, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    set(&mut chunks[current], original_key, line);
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
    get(&mut chunks[current], original_key, line);
    get(&mut chunks[current], old_value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    emit_unwrap_identity_map_value(chunks, current, map, old_value, line);
    chunks[current].emit_end(line);
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
    emit_pack_identity_map_value(chunks, current, map, original_key, result, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], result, line);
}

pub fn emit_map_compute_if_present(chunks: &mut [Chunk], current: usize, line: u32) {
    let fn_slot = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let original_key = chunks[current].alloc_scratch(1);
    let old_value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fn_slot, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    set(&mut chunks[current], original_key, line);
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
    get(&mut chunks[current], original_key, line);
    emit_unwrap_identity_map_value(chunks, current, map, old_value, line);
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
    emit_pack_identity_map_value(chunks, current, map, original_key, result, line);
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
    let original_key = chunks[current].alloc_scratch(1);
    let old_value = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let had_key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], fn_slot, line);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    set(&mut chunks[current], original_key, line);
    emit_canonicalize_map_key(chunks, current, map, key, line);

    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    set(&mut chunks[current], had_key, line);
    get(&mut chunks[current], had_key, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    set(&mut chunks[current], old_value, line);
    get(&mut chunks[current], fn_slot, line);
    emit_unwrap_identity_map_value(chunks, current, map, old_value, line);
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
    emit_pack_identity_map_value(chunks, current, map, original_key, result, line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], callback, line);
    emit_identity_entry_key(chunks, current, map, pair, line);
    emit_identity_entry_value(chunks, current, map, pair, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_concurrent_items(chunks: &mut [Chunk], current: usize, map: u16, mode: u8, line: u32) {
    get(&mut chunks[current], map, line);
    match mode {
        0 => host::emit(&mut chunks[current], "ecma:map", "keys", 1, line),
        1 => host::emit(&mut chunks[current], "ecma:map", "values", 1, line),
        _ => host::emit(&mut chunks[current], "ecma:map", "entries", 1, line),
    }
}

pub fn emit_concurrent_for_each(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
    let callback = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let items = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], callback, line);
    chunks[current].emit_op(Op::DROP, line);
    set(&mut chunks[current], map, line);
    emit_concurrent_items(chunks, current, map, mode, line);
    set(&mut chunks[current], items, line);
    get(&mut chunks[current], items, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], items, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], item, line);
    get(&mut chunks[current], callback, line);
    get(&mut chunks[current], item, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_concurrent_reduce(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
    let reducer = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let items = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let acc = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], reducer, line);
    chunks[current].emit_op(Op::DROP, line);
    set(&mut chunks[current], map, line);
    emit_concurrent_items(chunks, current, map, mode, line);
    set(&mut chunks[current], items, line);
    get(&mut chunks[current], items, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    get(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], items, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], acc, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    set(&mut chunks[current], index, line);
    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], reducer, line);
    get(&mut chunks[current], acc, line);
    get(&mut chunks[current], items, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    set(&mut chunks[current], acc, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], acc, line);
    chunks[current].emit_end(line);
}

pub fn emit_concurrent_search(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
    let searcher = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    let items = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], searcher, line);
    chunks[current].emit_op(Op::DROP, line);
    set(&mut chunks[current], map, line);
    emit_concurrent_items(chunks, current, map, mode, line);
    set(&mut chunks[current], items, line);
    get(&mut chunks[current], items, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::NULL, line);
    set(&mut chunks[current], found, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);
    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], searcher, line);
    get(&mut chunks[current], items, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    set(&mut chunks[current], found, line);
    get(&mut chunks[current], found, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], found, line);
}

pub fn emit_semaphore_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let fair = chunks[current].alloc_scratch(1);
    let permits = chunks[current].alloc_scratch(1);
    match argc {
        0 => {
            chunks[current].emit_bool_const(false, line);
            set(&mut chunks[current], fair, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            set(&mut chunks[current], permits, line);
        }
        1 => {
            chunks[current].emit_bool_const(false, line);
            set(&mut chunks[current], fair, line);
            set(&mut chunks[current], permits, line);
        }
        _ => {
            set(&mut chunks[current], fair, line);
            set(&mut chunks[current], permits, line);
        }
    }
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    let sem = chunks[current].alloc_scratch(1);
    let cells = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sem, line);

    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    set(&mut chunks[current], cells, line);
    get(&mut chunks[current], cells, line);
    chunks[current].emit_string_const("__shared_int32_len", line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], cells, line);
    chunks[current].emit_string_const("0", line);
    get(&mut chunks[current], permits, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], cells, line);
    chunks[current].emit_string_const("1", line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    set_object_prop_from_local(chunks, current, sem, SEMAPHORE_CELLS_KEY, cells, line);
    set_object_prop_from_local(chunks, current, sem, SEMAPHORE_PERMITS_KEY, permits, line);
    set_object_prop_from_local(chunks, current, sem, SEMAPHORE_FAIR_KEY, fair, line);
    set_object_prop_i32(chunks, current, sem, SEMAPHORE_QUEUED_KEY, 0, line);
    get(&mut chunks[current], sem, line);
}

pub fn emit_semaphore_available(chunks: &mut [Chunk], current: usize, line: u32) {
    let sem = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sem, line);
    emit_semaphore_get_permits(chunks, current, sem, line);
}

fn emit_semaphore_take_args(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    timed: bool,
    line: u32,
) -> (u16, u16) {
    if timed && argc >= 4 {
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op(Op::DROP, line);
    } else if timed && argc == 3 {
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op(Op::DROP, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
    } else if argc <= 1 {
        core_wasm::i32_const(&mut chunks[current], line, 1);
    }
    let permits = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], permits, line);
    let sem = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sem, line);
    (sem, permits)
}

fn emit_semaphore_get_permits(chunks: &mut [Chunk], current: usize, sem: u16, line: u32) {
    get_object_prop(chunks, current, sem, SEMAPHORE_CELLS_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:atomics", "load", 2, line);
}

fn emit_semaphore_set_permits_from_top(chunks: &mut [Chunk], current: usize, sem: u16, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get_object_prop(chunks, current, sem, SEMAPHORE_CELLS_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:atomics", "store", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    set_object_prop_from_local(chunks, current, sem, SEMAPHORE_PERMITS_KEY, value, line);
}

fn emit_semaphore_get_queued(chunks: &mut [Chunk], current: usize, sem: u16, line: u32) {
    get_object_prop(chunks, current, sem, SEMAPHORE_CELLS_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    host::emit(&mut chunks[current], "ecma:atomics", "load", 2, line);
}

fn emit_semaphore_atomic_add_cell(
    chunks: &mut [Chunk],
    current: usize,
    sem: u16,
    cell: i32,
    delta: i32,
    line: u32,
) {
    get_object_prop(chunks, current, sem, SEMAPHORE_CELLS_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, cell);
    core_wasm::i32_const(&mut chunks[current], line, delta.abs());
    let op = if delta >= 0 { "add" } else { "sub" };
    host::emit(&mut chunks[current], "ecma:atomics", op, 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_dyn_sub(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_emitter::ops::emit_dyn_neg(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
}

pub fn emit_semaphore_acquire(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (sem, permits) = emit_semaphore_take_args(chunks, current, argc, false, line);
    let available = chunks[current].alloc_scratch(1);
    let observed = chunks[current].alloc_scratch(1);
    let queued = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, queued, line);

    let sub_dur_idx = chunks[0].add_import("wasi:clocks/monotonic-clock", "subscribe-duration");
    let block_idx = chunks[0].add_import("wasi:io/poll", "[method]pollable.block");

    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    emit_semaphore_get_permits(chunks, current, sem, line);
    set(&mut chunks[current], available, line);
    get(&mut chunks[current], available, line);
    get(&mut chunks[current], permits, line);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    get_object_prop(chunks, current, sem, SEMAPHORE_CELLS_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(&mut chunks[current], available, line);
    get(&mut chunks[current], available, line);
    get(&mut chunks[current], permits, line);
    emit_dyn_sub(chunks, current, line);
    host::emit(
        &mut chunks[current],
        "ecma:atomics",
        "compareExchange",
        4,
        line,
    );
    set(&mut chunks[current], observed, line);
    get(&mut chunks[current], observed, line);
    get(&mut chunks[current], available, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, queued, line);
    chunks[current].emit_if_value(line);
    emit_semaphore_atomic_add_cell(chunks, current, sem, 1, -1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(3, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, queued, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_semaphore_atomic_add_cell(chunks, current, sem, 1, 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, queued, line);
    chunks[current].emit_end(line);
    chunks[current].emit_f64_const(5.0, line);
    vybe_emitter::threading::emit_sleep(&mut chunks[current], sub_dur_idx, block_idx, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_semaphore_release(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (sem, permits) = emit_semaphore_take_args(chunks, current, argc, false, line);
    get_object_prop(chunks, current, sem, SEMAPHORE_CELLS_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(&mut chunks[current], permits, line);
    host::emit(&mut chunks[current], "ecma:atomics", "add", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_semaphore_try_acquire(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (sem, permits) = emit_semaphore_take_args(chunks, current, argc, argc >= 3, line);
    emit_semaphore_get_permits(chunks, current, sem, line);
    get(&mut chunks[current], permits, line);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_semaphore_get_permits(chunks, current, sem, line);
    get(&mut chunks[current], permits, line);
    emit_dyn_sub(chunks, current, line);
    emit_semaphore_set_permits_from_top(chunks, current, sem, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_semaphore_drain(chunks: &mut [Chunk], current: usize, line: u32) {
    let sem = chunks[current].alloc_scratch(1);
    let permits = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sem, line);
    emit_semaphore_get_permits(chunks, current, sem, line);
    set(&mut chunks[current], permits, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    emit_semaphore_set_permits_from_top(chunks, current, sem, line);
    get(&mut chunks[current], permits, line);
}

pub fn emit_semaphore_has_queued(chunks: &mut [Chunk], current: usize, line: u32) {
    let sem = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sem, line);
    emit_semaphore_get_queued(chunks, current, sem, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_semaphore_queue_length(chunks: &mut [Chunk], current: usize, line: u32) {
    let sem = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sem, line);
    emit_semaphore_get_queued(chunks, current, sem, line);
}

pub fn emit_semaphore_is_fair(chunks: &mut [Chunk], current: usize, line: u32) {
    let sem = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sem, line);
    get_object_prop(chunks, current, sem, SEMAPHORE_FAIR_KEY, line);
}

pub fn emit_java_thread_start_with(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let target = chunks[current].alloc_scratch(1);
    let thread = chunks[current].alloc_scratch(1);
    let task = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], target, line);
    set(&mut chunks[current], thread, line);

    get(&mut chunks[current], thread, line);
    chunks[current].emit_string_const(JAVA_THREAD_ALIVE_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], thread, line);
    chunks[current].emit_string_const("__target", line);
    get(&mut chunks[current], target, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    let mut worker = create_function_chunk("__java_thread_worker", 1);
    let current_key = worker.add_constant(Value::String(Arc::from("__j_current_thread")));
    let run_key = worker.add_constant(Value::String(Arc::from("__j_runnable_run")));
    let target_key = worker.add_constant(Value::String(Arc::from("__target")));
    worker.emit_op_u16(Op::LOCAL_GET, 0, line);
    worker.emit_op_u16(Op::GLOBAL_SET, current_key, line);
    worker.emit_op_u16(Op::GLOBAL_GET, run_key, line);
    worker.emit_op_u16(Op::LOCAL_GET, 0, line);
    worker.emit_op_u16(Op::STRUCT_GET, target_key, line);
    worker.emit_op_u8(Op::CALL_REF, 1, line);
    worker.emit_op(Op::RETURN, line);
    worker.local_count = 1;
    chunks.push(worker);
    let worker_idx = chunks.len() - 1;

    get(&mut chunks[current], thread, line);
    chunks[current].emit_op_u16(Op::REF_FUNC, worker_idx as u16, line);
    chunks[current].emit(0, line);
    chunks[current].emit_op(Op::THREAD_SPAWN, line);
    set(&mut chunks[current], task, line);

    get(&mut chunks[current], thread, line);
    chunks[current].emit_string_const(JAVA_THREAD_HANDLE_KEY, line);
    get(&mut chunks[current], task, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    let sub_dur_idx = chunks[0].add_import("wasi:clocks/monotonic-clock", "subscribe-duration");
    let block_idx = chunks[0].add_import("wasi:io/poll", "[method]pollable.block");
    chunks[current].emit_f64_const(1.0, line);
    vybe_emitter::threading::emit_sleep(&mut chunks[current], sub_dur_idx, block_idx, line);

    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_java_thread_join(chunks: &mut [Chunk], current: usize, line: u32) {
    let thread = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], thread, line);

    get_object_prop(chunks, current, thread, JAVA_THREAD_HANDLE_KEY, line);
    vybe_emitter::threading::emit_thread_join(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], thread, line);
    chunks[current].emit_string_const(JAVA_THREAD_ALIVE_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_java_thread_sleep(chunks: &mut [Chunk], current: usize, line: u32) {
    let sub_dur_idx = chunks[0].add_import("wasi:clocks/monotonic-clock", "subscribe-duration");
    let block_idx = chunks[0].add_import("wasi:io/poll", "[method]pollable.block");
    chunks[current].emit_f64_const(25.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    vybe_emitter::threading::emit_sleep(&mut chunks[current], sub_dur_idx, block_idx, line);
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
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], cloned, line);
    emit_mark_identity_map(chunks, current, line);
    set(&mut chunks[current], cloned, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], source, line);
    chunks[current].emit_string_const(ACCESS_ORDER_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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

pub fn emit_map_key_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    emit_identity_map_flag(chunks, current, map, line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], out, line);
    emit_identity_entry_key(chunks, current, map, pair, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], out, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "keys", 1, line);
    chunks[current].emit_end(line);
}

pub fn emit_map_values(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    emit_identity_map_flag(chunks, current, map, line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], out, line);
    emit_identity_entry_value(chunks, current, map, pair, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], out, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], map, line);
    host::emit(&mut chunks[current], "ecma:map", "values", 1, line);
    chunks[current].emit_end(line);
}

pub fn emit_map_entry_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let entries = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let entry = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    collections::emit_array_new(chunks, current, 0, line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    emit_identity_entry_key(chunks, current, map, pair, line);
    emit_identity_entry_value(chunks, current, map, pair, line);
    get(&mut chunks[current], map, line);
    collections::emit_array_new(chunks, current, 3, line);
    set(&mut chunks[current], entry, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], entry, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], out, line);
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
    let iterator = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__index", line);
    set(&mut chunks[current], index, line);
    set_object_prop_from_local(chunks, current, iterator, "__last", index, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    let next = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], next, line);
    set_object_prop_from_local(chunks, current, iterator, "__index", next, line);
    get_iterator_list(chunks, current, iterator, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
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
    get_iterator_list(chunks, current, iterator, line);
    collections::emit_len(chunks, current, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_list_iterator(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let index = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    if argc > 1 {
        set(&mut chunks[current], index, line);
    } else {
        chunks[current].emit_i32_const(0, line);
        set(&mut chunks[current], index, line);
    }
    set(&mut chunks[current], list, line);

    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    let iterator = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get(&mut chunks[current], iterator, line);
    chunks[current].emit_string_const("__java_list_iterator", line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    set_object_prop_from_local(chunks, current, iterator, "__list", list, line);
    set_object_prop_from_local(chunks, current, iterator, "__index", index, line);
    set_object_prop_i32(chunks, current, iterator, "__last", -1, line);
    get(&mut chunks[current], iterator, line);
}

pub fn emit_iterator_previous(chunks: &mut [Chunk], current: usize, line: u32) {
    let iterator = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__index", line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], index, line);
    set_object_prop_from_local(chunks, current, iterator, "__index", index, line);
    set_object_prop_from_local(chunks, current, iterator, "__last", index, line);
    get_iterator_list(chunks, current, iterator, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_iterator_has_previous(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("__index", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_iterator_next_index(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("__index", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

pub fn emit_iterator_previous_index(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_string_const("__index", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], result, line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], other, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    get(&mut chunks[current], value, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
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
        chunks[current].emit_string_const("__java_list_iterator", line);
        host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        let iterator = list;
        let index = chunks[current].alloc_scratch(1);
        get_object_prop(chunks, current, iterator, "__index", line);
        set(&mut chunks[current], index, line);
        let backing = chunks[current].alloc_scratch(1);
        get_iterator_list(chunks, current, iterator, line);
        set(&mut chunks[current], backing, line);
        emit_throw_if_immutable_list(chunks, current, backing, line);
        get_iterator_list(chunks, current, iterator, line);
        get(&mut chunks[current], index, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        get(&mut chunks[current], value, line);
        collections::emit_insert(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        get(&mut chunks[current], index, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        let next = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], next, line);
        set_object_prop_from_local(chunks, current, iterator, "__index", next, line);
        set_object_prop_i32(chunks, current, iterator, "__last", -1, line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        emit_throw_if_immutable_list(chunks, current, list, line);
        emit_is_sublist(chunks, current, list, line);
        chunks[current].emit_if_value(line);
        emit_sublist_parent(chunks, current, list, line);
        emit_sublist_to(chunks, current, list, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        get(&mut chunks[current], value, line);
        collections::emit_insert(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        emit_sublist_increment_to(chunks, current, list, 1, line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        emit_is_set_collection(chunks, current, list, line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], value, line);
        collections::emit_contains(chunks, current, line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], value, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], value, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    } else {
        let index = chunks[current].alloc_scratch(1);
        let list = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], index, line);
        set(&mut chunks[current], list, line);
        emit_throw_if_immutable_list(chunks, current, list, line);
        emit_is_sublist(chunks, current, list, line);
        chunks[current].emit_if_value(line);
        emit_sublist_parent(chunks, current, list, line);
        emit_sublist_absolute_index(chunks, current, list, index, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        get(&mut chunks[current], value, line);
        collections::emit_insert(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        emit_sublist_increment_to(chunks, current, list, 1, line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], index, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        get(&mut chunks[current], value, line);
        collections::emit_insert(chunks, current, line);
        chunks[current].emit_end(line);
    }
}

pub fn emit_sorted_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_sort_if_ordered(chunks, current, list, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
}

pub fn emit_priority_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let comparator = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], list, line);
    emit_comparator(chunks, current, list, line);
    set(&mut chunks[current], comparator, line);
    get(&mut chunks[current], comparator, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], comparator, line);
    heap::emit_push_with_comparator(chunks, current, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    heap::emit_push(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_bool_const(true, line);
}

pub fn emit_stack_push(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], value, line);
}

pub fn emit_list_clone(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, i32::MAX);
    collections::emit_slice(chunks, current, line);
}

pub fn emit_stack_search(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_last_index_of(chunks, current, line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_end(line);
}

pub fn emit_vector_capacity(chunks: &mut [Chunk], current: usize, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    let capacity = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    get_object_prop(chunks, current, list, VECTOR_CAPACITY_KEY, line);
    set(&mut chunks[current], capacity, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    get(&mut chunks[current], capacity, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], capacity, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], len, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], capacity, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_vector_ensure_capacity(chunks: &mut [Chunk], current: usize, line: u32) {
    let min_capacity = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], min_capacity, line);
    set(&mut chunks[current], list, line);
    set_object_prop_from_local(
        chunks,
        current,
        list,
        VECTOR_CAPACITY_KEY,
        min_capacity,
        line,
    );
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_vector_trim_to_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    set_object_prop_from_local(chunks, current, list, VECTOR_CAPACITY_KEY, len, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_vector_set_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let size = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], size, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const("length", line);
    get(&mut chunks[current], size, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_enumeration_from_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let items = chunks[current].alloc_scratch(1);
    let enumeration = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], items, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    set(&mut chunks[current], enumeration, line);
    set_object_prop_from_local(chunks, current, enumeration, ENUM_ITEMS_KEY, items, line);
    set_object_prop_i32(chunks, current, enumeration, ENUM_INDEX_KEY, 0, line);
    get(&mut chunks[current], enumeration, line);
}

pub fn emit_enumeration_has_more(chunks: &mut [Chunk], current: usize, line: u32) {
    let enumeration = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], enumeration, line);
    get_object_prop(chunks, current, enumeration, ENUM_INDEX_KEY, line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], index, line);
    get_object_prop(chunks, current, enumeration, ENUM_ITEMS_KEY, line);
    collections::emit_len(chunks, current, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_enumeration_next(chunks: &mut [Chunk], current: usize, line: u32) {
    let enumeration = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let next_index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], enumeration, line);
    get_object_prop(chunks, current, enumeration, ENUM_INDEX_KEY, line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], next_index, line);
    set_object_prop_from_local(
        chunks,
        current,
        enumeration,
        ENUM_INDEX_KEY,
        next_index,
        line,
    );
    get_object_prop(chunks, current, enumeration, ENUM_ITEMS_KEY, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_hashtable_put(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_throw_if_null(chunks, current, key, line);
    emit_throw_if_null(chunks, current, value, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    emit_map_put(chunks, current, line);
}

pub fn emit_hashtable_keys(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:map", "keys", 1, line);
    emit_enumeration_from_array(chunks, current, line);
}

pub fn emit_hashtable_elements(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:map", "values", 1, line);
    emit_enumeration_from_array(chunks, current, line);
}

pub fn emit_queue_poll(chunks: &mut [Chunk], current: usize, line: u32) {
    let queue = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], queue, line);
    emit_sort_if_ordered(chunks, current, queue, line);
    get(&mut chunks[current], queue, line);
    emit_poll(chunks, current, false, line);
}

pub fn emit_priority_peek(chunks: &mut [Chunk], current: usize, line: u32) {
    let queue = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], queue, line);
    emit_sort_if_ordered(chunks, current, queue, line);
    get(&mut chunks[current], queue, line);
    emit_peek(chunks, current, false, line);
}

pub fn emit_peek(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    if last {
        get(&mut chunks[current], list, line);
        collections::emit_len(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_SUB, line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_poll(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    if last {
        collections::emit_pop(chunks, current, line);
    } else {
        collections::emit_shift(chunks, current, line);
    }
    chunks[current].emit_end(line);
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

pub fn emit_sorted_bound(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
    let bound = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    let candidate = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], bound, line);
    set(&mut chunks[current], list, line);
    emit_sort_if_ordered(chunks, current, list, line);
    get(&mut chunks[current], list, line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    emit_bound_condition(chunks, current, value, bound, mode, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if matches!(mode, 0 | 2) {
        get(&mut chunks[current], found, line);
        vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], value, line);
        set(&mut chunks[current], candidate, line);
        chunks[current].emit_bool_const(true, line);
        set(&mut chunks[current], found, line);
        chunks[current].emit_end(line);
    } else {
        get(&mut chunks[current], value, line);
        set(&mut chunks[current], candidate, line);
        chunks[current].emit_bool_const(true, line);
        set(&mut chunks[current], found, line);
    }
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], candidate, line);
}

pub fn emit_sorted_descending_set(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_clone(chunks, current, line);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    emit_sort_if_ordered(chunks, current, list, line);
    get(&mut chunks[current], list, line);
    collections::emit_reverse(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(DESCENDING_SET_KEY, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], list, line);
    emit_mark_set_collection(chunks, current, line);
}

pub fn emit_sorted_set_range_view(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
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
    let list = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    emit_sort_if_ordered(chunks, current, list, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    emit_range_condition(chunks, current, value, lower, upper, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], out, line);
    emit_mark_set_collection(chunks, current, line);
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
    collections::emit_sort_with_comparator(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], keys, line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(DESCENDING_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], use_last, line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    set(&mut chunks[current], use_last, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], use_last, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    collections::emit_sort_with_comparator(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(DESCENDING_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], keys, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], key, line);

    emit_bound_condition(chunks, current, key, bound, mode, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if matches!(mode, 0 | 2) {
        get(&mut chunks[current], found, line);
        vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
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
            vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
            vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
        }
        1 => {
            vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
            vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
        }
        2 => vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line),
        _ => vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line),
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
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
        vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
        vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
        vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
        vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    emit_throw_if_immutable_list(chunks, current, list, line);

    emit_is_set_collection(chunks, current, list, line);
    chunks[current].emit_if_value(line);
    let index_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);
    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index_slot, line);
    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index_slot, line);
    get(&mut chunks[current], len_slot, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], source, line);
    get(&mut chunks[current], index_slot, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value_slot, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    emit_sort_if_ordered(chunks, current, list, line);
    get(&mut chunks[current], changed, line);
    chunks[current].emit_else(line);

    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    chunks[current].emit_end(line);
}

pub fn emit_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    if argc == 2 {
        let iterator = chunks[current].alloc_scratch(1);
        let index = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], iterator, line);
        get_object_prop(chunks, current, iterator, "__last", line);
        set(&mut chunks[current], index, line);
        get_iterator_list(chunks, current, iterator, line);
        get(&mut chunks[current], index, line);
        get(&mut chunks[current], value, line);
        collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let index = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let previous = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], index, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
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

pub fn emit_remove_at(chunks: &mut [Chunk], current: usize, line: u32) {
    let index = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], index, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    emit_is_sublist(chunks, current, list, line);
    chunks[current].emit_if_value(line);
    emit_sublist_parent(chunks, current, list, line);
    emit_sublist_absolute_index(chunks, current, list, index, line);
    collections::emit_remove_at(chunks, current, line);
    emit_sublist_increment_to(chunks, current, list, -1, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_remove_value_checked(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_remove_value(chunks, current, line);
}

pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    emit_is_sublist(chunks, current, list, line);
    chunks[current].emit_if_value(line);
    let count = chunks[current].alloc_scratch(1);
    emit_sublist_to(chunks, current, list, line);
    emit_sublist_from(chunks, current, list, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    set(&mut chunks[current], count, line);
    emit_sublist_parent(chunks, current, list, line);
    emit_sublist_from(chunks, current, list, line);
    get(&mut chunks[current], count, line);
    collections::emit_remove_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_sublist_from(chunks, current, list, line);
    let to = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], to, line);
    set_object_prop_from_local(chunks, current, list, SUBLIST_TO_KEY, to, line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    host::emit(&mut chunks[current], "ecma:array", "clear", 1, line);
    chunks[current].emit_end(line);
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
    collections::emit_sort_with_comparator(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_remove_all(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_filter_members(chunks, current, line, false);
}

pub fn emit_retain_all(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_filter_members(chunks, current, line, true);
}

pub fn emit_remove_if(chunks: &mut [Chunk], current: usize, line: u32) {
    let predicate = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let length = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], predicate, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    get(&mut chunks[current], list, line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], predicate, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], length, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], length, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_end(line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], changed, line);
}

pub fn emit_replace_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let operator = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let length = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], operator, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], length, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], length, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], operator, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_op(Op::NULL, line);
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
    emit_throw_if_immutable_list(chunks, current, list, line);
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
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], snapshot, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], members, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], changed, line);
}
