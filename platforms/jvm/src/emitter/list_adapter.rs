//! JVM `java.util.List` / `ArrayList` / `HashMap` / iterators.
//!
//! Moved from `languages/java`: these are `java.util` classes, so they belong
//! with the platform that models the JDK, reachable by every JVM language
//! rather than only by the one whose crate happened to hold them.
//!
//! ## Why this sits BESIDE `collection_adapter` rather than merging into it
//!
//! The platform already had a collection adapter, and 29 function names appear
//! in both — with DIFFERENT bodies. That is not accidental duplication: Kotlin
//! is a live consumer of `collection_adapter` (six profile rows plus a direct
//! call to `emit_add`), and Java's semantics are not Kotlin's. Merging them is
//! a real question about which behaviour is correct for whom, answerable only
//! with both suites measured — and the concat helper in `stream_adapter` is
//! the cautionary case: two "equivalent" stringifiers where one reached the
//! object's `toString` slot and the other did not.
//!
//! So the code moves now, and the reconciliation becomes a platform-internal
//! question instead of a cross-crate one.
use std::sync::Arc;
use vybe_compiler::primitives::{
    callable, collections,
    functions::create_function_chunk,
    heap,
    instructions::{core_wasm, host},
    ops, sets, sorted_collection,
};
use vybe_runtime::Chunk;
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

/// A set VIEW's backing map — `keySet()` / `entrySet()` keep it in step.
const BACKING_MAP_KEY: &str = "__java_backing_map";

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

pub fn emit_atomic_new(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    chunks[current].emit_struct_new(0, 0, line);
    let cell = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], cell, line);
    set_object_prop_from_local(chunks, current, cell, "value", value, line);
    get(&mut chunks[current], cell, line);
}

pub fn emit_atomic_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let cell = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], cell, line);
    get_object_prop(chunks, current, cell, "value", line);
}

pub fn emit_atomic_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let cell = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], cell, line);
    set_object_prop_from_local(chunks, current, cell, "value", value, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_atomic_get_and_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let cell = chunks[current].alloc_scratch(1);
    let old = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], cell, line);
    get_object_prop(chunks, current, cell, "value", line);
    set(&mut chunks[current], old, line);
    set_object_prop_from_local(chunks, current, cell, "value", value, line);
    get(&mut chunks[current], old, line);
}

pub fn emit_atomic_compare_and_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let update = chunks[current].alloc_scratch(1);
    let expected = chunks[current].alloc_scratch(1);
    let cell = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], update, line);
    set(&mut chunks[current], expected, line);
    set(&mut chunks[current], cell, line);
    get_object_prop(chunks, current, cell, "value", line);
    get(&mut chunks[current], expected, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    set_object_prop_from_local(chunks, current, cell, "value", update, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_atomic_delta(
    chunks: &mut [Chunk],
    current: usize,
    delta: f64,
    return_old: bool,
    line: u32,
) {
    let cell = chunks[current].alloc_scratch(1);
    let old = chunks[current].alloc_scratch(1);
    let next = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], cell, line);
    get_object_prop(chunks, current, cell, "value", line);
    set(&mut chunks[current], old, line);
    get(&mut chunks[current], old, line);
    chunks[current].emit_f64_const(delta, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], next, line);
    set_object_prop_from_local(chunks, current, cell, "value", next, line);
    get(
        &mut chunks[current],
        if return_old { old } else { next },
        line,
    );
}

pub fn emit_atomic_add_and_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let delta = chunks[current].alloc_scratch(1);
    let cell = chunks[current].alloc_scratch(1);
    let next = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], delta, line);
    set(&mut chunks[current], cell, line);
    get_object_prop(chunks, current, cell, "value", line);
    get(&mut chunks[current], delta, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], next, line);
    set_object_prop_from_local(chunks, current, cell, "value", next, line);
    get(&mut chunks[current], next, line);
}

fn get_iterator_list(chunks: &mut [Chunk], current: usize, iterator: u16, line: u32) {
    get_object_prop(chunks, current, iterator, "__list", line);
}

fn emit_is_sublist(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(SUBLIST_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], next, line);
    set_object_prop_from_local(chunks, current, list, SUBLIST_TO_KEY, next, line);
}

const COMPARATOR_KEY: &str = sorted_collection::COMPARATOR_KEY;
const IMMUTABLE_MAP_KEY: &str = "__java_immutable_map";
const IMMUTABLE_LIST_KEY: &str = "__java_immutable_list";
const SET_COLLECTION_KEY: &str = sorted_collection::SET_COLLECTION_KEY;
const SUBLIST_KEY: &str = "__java_sublist";
const SUBLIST_PARENT_KEY: &str = "__java_sublist_parent";
const SUBLIST_FROM_KEY: &str = "__java_sublist_from";
const SUBLIST_TO_KEY: &str = "__java_sublist_to";
const DESCENDING_MAP_KEY: &str = sorted_collection::DESCENDING_MAP_KEY;
const DESCENDING_SET_KEY: &str = sorted_collection::DESCENDING_SET_KEY;
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
const BLOCKING_QUEUE_CAPACITY_KEY: &str = "__java_blocking_queue_capacity";
const ENUM_ITEMS_KEY: &str = "__java_enum_items";
const ENUM_INDEX_KEY: &str = "__java_enum_index";

fn emit_comparator(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    sorted_collection::emit_comparator(chunks, current, value, line);
}

fn emit_sort_if_ordered(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    sorted_collection::emit_sort_if_ordered(chunks, current, value, line);
}

fn emit_java_exception_throw(chunks: &mut [Chunk], current: usize, name: &str, line: u32) {
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        name,
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

fn emit_throw_if_null(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
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
    sorted_collection::emit_mark_set_collection(chunks, current, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_java_exception_throw(chunks, current, "UnsupportedOperationException", line);
    chunks[current].emit_end(line);
}

fn emit_throw_if_immutable_list(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(IMMUTABLE_LIST_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_java_exception_throw(chunks, current, "UnsupportedOperationException", line);
    chunks[current].emit_end(line);
}

fn emit_is_set_collection(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(SET_COLLECTION_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn emit_canonicalize_map_key(chunks: &mut [Chunk], current: usize, map: u16, key: u16, line: u32) {
    get(&mut chunks[current], map, line);
    chunks[current].emit_string_const(IDENTITY_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    get(&mut chunks[current], key, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    let identity = chunks[current].alloc_scratch(1);
    get_object_prop(chunks, current, key, IDENTITY_KEY_PROP, line);
    set(&mut chunks[current], identity, line);
    get(&mut chunks[current], identity, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("undefined", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get_object_prop(chunks, current, map, IDENTITY_NEXT_KEY, line);
    set(&mut chunks[current], identity, line);
    get(&mut chunks[current], identity, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("undefined", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    set(&mut chunks[current], identity, line);
    chunks[current].emit_end(line);
    set_object_prop_from_local(chunks, current, key, IDENTITY_KEY_PROP, identity, line);
    get(&mut chunks[current], identity, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    sorted_collection::emit_sorted_collection_new(chunks, current, argc, map, line);
}

pub fn emit_sub_list(chunks: &mut [Chunk], current: usize, line: u32) {
    let to = chunks[current].alloc_scratch(1);
    let from = chunks[current].alloc_scratch(1);
    let parent = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], to, line);
    set(&mut chunks[current], from, line);
    set(&mut chunks[current], parent, line);

    chunks[current].emit_struct_new(0, 0, line);
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
    emit_dyn_sub(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_double_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:math", "floor", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], result, line);
    get(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "wasm:js-boolean", "test", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    emit_unwrap_identity_map_value(chunks, current, map, result, line);
    chunks[current].emit_end(line);
}

pub fn emit_get_or_map_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let key = chunks[current].alloc_scratch(1);
    let collection = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], collection, line);

    get(&mut chunks[current], collection, line);
    get(&mut chunks[current], key, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], result, line);

    get(&mut chunks[current], result, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], collection, line);
    get(&mut chunks[current], key, line);
    emit_map_get(chunks, current, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], result, line);
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

/// `new LinkedHashMap()` / `(capacity)` / `(capacity, loadFactor)` /
/// `(capacity, loadFactor, accessOrder)`.
///
/// Insertion order costs nothing — an ecma Map already iterates that way.
/// The THIRD constructor argument is the one observable choice: with
/// `accessOrder = true` every `get`/`put` of a present key moves the entry
/// to the tail (LRU), which `emit_touch_access_order_map` performs wherever
/// the mark is set. Capacity and load factor are sizing hints and drop.
pub fn emit_linked_hash_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let access_order = if argc == 3 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op(Op::DROP, line);
        Some(slot)
    } else {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        None
    };
    collections::emit_map_new(chunks, current, line);
    if let Some(flag) = access_order {
        let map = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], map, line);
        get(&mut chunks[current], flag, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        get(&mut chunks[current], map, line);
        emit_mark_access_order_map(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
        get(&mut chunks[current], map, line);
    }
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_throw_if_null(chunks, current, key, line);
    chunks[current].emit_end(line);
    emit_canonicalize_map_key(chunks, current, map, key, line);
    emit_throw_if_immutable_map(chunks, current, map, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    set(&mut chunks[current], had_key, line);
    get(&mut chunks[current], had_key, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    emit_unwrap_identity_map_value(chunks, current, map, previous, line);
    chunks[current].emit_end(line);
}

pub fn emit_concurrent_map_put(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], map, line);
    emit_throw_if_null(chunks, current, key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    emit_map_put(chunks, current, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], found, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_br(1, line);
    chunks[current].emit_end(line);
    emit_identity_entry_value(chunks, current, map, pair, line);
    get(&mut chunks[current], expected, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], found, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    callable::emit_direct_invoke(chunks, current, 1, line);
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
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], map, line);
        get(&mut chunks[current], key, line);
        host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
        set(&mut chunks[current], current_value, line);
        emit_unwrap_identity_map_value(chunks, current, map, current_value, line);
        get(&mut chunks[current], expected, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], previous, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], previous, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], old_value, line);

    get(&mut chunks[current], fn_slot, line);
    get(&mut chunks[current], original_key, line);
    get(&mut chunks[current], old_value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    emit_unwrap_identity_map_value(chunks, current, map, old_value, line);
    chunks[current].emit_end(line);
    callable::emit_direct_invoke(chunks, current, 2, line);
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
    callable::emit_direct_invoke(chunks, current, 2, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    callable::emit_direct_invoke(chunks, current, 2, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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
    callable::emit_direct_invoke(chunks, current, 2, line);
    host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], entries, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], pair, line);
    get(&mut chunks[current], callback, line);
    emit_identity_entry_key(chunks, current, map, pair, line);
    emit_identity_entry_value(chunks, current, map, pair, line);
    callable::emit_direct_invoke(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], items, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], item, line);
    get(&mut chunks[current], callback, line);
    get(&mut chunks[current], item, line);
    callable::emit_direct_invoke(chunks, current, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], reducer, line);
    get(&mut chunks[current], acc, line);
    get(&mut chunks[current], items, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    callable::emit_direct_invoke(chunks, current, 2, line);
    set(&mut chunks[current], acc, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set(&mut chunks[current], found, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);
    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], searcher, line);
    get(&mut chunks[current], items, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    callable::emit_direct_invoke(chunks, current, 1, line);
    set(&mut chunks[current], found, line);
    get(&mut chunks[current], found, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    chunks[current].emit_struct_new(0, 0, line);
    let sem = chunks[current].alloc_scratch(1);
    let cells = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], sem, line);

    chunks[current].emit_struct_new(0, 0, line);
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
    vybe_compiler::primitives::ops::emit_dyn_neg(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
}

pub fn emit_semaphore_acquire(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (sem, permits) = emit_semaphore_take_args(chunks, current, argc, false, line);
    let available = chunks[current].alloc_scratch(1);
    let observed = chunks[current].alloc_scratch(1);
    let queued = chunks[current].alloc_scratch(1);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, queued, line);

    let wait_for_idx = chunks[0].add_import("wasi:clocks/monotonic-clock", "wait-for");

    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    emit_semaphore_get_permits(chunks, current, sem, line);
    set(&mut chunks[current], available, line);
    get(&mut chunks[current], available, line);
    get(&mut chunks[current], permits, line);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_semaphore_atomic_add_cell(chunks, current, sem, 1, 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, queued, line);
    chunks[current].emit_end(line);
    chunks[current].emit_f64_const(5.0, line);
    vybe_compiler::primitives::threading::emit_sleep(&mut chunks[current], wait_for_idx, line);
    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_semaphore_release(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (sem, permits) = emit_semaphore_take_args(chunks, current, argc, false, line);
    get_object_prop(chunks, current, sem, SEMAPHORE_CELLS_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(&mut chunks[current], permits, line);
    host::emit(&mut chunks[current], "ecma:atomics", "add", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_semaphore_try_acquire(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (sem, permits) = emit_semaphore_take_args(chunks, current, argc, argc >= 3, line);
    emit_semaphore_get_permits(chunks, current, sem, line);
    get(&mut chunks[current], permits, line);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
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
    let target_key = worker.add_constant(Value::String(Arc::from("__target")));
    // Slot 0 arrives as the thread object's TABLE INDEX (the wasi-threads
    // record's user_arg is an i32; objects cross via funcref table 0).
    worker.emit_op_u16(Op::LOCAL_GET, 0, line);
    worker.emit_op_u16(Op::TABLE_GET, 0, line);
    worker.emit_op_u16(Op::LOCAL_SET, 0, line);
    worker.emit_op_u16(Op::LOCAL_GET, 0, line);
    vybe_compiler::primitives::globals::emit_write(&mut worker, "__j_current_thread", line);
    vybe_compiler::primitives::globals::emit_read(&mut worker, "__j_runnable_run", line);
    worker.emit_op_u16(Op::LOCAL_GET, 0, line);
    worker.emit_struct_field_op(Op::STRUCT_GET, 0, target_key, line);
    callable::emit_direct_invoke_chunk(&mut worker, 1, line);
    worker.emit_op(Op::RETURN, line);
    worker.local_count = 1;
    chunks.push(worker);
    let worker_idx = chunks.len() - 1;

    // [thread_obj] → table 0 (index = user_arg), then the wasi spawn.
    get(&mut chunks[current], thread, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::TABLE_GROW, 0, line); // table 0 (u16 index)
    chunks[current].emit_op_u16(Op::REF_FUNC, worker_idx as u16, line);
    chunks[current].emit(0, line);
    vybe_compiler::primitives::threading::emit_thread_spawn(chunks, current, line);
    set(&mut chunks[current], task, line);

    get(&mut chunks[current], thread, line);
    chunks[current].emit_string_const(JAVA_THREAD_HANDLE_KEY, line);
    get(&mut chunks[current], task, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    let wait_for_idx = chunks[0].add_import("wasi:clocks/monotonic-clock", "wait-for");
    chunks[current].emit_f64_const(1.0, line);
    vybe_compiler::primitives::threading::emit_sleep(&mut chunks[current], wait_for_idx, line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_java_thread_join(chunks: &mut [Chunk], current: usize, line: u32) {
    let thread = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], thread, line);

    get_object_prop(chunks, current, thread, JAVA_THREAD_HANDLE_KEY, line);
    vybe_compiler::primitives::threading::emit_thread_join(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], thread, line);
    chunks[current].emit_string_const(JAVA_THREAD_ALIVE_KEY, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_java_thread_sleep(chunks: &mut [Chunk], current: usize, line: u32) {
    let wait_for_idx = chunks[0].add_import("wasi:clocks/monotonic-clock", "wait-for");
    chunks[current].emit_f64_const(25.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    vybe_compiler::primitives::threading::emit_sleep(&mut chunks[current], wait_for_idx, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], cloned, line);
    emit_mark_identity_map(chunks, current, line);
    set(&mut chunks[current], cloned, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], source, line);
    chunks[current].emit_string_const(ACCESS_ORDER_MAP_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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

/// `Map.values()`, identity-aware: an identity map stores `[originalKey,
/// value]` pairs as its ecma-map values, so the raw `ecma:map.values` answer
/// would hand back the PAIRS. Unwrap element 1 for identity maps; everything
/// else is the plain values view.
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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

    chunks[current].emit_struct_new(0, 0, line);
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
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], result, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], other, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    get(&mut chunks[current], value, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
        // A set VIEW over a map keeps the map in step — `keySet().add(x)` has
        // to reach the backing map, not just the view's array.
        get_object_prop(chunks, current, list, BACKING_MAP_KEY, line);
        let backing = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], backing, line);
        get(&mut chunks[current], backing, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        get(&mut chunks[current], backing, line);
        get(&mut chunks[current], value, line);
        chunks[current].emit_bool_const(true, line);
        host::emit(&mut chunks[current], "ecma:map", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
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
        // ⛔ FOUR receiver kinds here, not two. A dict-backed `HashSet` is a
        // REAL ecma Set — pushing onto it silently does nothing — while a
        // marker-flagged set collection is an array that must stay distinct.
        // Both branches existed, in two different emitters, each blind to the
        // other's case; this is the union, not a new rule.
        crate::emitter::collection_adapter::emit_is_ecma_set(chunks, current, list, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], value, line);
        sets::emit_add_mode(
            chunks,
            current,
            crate::emitter::collection_adapter::JAVA_SET_SEMANTICS,
            line,
        );
        chunks[current].emit_else(line);
        emit_is_set_collection(chunks, current, list, line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], value, line);
        collections::emit_contains(chunks, current, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    sorted_collection::emit_sorted_add(chunks, current, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_vector_trim_to_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    set_object_prop_from_local(chunks, current, list, VECTOR_CAPACITY_KEY, len, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_enumeration_from_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let items = chunks[current].alloc_scratch(1);
    let enumeration = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], items, line);
    chunks[current].emit_struct_new(0, 0, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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

fn emit_set_blocking_queue_capacity(
    chunks: &mut [Chunk],
    current: usize,
    queue: u16,
    capacity: u16,
    line: u32,
) {
    set_object_prop_from_local(
        chunks,
        current,
        queue,
        BLOCKING_QUEUE_CAPACITY_KEY,
        capacity,
        line,
    );
}

fn emit_get_blocking_queue_capacity(chunks: &mut [Chunk], current: usize, queue: u16, line: u32) {
    get_object_prop(chunks, current, queue, BLOCKING_QUEUE_CAPACITY_KEY, line);
}

pub fn emit_blocking_queue_remaining_capacity(chunks: &mut [Chunk], current: usize, line: u32) {
    let queue = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], queue, line);
    emit_get_blocking_queue_capacity(chunks, current, queue, line);
    get(&mut chunks[current], queue, line);
    collections::emit_len(chunks, current, line);
    emit_dyn_sub(chunks, current, line);
}

pub fn emit_blocking_queue_offer(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    throw_on_full: bool,
    line: u32,
) {
    if argc > 2 {
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    let value = chunks[current].alloc_scratch(1);
    let queue = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], queue, line);
    get(&mut chunks[current], queue, line);
    collections::emit_len(chunks, current, line);
    emit_get_blocking_queue_capacity(chunks, current, queue, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], queue, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    if throw_on_full {
        emit_java_exception_throw(chunks, current, "IllegalStateException", line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    chunks[current].emit_end(line);
}

pub fn emit_blocking_queue_put(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let queue = chunks[current].alloc_scratch(1);
    let done = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], queue, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], done, line);
    let wait_for_idx = chunks[0].add_import("wasi:clocks/monotonic-clock", "wait-for");
    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], queue, line);
    collections::emit_len(chunks, current, line);
    emit_get_blocking_queue_capacity(chunks, current, queue, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], queue, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], done, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_f64_const(1.0, line);
    vybe_compiler::primitives::threading::emit_sleep(&mut chunks[current], wait_for_idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], done, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_blocking_queue_take(chunks: &mut [Chunk], current: usize, line: u32) {
    let queue = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], queue, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set(&mut chunks[current], result, line);
    let wait_for_idx = chunks[0].add_import("wasi:clocks/monotonic-clock", "wait-for");
    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], queue, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], queue, line);
    emit_poll(chunks, current, false, line);
    set(&mut chunks[current], result, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_f64_const(1.0, line);
    vybe_compiler::primitives::threading::emit_sleep(&mut chunks[current], wait_for_idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], result, line);
}

pub fn emit_blocking_queue_poll(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 1 {
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
    }
    emit_queue_poll(chunks, current, line);
}

pub fn emit_queue_remove_checked(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_queue_poll(chunks, current, line);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_java_exception_throw(chunks, current, "NoSuchElementException", line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], value, line);
}

pub fn emit_queue_element_checked(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_peek(chunks, current, false, line);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_java_exception_throw(chunks, current, "NoSuchElementException", line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], value, line);
}

pub fn emit_blocking_queue_drain_to(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let max = if argc >= 3 {
        let slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    let dest = chunks[current].alloc_scratch(1);
    let queue = chunks[current].alloc_scratch(1);
    let moved = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], dest, line);
    set(&mut chunks[current], queue, line);
    emit_throw_if_immutable_list(chunks, current, dest, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], moved, line);
    let outer = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], queue, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    if let Some(max) = max {
        get(&mut chunks[current], moved, line);
        get(&mut chunks[current], max, line);
        vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_br_if(1, line);
    }
    get(&mut chunks[current], dest, line);
    get(&mut chunks[current], queue, line);
    emit_poll(chunks, current, false, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], moved, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], moved, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], moved, line);
}

pub fn emit_set_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    if argc >= 3 {
        let index = chunks[current].alloc_scratch(1);
        let list = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], index, line);
        set(&mut chunks[current], list, line);
        emit_throw_if_immutable_list(chunks, current, list, line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], index, line);
        get(&mut chunks[current], value, line);
        crate::emitter::collection_adapter::emit_add(chunks, current, argc, line);
        return;
    }
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    crate::emitter::collection_adapter::emit_add(chunks, current, argc, line);
}

pub fn emit_copy_on_write_add_if_absent(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
}

pub fn emit_queue_poll(chunks: &mut [Chunk], current: usize, line: u32) {
    let queue = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], queue, line);
    emit_sort_if_ordered(chunks, current, queue, line);
    get(&mut chunks[current], queue, line);
    emit_poll(chunks, current, false, line);
}

pub fn emit_peek(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    sorted_collection::emit_sorted_end(chunks, current, last, line);
}

pub fn emit_sorted_set_range_view(chunks: &mut [Chunk], current: usize, mode: u8, line: u32) {
    // Java `subSet`/`headSet` use a half-open upper bound (`key < upper`).
    sorted_collection::emit_sorted_set_range_view(chunks, current, mode, false, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], use_last, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    set(&mut chunks[current], use_last, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], use_last, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    sorted_collection::emit_sorted_map_key_set(chunks, current, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set(&mut chunks[current], candidate, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], keys, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], key, line);

    emit_bound_condition(chunks, current, key, bound, mode, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if matches!(mode, 0 | 2) {
        get(&mut chunks[current], found, line);
        vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    sorted_collection::emit_bound_condition(chunks, current, key, bound, mode, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    // Java map range views use a half-open upper bound (`key < upper`).
    sorted_collection::emit_range_condition(chunks, current, key, lower, upper, false, line);
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

    // The source Collection is read by position below — a `HashSet` argument
    // is a real ECMA Set, so flatten it to its values first.
    get(&mut chunks[current], source, line);
    crate::emitter::collection_adapter::emit_values_snapshot(chunks, current, line);
    set(&mut chunks[current], source, line);

    let index_slot = chunks[current].alloc_scratch(1);
    let len_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);

    // Receiver #1: a real ECMA Set (`HashSet`/`LinkedHashSet`) — the marked
    // branch below is ARRAY code and silently no-ops on a Set, which is how
    // `hashSet.addAll(list)` reported changed=true while adding nothing.
    crate::emitter::collection_adapter::emit_is_ecma_set(chunks, current, list, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);
    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index_slot, line);
    let set_outer = chunks[current].emit_block(line);
    let (set_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index_slot, line);
    get(&mut chunks[current], len_slot, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], source, line);
    get(&mut chunks[current], index_slot, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value_slot, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value_slot, line);
    sets::emit_add_mode(
        chunks,
        current,
        crate::emitter::collection_adapter::JAVA_SET_SEMANTICS,
        line,
    );
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(set_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(set_outer);
    get(&mut chunks[current], changed, line);
    chunks[current].emit_else(line);

    emit_is_set_collection(chunks, current, list, line);
    chunks[current].emit_if_value(line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], source, line);
    get(&mut chunks[current], index_slot, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value_slot, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
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
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    let removed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], index, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], removed, line);
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
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], removed, line);
}

pub fn emit_iterator_remove_unsupported(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    emit_java_exception_throw(chunks, current, "UnsupportedOperationException", line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    // ⛔ THREE receiver kinds, not two.
    //
    // This existed twice — once here knowing about sublists and falling back to
    // `ecma:array.clear`, once in `collection_adapter` knowing about dict-backed
    // Sets and falling back to `collections::clear`. Each was a superset of the
    // other's fallback and blind to its special case, so unifying on either one
    // broke the other language: java's lost `hashset_clear`, kotlin's lost
    // `sublist_add_at_end_extends_parent_tail`.
    //
    // The real fix is a protocol slot — the receiver should answer "how do I
    // clear" instead of the emitter testing what it is. Until that exists, the
    // union is the honest form, and it is written as one chain so a future
    // reader sees all three cases in one place rather than two files.
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    emit_throw_if_immutable_list(chunks, current, value, line);

    emit_is_sublist(chunks, current, value, line);
    chunks[current].emit_if_value(line);
    let count = chunks[current].alloc_scratch(1);
    emit_sublist_to(chunks, current, value, line);
    emit_sublist_from(chunks, current, value, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    set(&mut chunks[current], count, line);
    emit_sublist_parent(chunks, current, value, line);
    emit_sublist_from(chunks, current, value, line);
    get(&mut chunks[current], count, line);
    collections::emit_remove_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_sublist_from(chunks, current, value, line);
    let to = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], to, line);
    set_object_prop_from_local(chunks, current, value, SUBLIST_TO_KEY, to, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    crate::emitter::collection_adapter::emit_is_ecma_set(chunks, current, value, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    vybe_compiler::primitives::sets::emit_clear_snapshot(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:array", "clear", 1, line);
    chunks[current].emit_end(line);
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

pub fn emit_contains_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let ok = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    set(&mut chunks[current], target, line);
    // Both operands are read by position below — flatten Set shapes to their
    // values (the receiver is only READ here, so a snapshot is sound).
    get(&mut chunks[current], source, line);
    crate::emitter::collection_adapter::emit_values_snapshot(chunks, current, line);
    set(&mut chunks[current], source, line);
    get(&mut chunks[current], target, line);
    crate::emitter::collection_adapter::emit_values_snapshot(chunks, current, line);
    set(&mut chunks[current], target, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], ok, line);
    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], source, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], target, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], ok, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], ok, line);
}

pub fn emit_list_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_sequence_equal(chunks, current, line);
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

    // Receiver #1: a real ECMA Set — iterate a values snapshot, delete through
    // the set primitive; the index-based removal below no-ops on a Set.
    let set_snapshot = chunks[current].alloc_scratch(1);
    crate::emitter::collection_adapter::emit_is_ecma_set(chunks, current, list, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    sets::emit_values_array(chunks, current, line);
    set(&mut chunks[current], set_snapshot, line);
    get(&mut chunks[current], set_snapshot, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], length, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);
    let set_outer = chunks[current].emit_block(line);
    let (set_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], length, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], set_snapshot, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], predicate, line);
    get(&mut chunks[current], value, line);
    callable::emit_direct_invoke(chunks, current, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    sets::emit_delete_mode(
        chunks,
        current,
        crate::emitter::collection_adapter::JAVA_SET_SEMANTICS,
        line,
    );
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(set_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(set_outer);
    get(&mut chunks[current], changed, line);
    chunks[current].emit_else(line);

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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], predicate, line);
    get(&mut chunks[current], value, line);
    callable::emit_direct_invoke(chunks, current, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], length, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], length, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_end(line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], operator, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    callable::emit_direct_invoke(chunks, current, 1, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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

    // The membership operand is read by position — flatten a Set argument.
    get(&mut chunks[current], members, line);
    crate::emitter::collection_adapter::emit_values_snapshot(chunks, current, line);
    set(&mut chunks[current], members, line);

    // Receiver #1: a real ECMA Set — iterate a values snapshot and delete
    // through the set primitive; the array path below no-ops on a Set.
    crate::emitter::collection_adapter::emit_is_ecma_set(chunks, current, list, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    sets::emit_values_array(chunks, current, line);
    set(&mut chunks[current], snapshot, line);
    get(&mut chunks[current], snapshot, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], length, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);
    let set_outer = chunks[current].emit_block(line);
    let (set_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], length, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], snapshot, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], members, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    if retain {
        chunks[current].emit_op(Op::I32_EQZ, line);
    }
    chunks[current].emit_if(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    sets::emit_delete_mode(
        chunks,
        current,
        crate::emitter::collection_adapter::JAVA_SET_SEMANTICS,
        line,
    );
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(set_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(set_outer);
    get(&mut chunks[current], changed, line);
    chunks[current].emit_else(line);

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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], snapshot, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], members, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);
}
