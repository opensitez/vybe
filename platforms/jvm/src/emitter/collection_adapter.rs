//! JVM `java.util` collection constructor adapters.
//!
//! These are platform surface, not Java-language surface: Kotlin, Java and any
//! later JVM frontend should reach them by resolving `java.util.*` through the
//! JVM namespace tree.

use vybe_compiler::primitives::{
    collections, errors, heap,
    instructions::{core_wasm, host},
    object, ops, sets, sorted_collection,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const VECTOR_CAPACITY_KEY: &str = "__java_vector_capacity";
const BLOCKING_QUEUE_CAPACITY_KEY: &str = "__java_blocking_queue_capacity";
const IMMUTABLE_LIST_KEY: &str = "__java_immutable_list";
const BACKING_MAP_KEY: &str = "__java_backing_map";
const DESCENDING_SET_KEY: &str = sorted_collection::DESCENDING_SET_KEY;

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

fn get_object_prop(chunks: &mut [Chunk], current: usize, object: u16, key: &str, line: u32) {
    get(&mut chunks[current], object, line);
    chunks[current].emit_string_const(key, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

fn mark_bool(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], value, line);
}

fn get_iterator_list(chunks: &mut [Chunk], current: usize, iterator: u16, line: u32) {
    get_object_prop(chunks, current, iterator, "__list", line);
}

fn get_iterator_view(chunks: &mut [Chunk], current: usize, iterator: u16, line: u32) {
    get_object_prop(chunks, current, iterator, "__values", line);
    let values = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], values, line);
    get(&mut chunks[current], values, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    get_iterator_list(chunks, current, iterator, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], values, line);
    chunks[current].emit_end(line);
}

fn emit_is_ecma_set(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:object", "toStringTag", 1, line);
    chunks[current].emit_string_const("[object Set]", line);
    host::emit(&mut chunks[current], "wasm:js-string", "equals", 2, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_string_const(sorted_collection::SET_COLLECTION_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_sort_if_ordered(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    sorted_collection::emit_sort_if_ordered(chunks, current, value, line);
}

fn emit_comparator(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    sorted_collection::emit_comparator(chunks, current, value, line);
}

fn emit_jvm_exception_throw(chunks: &mut [Chunk], current: usize, name: &str, line: u32) {
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        name,
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
}

fn emit_throw_if_immutable_list(chunks: &mut [Chunk], current: usize, list: u16, line: u32) {
    get(&mut chunks[current], list, line);
    chunks[current].emit_string_const(IMMUTABLE_LIST_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_jvm_exception_throw(chunks, current, "UnsupportedOperationException", line);
    chunks[current].emit_end(line);
}

pub fn emit_mark_immutable_list(chunks: &mut [Chunk], current: usize, line: u32) {
    mark_bool(chunks, current, IMMUTABLE_LIST_KEY, line);
}

pub fn emit_mutable_list_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    if argc == 1 {
        let arg = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], arg, line);
        get(&mut chunks[current], arg, line);
        host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], arg, line);
        collections::emit_clone(chunks, current, line);
        chunks[current].emit_else(line);
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_end(line);
        return;
    }
    collections::emit_array_new(chunks, current, argc as u16, line);
}

pub fn emit_hash_set_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let marked = if argc == 0 {
        sets::emit_new(chunks, current, line);
        true
    } else if argc == 1 {
        sets::emit_from_iterable(chunks, current, line);
        true
    } else {
        let base = chunks[current].alloc_scratch(argc as u16);
        collections::emit_pack_n(chunks, current, argc as u16, base, line);
        sets::emit_from_iterable(chunks, current, line);
        true
    };
    if marked {
        sorted_collection::emit_mark_set_collection(chunks, current, line);
    }
}

pub fn emit_list_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    collections::emit_array_new(chunks, current, argc as u16, line);
    mark_bool(chunks, current, IMMUTABLE_LIST_KEY, line);
}

pub fn emit_list_copy_of(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_clone(chunks, current, line);
    mark_bool(chunks, current, IMMUTABLE_LIST_KEY, line);
}

pub fn emit_set_of(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    sets::emit_literal(chunks, current, argc, line);
    sorted_collection::emit_mark_set_collection(chunks, current, line);
    if argc > 0 {
        let set_slot = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], set_slot, line);
        get(&mut chunks[current], set_slot, line);
        sets::emit_size(chunks, current, line);
        core_wasm::i32_const(&mut chunks[current], line, argc as i32);
        chunks[current].emit_op(Op::I32_NE, line);
        chunks[current].emit_if(line);
        emit_jvm_exception_throw(chunks, current, "IllegalArgumentException", line);
        chunks[current].emit_end(line);
        get(&mut chunks[current], set_slot, line);
    }
    mark_bool(chunks, current, IMMUTABLE_LIST_KEY, line);
}

pub fn emit_set_copy_of(chunks: &mut [Chunk], current: usize, line: u32) {
    sets::emit_from_iterable(chunks, current, line);
    mark_bool(chunks, current, IMMUTABLE_LIST_KEY, line);
}

pub fn emit_new_set_from_map(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    collections::emit_array_new(chunks, current, 0, line);
    let set_slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], set_slot, line);
    sorted_collection::emit_mark_set_collection(chunks, current, line);
    set(&mut chunks[current], set_slot, line);
    set_object_prop_from_local(chunks, current, set_slot, BACKING_MAP_KEY, map, line);
    get(&mut chunks[current], set_slot, line);
}

pub fn emit_reverse_order(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:object", "new", 0, line);
    let comparator = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], comparator, line);
    get(&mut chunks[current], comparator, line);
    chunks[current].emit_string_const("__java_reverse_order", line);
    chunks[current].emit_bool_const(true, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], comparator, line);
}

pub fn emit_add_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let argc = argc as u16;
    if argc == 0 {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc);
    for i in (0..argc).rev() {
        set(&mut chunks[current], base + i, line);
    }
    emit_throw_if_immutable_list(chunks, current, base, line);
    for i in 1..argc {
        get(&mut chunks[current], base, line);
        get(&mut chunks[current], base + i, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_bool_const(argc > 1, line);
}

pub fn emit_swap(chunks: &mut [Chunk], current: usize, line: u32) {
    let j = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let tmp = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], j, line);
    set(&mut chunks[current], i, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], tmp, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], j, line);
    collections::emit_get(chunks, current, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], j, line);
    get(&mut chunks[current], tmp, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_replace_all_values(chunks: &mut [Chunk], current: usize, line: u32) {
    let new_value = chunks[current].alloc_scratch(1);
    let old_value = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], new_value, line);
    set(&mut chunks[current], old_value, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);
    get(&mut chunks[current], list, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], old_value, line);
    object::emit_equals(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], new_value, line);
    collections::emit_set(chunks, current, line);
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
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], changed, line);
}

pub fn emit_rotate(chunks: &mut [Chunk], current: usize, line: u32) {
    let distance = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let count = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], distance, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    get(&mut chunks[current], distance, line);
    set(&mut chunks[current], count, line);

    let positive_block = chunks[current].emit_block(line);
    let (positive_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], count, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_gt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    collections::emit_pop(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], list, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(&mut chunks[current], value, line);
    collections::emit_insert(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], count, line);
    core_wasm::i32_const(&mut chunks[current], line, -1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], count, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(positive_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(positive_block);

    let negative_block = chunks[current].emit_block(line);
    let (negative_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], count, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    collections::emit_shift(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], count, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], count, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(negative_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(negative_block);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_index_of_sublist(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let needle = chunks[current].alloc_scratch(1);
    let haystack = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let j = chunks[current].alloc_scratch(1);
    let hay_len = chunks[current].alloc_scratch(1);
    let needle_len = chunks[current].alloc_scratch(1);
    let found = chunks[current].alloc_scratch(1);
    let ok = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], needle, line);
    set(&mut chunks[current], haystack, line);
    chunks[current].emit_i32_const(-1, line);
    set(&mut chunks[current], found, line);
    get(&mut chunks[current], haystack, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], hay_len, line);
    get(&mut chunks[current], needle, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], needle_len, line);
    if last {
        get(&mut chunks[current], hay_len, line);
        get(&mut chunks[current], needle_len, line);
        ops::emit_dyn_neg(&mut chunks[current], line);
        ops::emit_dyn_add(&mut chunks[current], line);
    } else {
        core_wasm::i32_const(&mut chunks[current], line, 0);
    }
    set(&mut chunks[current], i, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    if last {
        get(&mut chunks[current], i, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        ops::emit_dyn_lt(&mut chunks[current], line);
    } else {
        get(&mut chunks[current], i, line);
        get(&mut chunks[current], hay_len, line);
        get(&mut chunks[current], needle_len, line);
        ops::emit_dyn_neg(&mut chunks[current], line);
        ops::emit_dyn_add(&mut chunks[current], line);
        ops::emit_dyn_gt(&mut chunks[current], line);
    }
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], ok, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], j, line);

    let inner = chunks[current].emit_block(line);
    let (inner_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], j, line);
    get(&mut chunks[current], needle_len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], ok, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], haystack, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], j, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    collections::emit_get(chunks, current, line);
    get(&mut chunks[current], needle, line);
    get(&mut chunks[current], j, line);
    collections::emit_get(chunks, current, line);
    object::emit_equals(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], ok, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], j, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inner_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner);

    get(&mut chunks[current], ok, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], i, line);
    set(&mut chunks[current], found, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], i, line);
    core_wasm::i32_const(&mut chunks[current], line, if last { -1 } else { 1 });
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], found, line);
}

pub fn emit_n_copies(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    collections::emit_new_with_length(chunks, current, line);
    get(&mut chunks[current], value, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_i32_const(i32::MAX, line);
    collections::emit_fill(chunks, current, line);
    mark_bool(chunks, current, IMMUTABLE_LIST_KEY, line);
}

pub fn emit_sorted_set_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 1 {
        let arg = chunks[current].alloc_scratch(1);
        let collection = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], arg, line);
        get(&mut chunks[current], arg, line);
        host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], arg, line);
        collections::emit_clone(chunks, current, line);
        set(&mut chunks[current], collection, line);
        get(&mut chunks[current], collection, line);
        chunks[current].emit_string_const(sorted_collection::COMPARATOR_KEY, line);
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
        get(&mut chunks[current], collection, line);
        sorted_collection::emit_mark_set_collection(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        emit_sort_if_ordered(chunks, current, collection, line);
        get(&mut chunks[current], collection, line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], arg, line);
        sorted_collection::emit_sorted_collection_new(chunks, current, 1, false, line);
        chunks[current].emit_end(line);
        return;
    }
    sorted_collection::emit_sorted_collection_new(chunks, current, argc, false, line);
}

pub fn emit_sorted_map_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    sorted_collection::emit_sorted_collection_new(chunks, current, argc, true, line);
}

pub fn emit_priority_queue_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let comparator = chunks[current].alloc_scratch(1);
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    set(&mut chunks[current], comparator, line);
    collections::emit_array_new(chunks, current, 0, line);
    let collection = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], collection, line);
    get(&mut chunks[current], collection, line);
    chunks[current].emit_string_const(sorted_collection::COMPARATOR_KEY, line);
    get(&mut chunks[current], comparator, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], collection, line);
}

pub fn emit_passthrough_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
}

pub fn emit_copy_on_write_list_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        collections::emit_array_new(chunks, current, 0, line);
    } else {
        collections::emit_clone(chunks, current, line);
    }
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
    chunks[current].emit_i32_const(-1, line);
    set(&mut chunks[current], index, line);
    set_object_prop_from_local(chunks, current, iterator, "__last", index, line);
    emit_is_ecma_set(chunks, current, list, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    sets::emit_values_array(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    let values = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], values, line);
    set_object_prop_from_local(chunks, current, iterator, "__values", values, line);
    get(&mut chunks[current], iterator, line);
}

pub fn emit_iterator_next(chunks: &mut [Chunk], current: usize, line: u32) {
    let iterator = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__index", line);
    set(&mut chunks[current], index, line);
    // `next()` past the end THROWS NoSuchElementException (Iterator contract);
    // the bare index read answered undefined.
    get(&mut chunks[current], index, line);
    get_iterator_view(chunks, current, iterator, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_ge(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "NoSuchElementException",
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    set_object_prop_from_local(chunks, current, iterator, "__last", index, line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    let next = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], next, line);
    set_object_prop_from_local(chunks, current, iterator, "__index", next, line);
    get_iterator_view(chunks, current, iterator, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_iterator_has_next(chunks: &mut [Chunk], current: usize, line: u32) {
    let iterator = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__index", line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], index, line);
    get_iterator_view(chunks, current, iterator, line);
    collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

pub fn emit_iterator_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let iterator = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let values = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__last", line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_jvm_exception_throw(chunks, current, "IllegalStateException", line);
    chunks[current].emit_end(line);

    get_iterator_list(chunks, current, iterator, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    get_iterator_view(chunks, current, iterator, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], item, line);
    get(&mut chunks[current], list, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], item, line);
    sets::emit_delete(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], values, line);
    chunks[current].emit_end(line);

    get_object_prop(chunks, current, iterator, "__values", line);
    set(&mut chunks[current], values, line);
    get(&mut chunks[current], values, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], values, line);
    get(&mut chunks[current], index, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    get_object_prop(chunks, current, iterator, "__index", line);
    get(&mut chunks[current], index, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get_object_prop(chunks, current, iterator, "__index", line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    let next_index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], next_index, line);
    set_object_prop_from_local(chunks, current, iterator, "__index", next_index, line);
    chunks[current].emit_end(line);

    chunks[current].emit_i32_const(-1, line);
    set(&mut chunks[current], index, line);
    set_object_prop_from_local(chunks, current, iterator, "__last", index, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_iterator_previous(chunks: &mut [Chunk], current: usize, line: u32) {
    let iterator = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__index", line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    set(&mut chunks[current], index, line);
    set_object_prop_from_local(chunks, current, iterator, "__index", index, line);
    set_object_prop_from_local(chunks, current, iterator, "__last", index, line);
    get_iterator_view(chunks, current, iterator, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_iterator_has_previous(chunks: &mut [Chunk], current: usize, line: u32) {
    let iterator = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__index", line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_gt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

pub fn emit_iterator_index(chunks: &mut [Chunk], current: usize, previous: bool, line: u32) {
    let iterator = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__index", line);
    if previous {
        core_wasm::i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_SUB, line);
    }
}

pub fn emit_iterator_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let iterator = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__last", line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_jvm_exception_throw(chunks, current, "IllegalStateException", line);
    chunks[current].emit_end(line);

    get_iterator_list(chunks, current, iterator, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_iterator_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let iterator = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    let next = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], iterator, line);
    get_object_prop(chunks, current, iterator, "__index", line);
    set(&mut chunks[current], index, line);
    get_iterator_list(chunks, current, iterator, line);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);

    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(&mut chunks[current], value, line);
    collections::emit_insert(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], next, line);
    set_object_prop_from_local(chunks, current, iterator, "__index", next, line);
    chunks[current].emit_i32_const(-1, line);
    set(&mut chunks[current], index, line);
    set_object_prop_from_local(chunks, current, iterator, "__last", index, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    if argc >= 3 {
        let index = chunks[current].alloc_scratch(1);
        let list = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], index, line);
        set(&mut chunks[current], list, line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], index, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        get(&mut chunks[current], value, line);
        collections::emit_insert(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
        return;
    }
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], list, line);
    emit_throw_if_immutable_list(chunks, current, list, line);
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
    emit_is_ecma_set(chunks, current, list, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    sets::emit_add_changed(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
}

pub fn emit_get(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_get(chunks, current, line);
}

pub fn emit_set(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_set(chunks, current, line);
}

pub fn emit_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    emit_is_ecma_set(chunks, current, value, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    sets::emit_size(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `AbstractCollection.toString()` — `[a, b, c]`, the rendering every
/// `List`/`Set`/`Queue`/`Deque` inherits.
///
/// It has to be declared ON THE TYPE, because no runtime probe can answer it.
/// `jvm.java.to_string` (`object_adapter`) deliberately routes an ECMA array to
/// the plain coercion, since Java renders a real array `int[]` as `[I@1b6d35`
/// and NOT as its elements. A `java.util.ArrayList` is backed by the same ECMA
/// array, so it took that leg too and `list.toString()` answered `1,2,3`
/// instead of `[1, 2, 3]` — while `println(list)` was right, which is what made
/// it look like a printing quirk. The type node is the only place that knows a
/// `Collection` is not an array.
///
/// The set leg is the file's standing idiom (`emit_size`, `emit_contains`): a
/// JVM `HashSet` is a real ECMA Set, so it is flattened to its values before
/// the shared array renderer — the same one `Arrays.toString` uses, so both
/// spellings agree by construction.
pub fn emit_collection_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    emit_is_ecma_set(chunks, current, value, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    sets::emit_values_array(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
    crate::emitter::arrays_adapter::emit_to_string(chunks, current, line);
}

pub fn emit_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let needle = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], needle, line);
    set(&mut chunks[current], value, line);
    emit_is_ecma_set(chunks, current, value, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], needle, line);
    sets::emit_has(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], needle, line);
    collections::emit_contains(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_size(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op(Op::I32_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    emit_throw_if_immutable_list(chunks, current, value, line);
    emit_is_ecma_set(chunks, current, value, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    sets::emit_clear(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    collections::emit_clear(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_remove(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        let value = chunks[current].alloc_scratch(1);
        let list = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], value, line);
        set(&mut chunks[current], list, line);
        emit_throw_if_immutable_list(chunks, current, list, line);
        emit_is_ecma_set(chunks, current, list, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], value, line);
        sets::emit_delete(chunks, current, line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], list, line);
        get(&mut chunks[current], value, line);
        collections::emit_remove_value(chunks, current, line);
        chunks[current].emit_end(line);
    } else {
        collections::emit_remove_at(chunks, current, line);
    }
}

pub fn emit_index_of(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_index_of(chunks, current, line);
}

pub fn emit_collection_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    let source = chunks[current].alloc_scratch(1);
    let dest = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], source, line);
    set(&mut chunks[current], dest, line);
    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    // `Collections.copy` THROWS IndexOutOfBoundsException when the
    // destination is shorter than the source (javadoc-specified).
    get(&mut chunks[current], dest, line);
    collections::emit_len(chunks, current, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Source does not fit in dest", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "IndexOutOfBoundsException",
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], index, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
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
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], target, line);
    object::emit_equals(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], count, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], count, line);
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
    get(&mut chunks[current], count, line);
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
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], result, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], left, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], right, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], result, line);
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
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
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
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
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
        ops::emit_dyn_lt(&mut chunks[current], line);
    } else {
        ops::emit_dyn_gt(&mut chunks[current], line);
    }
    chunks[current].emit_else(line);
    get(&mut chunks[current], comparator, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], best, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    if min {
        ops::emit_dyn_lt(&mut chunks[current], line);
    } else {
        ops::emit_dyn_gt(&mut chunks[current], line);
    }
    chunks[current].emit_end(line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    set(&mut chunks[current], best, line);
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
    get(&mut chunks[current], best, line);
}

pub fn emit_add_first(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:array", "unshift", 2, line);
}

pub fn emit_remove_first(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_shift(chunks, current, line);
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

pub fn emit_priority_poll(chunks: &mut [Chunk], current: usize, line: u32) {
    let queue = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], queue, line);
    emit_sort_if_ordered(chunks, current, queue, line);
    get(&mut chunks[current], queue, line);
    emit_poll(chunks, current, false, line);
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
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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

pub fn emit_priority_peek(chunks: &mut [Chunk], current: usize, line: u32) {
    let queue = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], queue, line);
    emit_sort_if_ordered(chunks, current, queue, line);
    get(&mut chunks[current], queue, line);
    emit_peek(chunks, current, false, line);
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
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    set(&mut chunks[current], candidate, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    sorted_collection::emit_bound_condition(chunks, current, value, bound, mode, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if matches!(mode, 0 | 2) {
        get(&mut chunks[current], found, line);
        ops::emit_dyn_not(&mut chunks[current], line);
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
    ops::emit_dyn_add(&mut chunks[current], line);
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
    sorted_collection::emit_mark_set_collection(chunks, current, line);
}

pub fn emit_linked_blocking_queue_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let queue = chunks[current].alloc_scratch(1);
    let capacity = chunks[current].alloc_scratch(1);
    if argc == 0 {
        collections::emit_array_new(chunks, current, 0, line);
        set(&mut chunks[current], queue, line);
        chunks[current].emit_i32_const(i32::MAX, line);
        set(&mut chunks[current], capacity, line);
    } else {
        let arg = chunks[current].alloc_scratch(1);
        set(&mut chunks[current], arg, line);
        get(&mut chunks[current], arg, line);
        host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], arg, line);
        collections::emit_clone(chunks, current, line);
        set(&mut chunks[current], queue, line);
        chunks[current].emit_i32_const(i32::MAX, line);
        set(&mut chunks[current], capacity, line);
        chunks[current].emit_else(line);
        collections::emit_array_new(chunks, current, 0, line);
        set(&mut chunks[current], queue, line);
        get(&mut chunks[current], arg, line);
        set(&mut chunks[current], capacity, line);
        chunks[current].emit_end(line);
    }
    set_object_prop_from_local(
        chunks,
        current,
        queue,
        BLOCKING_QUEUE_CAPACITY_KEY,
        capacity,
        line,
    );
    get(&mut chunks[current], queue, line);
}
