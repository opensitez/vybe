//! Kotlin collection adapters that need Kotlin-specific return contracts.

use std::sync::Arc;
use vybe_compiler::primitives::{collections as common_collections, dict, instructions::host, ops};
use vybe_runtime::Chunk;
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Kotlin `MutableCollection.add(value)`.
///
/// Arrays/lists append and return `true`; Kotlin's dict-backed sets use
/// `MutableSet.add` duplicate semantics and return whether the set changed.
pub fn emit_add(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let collection = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], value, line);
    set(&mut chunks[current], collection, line);

    get(&mut chunks[current], collection, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    get(&mut chunks[current], collection, line);
    get(&mut chunks[current], value, line);
    common_collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);

    chunks[current].emit_else(line);

    get(&mut chunks[current], collection, line);
    get(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    emit_set_add(chunks, current, 3, line);

    chunks[current].emit_end(line);
}

/// Kotlin `MutableSet.add(value)`.
///
/// Stack in: `[set, key, value]`; stack out: `[changed_bool]`.
pub fn emit_set_add(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    let existed = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], value, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], set_slot, line);

    get(&mut chunks[current], set_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], value, line);
    common_collections::emit_contains(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    set(&mut chunks[current], existed, line);

    get(&mut chunks[current], existed, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], value, line);
    common_collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], existed, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], key, line);
    dict::emit_method_has(chunks, current, line);
    set(&mut chunks[current], existed, line);

    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    common_collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], existed, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], set_slot, line);
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, keys_key, line);
    get(&mut chunks[current], key, line);
    common_collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], existed, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// Kotlin dict-backed set `size`. The marker property is implementation detail.
pub fn emit_set_size(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    dict::emit_method_size(chunks, current, line);
}

fn emit_collection_values_array(
    chunks: &mut Vec<Chunk>,
    current: usize,
    collection: u16,
    line: u32,
) {
    get(&mut chunks[current], collection, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], collection, line);
    common_collections::emit_clone(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], collection, line);
    dict::emit_values(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_mark_kotlin_set(chunks: &mut Vec<Chunk>, current: usize, out: u16, line: u32) {
    get(&mut chunks[current], out, line);
    let marker = chunks[current].add_constant(Value::String(Arc::from(
        crate::emitter::tostring::SET_MARKER,
    )));
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, marker, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Kotlin `toList()` / `toTypedArray()` for arrays and dict-backed sets.
pub fn emit_to_list(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let collection = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], collection, line);
    emit_collection_values_array(chunks, current, collection, line);
}

/// Kotlin `toSet()` / `toMutableSet()` for arrays and dict-backed sets.
pub fn emit_to_set(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let collection = chunks[current].alloc_scratch(1);
    let values = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], collection, line);
    emit_collection_values_array(chunks, current, collection, line);
    set(&mut chunks[current], values, line);

    dict::emit_new(chunks, current, line);
    set(&mut chunks[current], out, line);
    emit_mark_kotlin_set(chunks, current, out, line);

    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], values, line);
    common_collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], values, line);
    get(&mut chunks[current], index, line);
    common_collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);

    get(&mut chunks[current], value, line);
    crate::emitter::tostring::emit_to_string(chunks, current, line);
    set(&mut chunks[current], key, line);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    emit_set_add(chunks, current, 3, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    get(&mut chunks[current], out, line);
}

fn emit_set_from_filter(chunks: &mut Vec<Chunk>, current: usize, keep_present: bool, line: u32) {
    let right = chunks[current].alloc_scratch(1);
    let left = chunks[current].alloc_scratch(1);
    let right_set = chunks[current].alloc_scratch(1);
    let values = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], right, line);
    set(&mut chunks[current], left, line);
    get(&mut chunks[current], right, line);
    emit_to_set(chunks, current, 1, line);
    set(&mut chunks[current], right_set, line);
    emit_collection_values_array(chunks, current, left, line);
    set(&mut chunks[current], values, line);
    dict::emit_new(chunks, current, line);
    set(&mut chunks[current], out, line);
    emit_mark_kotlin_set(chunks, current, out, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], values, line);
    common_collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], values, line);
    get(&mut chunks[current], index, line);
    common_collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    crate::emitter::tostring::emit_to_string(chunks, current, line);
    set(&mut chunks[current], key, line);
    get(&mut chunks[current], right_set, line);
    get(&mut chunks[current], key, line);
    dict::emit_method_has(chunks, current, line);
    if !keep_present {
        ops::emit_dyn_not(&mut chunks[current], line);
    }
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], value, line);
    emit_set_add(chunks, current, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], out, line);
}

pub fn emit_set_union(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let right = chunks[current].alloc_scratch(1);
    let left = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], right, line);
    set(&mut chunks[current], left, line);
    get(&mut chunks[current], left, line);
    emit_to_set(chunks, current, 1, line);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], right, line);
    emit_add_all(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_set_intersect(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_set_from_filter(chunks, current, true, line);
}

pub fn emit_set_subtract(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_set_from_filter(chunks, current, false, line);
}

/// Kotlin `MutableCollection.addAll(values)`.
pub fn emit_add_all(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let collection = chunks[current].alloc_scratch(1);
    let values = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], other, line);
    set(&mut chunks[current], collection, line);
    emit_collection_values_array(chunks, current, other, line);
    set(&mut chunks[current], values, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], values, line);
    common_collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], values, line);
    get(&mut chunks[current], index, line);
    common_collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);

    get(&mut chunks[current], collection, line);
    get(&mut chunks[current], value, line);
    emit_add(chunks, current, 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    get(&mut chunks[current], changed, line);
}

fn emit_mutate_by_filter(chunks: &mut Vec<Chunk>, current: usize, delete_present: bool, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let collection = chunks[current].alloc_scratch(1);
    let other_set = chunks[current].alloc_scratch(1);
    let values = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let changed = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], other, line);
    set(&mut chunks[current], collection, line);
    get(&mut chunks[current], other, line);
    emit_to_set(chunks, current, 1, line);
    set(&mut chunks[current], other_set, line);
    emit_collection_values_array(chunks, current, collection, line);
    set(&mut chunks[current], values, line);
    chunks[current].emit_bool_const(false, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], values, line);
    common_collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], values, line);
    get(&mut chunks[current], index, line);
    common_collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    crate::emitter::tostring::emit_to_string(chunks, current, line);
    set(&mut chunks[current], key, line);
    get(&mut chunks[current], other_set, line);
    get(&mut chunks[current], key, line);
    dict::emit_method_has(chunks, current, line);
    if !delete_present {
        ops::emit_dyn_not(&mut chunks[current], line);
    }
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], collection, line);
    get(&mut chunks[current], key, line);
    dict::emit_method_delete(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    set(&mut chunks[current], changed, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
    get(&mut chunks[current], changed, line);
}

pub fn emit_remove_all(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_mutate_by_filter(chunks, current, true, line);
}

pub fn emit_retain_all(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_mutate_by_filter(chunks, current, false, line);
}

/// Kotlin `Collection.containsAll(values)`. Stack in: `[collection, values]`.
pub fn emit_contains_all(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let collection = chunks[current].alloc_scratch(1);
    let values = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], other, line);
    set(&mut chunks[current], collection, line);
    emit_collection_values_array(chunks, current, other, line);
    set(&mut chunks[current], values, line);

    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], index, line);
    get(&mut chunks[current], values, line);
    common_collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], index, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], values, line);
    get(&mut chunks[current], index, line);
    common_collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);

    get(&mut chunks[current], collection, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], collection, line);
    get(&mut chunks[current], value, line);
    common_collections::emit_contains(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    crate::emitter::tostring::emit_to_string(chunks, current, line);
    set(&mut chunks[current], key, line);
    get(&mut chunks[current], collection, line);
    get(&mut chunks[current], key, line);
    dict::emit_method_has(chunks, current, line);
    chunks[current].emit_end(line);

    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    get(&mut chunks[current], index, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);
}

/// Kotlin collection `isEmpty()`. Stack in: `[collection]`; stack out: `[bool]`.
pub fn emit_is_empty(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);

    // Strings answer by length — the dict probe below STRUCT_GETs, which
    // traps on a primitive.
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], value, line);
    common_collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_string_const("__java_immutable_map", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:map", "size", 1, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], value, line);
    dict::emit_method_size(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_is_not_empty(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_is_empty(chunks, current, argc, line);
    ops::emit_dyn_not(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// One `Map.Entry`: the `[k, v]` array, with `key`/`value` properties stamped
/// on it so both spellings work — `for ((k, v) in …)` destructures the array
/// positionally, `it.key`/`it.value` read the properties. Kotlin passes
/// entries to every map lambda (`filter`, `map`, `count`, `forEach`) and
/// yields them from `entries`, so the shape is built in ONE place.
/// Stack: `[k, v]` → `[entry]`.
pub fn emit_make_entry(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    set(&mut chunks[current], k, line);
    get(&mut chunks[current], k, line);
    get(&mut chunks[current], v, line);
    common_collections::emit_array_new(chunks, current, 2, line);
    for (prop, slot) in [("key", k), ("value", v)] {
        chunks[current].emit_dup(line);
        get(&mut chunks[current], slot, line);
        let key_idx = chunks[current].add_constant(Value::String(Arc::from(prop)));
        chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, key_idx, line);
        chunks[current].emit_op(Op::DROP, line);
    }
}

/// A LIST view of a dict receiver: a Set's elements, a Map's entries.
/// Arrays pass through untouched. This is what lets one predicate loop serve
/// `list.count { }`, `set.count { }` and `map.count { it.value > 0 }` — in
/// Kotlin the Map overloads iterate ENTRIES, and the entry shape is
/// [`emit_make_entry`]'s.
/// Stack: `[receiver]` → `[list]`.
pub fn emit_dict_as_list(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    get(&mut chunks[current], v, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    // A string iterates as its characters (`"ab".first()`, `s.count()`).
    get(&mut chunks[current], v, line);
    chunks[current].emit_string_const("", line);
    host::emit(&mut chunks[current], "ecma:string", "split", 2, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], v, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], v, line);
    chunks[current].emit_else(line);

    // A Set is its keys; a Map is its entries.
    get(&mut chunks[current], v, line);
    let marker = chunks[current].add_constant(Value::String(Arc::from(
        crate::emitter::tostring::SET_MARKER,
    )));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, marker, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    // The VALUES view, not `__keys`: a set literal's `__keys` also carries the
    // marker itself, and keys are the STRING spellings of the elements.
    emit_collection_values_array(chunks, current, v, line);
    chunks[current].emit_else(line);
    emit_entries_list(chunks, current, v, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Kotlin `map.entries` — insertion-ordered entry objects.
/// Stack: `[map]` → `[list of entries]`.
pub fn emit_entries(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    emit_entries_list(chunks, current, v, line);
}

fn emit_entries_list(chunks: &mut Vec<Chunk>, current: usize, v: u16, line: u32) {
    let keys = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], v, line);
    dict::emit_keys(chunks, current, line);
    set(&mut chunks[current], keys, line);
    common_collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
    let state = vybe_compiler::primitives::loops::emit_for_in_start(chunks, current, keys, idx, line);
    set(&mut chunks[current], key, line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], v, line);
    get(&mut chunks[current], key, line);
    dict::emit_get_dynamic(chunks, current, line);
    emit_make_entry(chunks, current, line);
    common_collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx, state, line);
    get(&mut chunks[current], out, line);
}
