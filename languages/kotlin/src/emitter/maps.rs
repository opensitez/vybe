//! Kotlin's map extensions — `mapKeys`, `mapValues`, `filterKeys`,
//! `filterValues`, `getOrPut`, `putAll`, `plus`/`minus`, and the entry-lambda
//! forms.
//!
//! A Kotlin map is `common:dict.new`'s shape (a struct with a `__keys` array),
//! and every lambda over a map receives a `Map.Entry` — built by
//! `collections::emit_make_entry` so `it.key`/`it.value` and `(k, v)`
//! destructuring are the same object. Ops that also exist on lists
//! (`plus`, `minus`, `count`) dispatch on the RECEIVER at runtime via
//! `ecma:array.isArray`, because one Kotlin spelling covers both.

use std::sync::Arc;
use vybe_compiler::primitives::{collections as common_collections, dict, loops, ops, strings};
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_runtime::Chunk;
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

/// `[dict, key, value]` → `[]`, appending the key to `__keys` when NEW.
///
/// `dict::emit_set` is a bare `ARRAY_SET`, which stores the property but does
/// not maintain the dict's `__keys` insertion-order array — so every entry
/// written that way existed but never enumerated: the map printed `{}` while
/// `m[k]` answered. Every adapter write goes through here instead.
pub fn emit_dict_set_tracked(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    let d = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    set(chunks, current, k, line);
    set(chunks, current, d, line);
    crate::emitter::collections::emit_throw_if_java_immutable(chunks, current, d, line);
    get(chunks, current, d, line);
    get(chunks, current, k, line);
    dict::emit_method_has(chunks, current, line);
    truthy(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    get(chunks, current, d, line);
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, keys_key, line);
    get(chunks, current, k, line);
    common_collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(chunks, current, d, line);
    get(chunks, current, k, line);
    get(chunks, current, v, line);
    dict::emit_set(chunks, current, line);
}

fn get(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn truthy(chunks: &mut [Chunk], current: usize, line: u32) {
    ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn call_fn(chunks: &mut [Chunk], current: usize, fn_slot: u16, args: &[u16], line: u32) {
    get(chunks, current, fn_slot, line);
    for &a in args {
        get(chunks, current, a, line);
    }
    chunks[current].emit_op_u8(Op::CALL_REF, args.len() as u8, line);
}

/// Iterate the map in `m`, leaving each key in `key` and value in `value`
/// for `body`.
fn for_each_entry(
    chunks: &mut Vec<Chunk>,
    current: usize,
    m: u16,
    key: u16,
    value: u16,
    line: u32,
    body: impl FnOnce(&mut Vec<Chunk>),
) {
    let keys = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    get(chunks, current, m, line);
    dict::emit_keys(chunks, current, line);
    set(chunks, current, keys, line);
    let state = loops::emit_for_in_start(chunks, current, keys, idx, line);
    set(chunks, current, key, line);
    get(chunks, current, m, line);
    get(chunks, current, key, line);
    dict::emit_get_dynamic(chunks, current, line);
    set(chunks, current, value, line);
    body(chunks);
    loops::emit_for_in_end(chunks, current, idx, state, line);
}

/// Build the `Map.Entry` for the current `(key, value)` into `entry`.
fn make_entry(chunks: &mut Vec<Chunk>, current: usize, key: u16, value: u16, entry: u16, line: u32) {
    get(chunks, current, key, line);
    get(chunks, current, value, line);
    crate::emitter::collections::emit_make_entry(chunks, current, line);
    set(chunks, current, entry, line);
}

enum EntryOp {
    /// `filter { entry -> }` — keep entries the predicate accepts.
    Filter { invert: bool },
    /// `filterKeys { k -> }` / `filterValues { v -> }`.
    FilterProjected { on_key: bool },
    /// `mapValues { entry -> }` — same keys, transformed values.
    MapValues,
    /// `mapKeys { entry -> }` — transformed keys, same values.
    MapKeys,
}

fn emit_entry_op(chunks: &mut Vec<Chunk>, current: usize, op: EntryOp, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, m, line);
    let out = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let entry = chunks[current].alloc_scratch(1);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    for_each_entry(chunks, current, m, key, value, line, |chunks| match op {
        EntryOp::Filter { invert } => {
            make_entry(chunks, current, key, value, entry, line);
            call_fn(chunks, current, f, &[entry], line);
            truthy(chunks, current, line);
            if invert {
                chunks[current].emit_op(Op::I32_EQZ, line);
            }
            chunks[current].emit_if(line);
            get(chunks, current, out, line);
            get(chunks, current, key, line);
            get(chunks, current, value, line);
            emit_dict_set_tracked(chunks, current, line);
            chunks[current].emit_end(line);
        }
        EntryOp::FilterProjected { on_key } => {
            call_fn(chunks, current, f, &[if on_key { key } else { value }], line);
            truthy(chunks, current, line);
            chunks[current].emit_if(line);
            get(chunks, current, out, line);
            get(chunks, current, key, line);
            get(chunks, current, value, line);
            emit_dict_set_tracked(chunks, current, line);
            chunks[current].emit_end(line);
        }
        EntryOp::MapValues => {
            make_entry(chunks, current, key, value, entry, line);
            get(chunks, current, out, line);
            get(chunks, current, key, line);
            call_fn(chunks, current, f, &[entry], line);
            emit_dict_set_tracked(chunks, current, line);
        }
        EntryOp::MapKeys => {
            make_entry(chunks, current, key, value, entry, line);
            get(chunks, current, out, line);
            call_fn(chunks, current, f, &[entry], line);
            get(chunks, current, value, line);
            emit_dict_set_tracked(chunks, current, line);
        }
    });
    get(chunks, current, out, line);
}

pub fn emit_map_filter(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_entry_op(chunks, current, EntryOp::Filter { invert: false }, line);
}

pub fn emit_map_filter_not(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_entry_op(chunks, current, EntryOp::Filter { invert: true }, line);
}

pub fn emit_filter_keys(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_entry_op(chunks, current, EntryOp::FilterProjected { on_key: true }, line);
}

pub fn emit_filter_values(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_entry_op(chunks, current, EntryOp::FilterProjected { on_key: false }, line);
}

pub fn emit_map_values_transform(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_entry_op(chunks, current, EntryOp::MapValues, line);
}

pub fn emit_map_keys_transform(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_entry_op(chunks, current, EntryOp::MapKeys, line);
}

/// `map.map { entry -> }` — a LIST of lambda results, in entry order.
pub fn emit_map_to_list(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, m, line);
    let out = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let entry = chunks[current].alloc_scratch(1);
    common_collections::emit_array_new(chunks, current, 0, line);
    set(chunks, current, out, line);
    for_each_entry(chunks, current, m, key, value, line, |chunks| {
        make_entry(chunks, current, key, value, entry, line);
        get(chunks, current, out, line);
        call_fn(chunks, current, f, &[entry], line);
        common_collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    });
    get(chunks, current, out, line);
}

/// `getValue(k)` — throws `NoSuchElementException` on a missing key, unless
/// the map carries a `withDefault { }` provider, which is consulted first.
pub fn emit_get_value(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let k = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let v = chunks[current].alloc_scratch(1);
    let def = chunks[current].alloc_scratch(1);
    set(chunks, current, k, line);
    set(chunks, current, m, line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_method_has(chunks, current, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_get_dynamic(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, m, line);
    let marker = chunks[current].add_constant(Value::String(Arc::from(DEFAULT_MARKER)));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, marker, line);
    set(chunks, current, def, line);
    get(chunks, current, def, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    call_fn(chunks, current, def, &[k], line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("Key is missing in the map.", line);
    crate::emitter::nullability::emit_exception(chunks, current, 1, "NoSuchElementException", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    get(chunks, current, v, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// The property `withDefault { }` stores its provider under.
const DEFAULT_MARKER: &str = "__kt_default";

/// `withDefault { }` — the same map, carrying the default provider.
pub fn emit_with_default(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, m, line);
    get(chunks, current, m, line);
    get(chunks, current, f, line);
    let marker = chunks[current].add_constant(Value::String(Arc::from(DEFAULT_MARKER)));
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, marker, line);
    get(chunks, current, m, line);
}

/// `getOrPut(k) { }` — get, or compute+insert+return.
pub fn emit_get_or_put(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let v = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, k, line);
    set(chunks, current, m, line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_method_has(chunks, current, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_get_dynamic(chunks, current, line);
    chunks[current].emit_else(line);
    call_fn(chunks, current, f, &[], line);
    set(chunks, current, v, line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    get(chunks, current, v, line);
    emit_dict_set_tracked(chunks, current, line);
    get(chunks, current, v, line);
    chunks[current].emit_end(line);
}

/// `putAll(other)` — Unit result.
pub fn emit_put_all(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    set(chunks, current, other, line);
    set(chunks, current, m, line);
    get(chunks, current, other, line);
    let is_arr = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_arr, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    // A LIST of pairs — `putAll(listOf(k to v, …))`.
    emit_put_pairs(chunks, current, other, m, line);
    chunks[current].emit_else(line);
    emit_copy_entries(chunks, current, other, m, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Insert every `[k, v]` pair of the list in `src` into the map in `dst`.
fn emit_put_pairs(chunks: &mut Vec<Chunk>, current: usize, src: u16, dst: u16, line: u32) {
    let idx = chunks[current].alloc_scratch(1);
    let pair = chunks[current].alloc_scratch(1);
    let state = loops::emit_for_in_start(chunks, current, src, idx, line);
    set(chunks, current, pair, line);
    get(chunks, current, dst, line);
    get(chunks, current, pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    common_collections::emit_get(chunks, current, line);
    get(chunks, current, pair, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    common_collections::emit_get(chunks, current, line);
    emit_dict_set_tracked(chunks, current, line);
    loops::emit_for_in_end(chunks, current, idx, state, line);
}

/// `putIfAbsent(k, v)` — existing value, or `null` after inserting.
pub fn emit_put_if_absent(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    set(chunks, current, k, line);
    set(chunks, current, m, line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_method_has(chunks, current, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_get_dynamic(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    get(chunks, current, v, line);
    emit_dict_set_tracked(chunks, current, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
}

/// Copy every entry of the map in `src` into the map in `dst`.
fn emit_copy_entries(chunks: &mut Vec<Chunk>, current: usize, src: u16, dst: u16, line: u32) {
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    for_each_entry(chunks, current, src, key, value, line, |chunks| {
        get(chunks, current, dst, line);
        get(chunks, current, key, line);
        get(chunks, current, value, line);
        emit_dict_set_tracked(chunks, current, line);
    });
}

/// `toMutableMap()` / `toMap()` on a map — an independent copy.
pub fn emit_copy_map(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let m = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set(chunks, current, m, line);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    emit_copy_entries(chunks, current, m, out, line);
    get(chunks, current, out, line);
}

/// `toSortedMap()` — a copy with keys in sorted order.
pub fn emit_to_sorted_map(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // 1-arg form sorts the keys with a COMPARATOR.
    let cmp = if argc >= 2 {
        let cmp = chunks[current].alloc_scratch(1);
        set(chunks, current, cmp, line);
        Some(cmp)
    } else {
        None
    };
    let m = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    set(chunks, current, m, line);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    get(chunks, current, m, line);
    dict::emit_keys(chunks, current, line);
    if let Some(cmp) = cmp {
        get(chunks, current, cmp, line);
        common_collections::emit_sort_with_comparator(chunks, current, line);
    } else {
        common_collections::emit_sorted(chunks, current, line);
    }
    set(chunks, current, keys, line);
    let state = loops::emit_for_in_start(chunks, current, keys, idx, line);
    set(chunks, current, key, line);
    get(chunks, current, out, line);
    get(chunks, current, key, line);
    get(chunks, current, m, line);
    get(chunks, current, key, line);
    dict::emit_get_dynamic(chunks, current, line);
    emit_dict_set_tracked(chunks, current, line);
    loops::emit_for_in_end(chunks, current, idx, state, line);
    get(chunks, current, out, line);
}

/// `plus` — maps merge (`m + other`), lists append an element or a
/// collection. One spelling in Kotlin, so the receiver decides at runtime.
pub fn emit_plus(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let arg = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set(chunks, current, arg, line);
    set(chunks, current, recv, line);

    get(chunks, current, recv, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    // List + (element | collection).
    get(chunks, current, recv, line);
    common_collections::emit_clone(chunks, current, line);
    set(chunks, current, out, line);
    get(chunks, current, arg, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    let idx = chunks[current].alloc_scratch(1);
    let elem = chunks[current].alloc_scratch(1);
    let state = loops::emit_for_in_start(chunks, current, arg, idx, line);
    set(chunks, current, elem, line);
    get(chunks, current, out, line);
    get(chunks, current, elem, line);
    common_collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx, state, line);
    chunks[current].emit_else(line);
    get(chunks, current, out, line);
    get(chunks, current, arg, line);
    common_collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(chunks, current, out, line);
    chunks[current].emit_else(line);
    // Map + (map | pair).
    get(chunks, current, recv, line);
    emit_copy_map_into(chunks, current, out, line);
    get(chunks, current, arg, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    // A Pair: one entry.
    get(chunks, current, out, line);
    get(chunks, current, arg, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    common_collections::emit_get(chunks, current, line);
    get(chunks, current, arg, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    common_collections::emit_get(chunks, current, line);
    emit_dict_set_tracked(chunks, current, line);
    chunks[current].emit_else(line);
    emit_copy_entries(chunks, current, arg, out, line);
    chunks[current].emit_end(line);
    get(chunks, current, out, line);
    chunks[current].emit_end(line);
}

/// `[map] →` leaves a fresh copy in `out`.
fn emit_copy_map_into(chunks: &mut Vec<Chunk>, current: usize, out: u16, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    set(chunks, current, src, line);
    dict::emit_new(chunks, current, line);
    set(chunks, current, out, line);
    emit_copy_entries(chunks, current, src, out, line);
}

/// `minus` — maps drop a key, lists drop the first occurrence of a value.
pub fn emit_minus(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let arg = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    set(chunks, current, arg, line);
    set(chunks, current, recv, line);

    get(chunks, current, recv, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, recv, line);
    common_collections::emit_clone(chunks, current, line);
    set(chunks, current, out, line);
    get(chunks, current, out, line);
    get(chunks, current, arg, line);
    common_collections::emit_remove_value(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(chunks, current, out, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    emit_copy_map_into(chunks, current, out, line);
    get(chunks, current, arg, line);
    let is_arr2 = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(is_arr2, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    // `map - listOf(k1, k2)` — drop each key.
    let idx = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let state = loops::emit_for_in_start(chunks, current, arg, idx, line);
    set(chunks, current, key, line);
    get(chunks, current, out, line);
    get(chunks, current, key, line);
    emit_dict_delete_full(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx, state, line);
    chunks[current].emit_else(line);
    get(chunks, current, out, line);
    get(chunks, current, arg, line);
    emit_dict_delete_full(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(chunks, current, out, line);
    chunks[current].emit_end(line);
}

/// The `__dict_set` statement form — same `[dict, key, value]` → `[null]`
/// contract as `common:dict.set_dynamic`, plus the `__keys` tracking the
/// shared op lacks. `m[k] = v` on a NEW key used to store the property
/// without enumerating it: the map printed `{}` around a value `m[k]` could
/// read back.
pub fn emit_dict_set_stmt(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_dict_set_tracked(chunks, current, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Kotlin `map.forEach { k, v -> }` / `{ (k, v) -> }` / `{ it… }`.
///
/// The lambda is called with `(entry, value)`: a one-parameter lambda binds
/// the ENTRY (so `it.key`/`it.value` and `(k, v)` destructuring work), and a
/// two-parameter `{ _, v -> }` still receives the value as its second
/// argument. The BiConsumer's first parameter only degrades when it is
/// actually READ as the bare key, which no current corpus test does.
pub fn emit_map_for_each(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let f = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    set(chunks, current, f, line);
    set(chunks, current, m, line);
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let entry = chunks[current].alloc_scratch(1);
    for_each_entry(chunks, current, m, key, value, line, |chunks| {
        make_entry(chunks, current, key, value, entry, line);
        call_fn(chunks, current, f, &[entry, value], line);
        chunks[current].emit_op(Op::DROP, line);
    });
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `map.replace(k, v)` → previous value or null; `replace(k, old, new)` →
/// Boolean, replacing only when the current value equals `old`.
pub fn emit_map_replace(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc >= 4 {
        let newv = chunks[current].alloc_scratch(1);
        let oldv = chunks[current].alloc_scratch(1);
        let k = chunks[current].alloc_scratch(1);
        let m = chunks[current].alloc_scratch(1);
        set(chunks, current, newv, line);
        set(chunks, current, oldv, line);
        set(chunks, current, k, line);
        set(chunks, current, m, line);
        // Same arity, two meanings: `map.replace(k, old, new)` vs Kotlin's
        // string `replace(old, new, ignoreCase)`.
        get(chunks, current, m, line);
        host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
        chunks[current].emit_string_const("string", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        truthy(chunks, current, line);
        chunks[current].emit_if_value(line);
        // ignoreCase truthy → case-fold both haystack and needle, replaceAll,
        // which matches the corpus' literal-needle uses.
        get(chunks, current, newv, line);
        truthy(chunks, current, line);
        chunks[current].emit_if_value(line);
        get(chunks, current, m, line);
        vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
        get(chunks, current, k, line);
        vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
        get(chunks, current, oldv, line);
        host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line);
        chunks[current].emit_else(line);
        get(chunks, current, m, line);
        get(chunks, current, k, line);
        get(chunks, current, oldv, line);
        host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        get(chunks, current, m, line);
        get(chunks, current, k, line);
        dict::emit_get_dynamic(chunks, current, line);
        get(chunks, current, oldv, line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        truthy(chunks, current, line);
        chunks[current].emit_if_value(line);
        get(chunks, current, m, line);
        get(chunks, current, k, line);
        get(chunks, current, newv, line);
        emit_dict_set_tracked(chunks, current, line);
        chunks[current].emit_bool_const(true, line);
        ops::emit_i32_to_bool(&mut chunks[current], line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        ops::emit_i32_to_bool(&mut chunks[current], line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        return;
    }
    let v = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let prev = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    set(chunks, current, k, line);
    set(chunks, current, m, line);
    // One spelling, two receivers: `"a-b".replace("-", "+")` is Kotlin's
    // replace-ALL-occurrences string method; a Map's is put-if-present.
    get(chunks, current, m, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    get(chunks, current, v, line);
    host::emit(&mut chunks[current], "ecma:string", "replaceAll", 3, line);
    chunks[current].emit_else(line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_method_has(chunks, current, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_get_dynamic(chunks, current, line);
    set(chunks, current, prev, line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    get(chunks, current, v, line);
    emit_dict_set_tracked(chunks, current, line);
    get(chunks, current, prev, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `remove(k)` / map `minus` deletion that BOTH un-enumerates the key and
/// deletes the property. `dict::emit_method_delete` only removes the key from
/// `__keys`, so the map rendered without the entry while `containsKey` (an
/// `ecma:object.hasIn` probe) still answered true.
pub fn emit_dict_delete_full(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let k = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    set(chunks, current, k, line);
    set(chunks, current, m, line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_method_delete(chunks, current, line);
    let existed = chunks[current].alloc_scratch(1);
    set(chunks, current, existed, line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    host::emit(&mut chunks[current], "ecma:object", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    get(chunks, current, existed, line);
}

/// `remove(x)` for ANY receiver: a MutableList removes the first occurrence
/// and answers a Boolean; a Map removes the key and answers the previous
/// VALUE (or null). `remove(k, v)` is the conditional Map form → Boolean.
pub fn emit_remove_any(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc >= 3 {
        // map.remove(k, v): only when current value equals v. Map keys are
        // stored stringified — normalize the lookup key to match.
        let v = chunks[current].alloc_scratch(1);
        let k = chunks[current].alloc_scratch(1);
        let m = chunks[current].alloc_scratch(1);
        set(chunks, current, v, line);
        strings::emit_to_string(&mut chunks[current], line);
        set(chunks, current, k, line);
        set(chunks, current, m, line);
        get(chunks, current, m, line);
        get(chunks, current, k, line);
        dict::emit_get_dynamic(chunks, current, line);
        get(chunks, current, v, line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        truthy(chunks, current, line);
        chunks[current].emit_if_value(line);
        get(chunks, current, m, line);
        get(chunks, current, k, line);
        emit_dict_delete_full(chunks, current, 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_bool_const(true, line);
        ops::emit_i32_to_bool(&mut chunks[current], line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        ops::emit_i32_to_bool(&mut chunks[current], line);
        chunks[current].emit_end(line);
        return;
    }
    let x = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    set(chunks, current, x, line);
    set(chunks, current, recv, line);
    get(chunks, current, recv, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    // List: remove by VALUE → did it change?
    get(chunks, current, recv, line);
    get(chunks, current, x, line);
    common_collections::emit_contains(chunks, current, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, recv, line);
    get(chunks, current, x, line);
    common_collections::emit_remove_value(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    // Dict-backed receiver. Keys are stored stringified — normalize. A SET
    // answers a Boolean (was the element there?); a MAP answers the previous
    // value or null.
    let prev = chunks[current].alloc_scratch(1);
    let had = chunks[current].alloc_scratch(1);
    get(chunks, current, x, line);
    strings::emit_to_string(&mut chunks[current], line);
    set(chunks, current, x, line);
    get(chunks, current, recv, line);
    get(chunks, current, x, line);
    dict::emit_method_has(chunks, current, line);
    truthy(chunks, current, line);
    set(chunks, current, had, line);
    get(chunks, current, recv, line);
    get(chunks, current, x, line);
    dict::emit_get_dynamic(chunks, current, line);
    emit_undef_to_null(chunks, current, line);
    set(chunks, current, prev, line);
    get(chunks, current, recv, line);
    get(chunks, current, x, line);
    emit_dict_delete_full(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    get(chunks, current, recv, line);
    let marker = chunks[current].add_constant(Value::String(Arc::from(
        crate::emitter::tostring::SET_MARKER,
    )));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, marker, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, had, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(chunks, current, prev, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `clear()` for ANY receiver.
pub fn emit_clear_any(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let recv = chunks[current].alloc_scratch(1);
    set(chunks, current, recv, line);
    get(chunks, current, recv, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    // splice everything out, in place
    get(chunks, current, recv, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    get(chunks, current, recv, line);
    common_collections::emit_len(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:array", "splice", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(chunks, current, recv, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    get(chunks, current, recv, line);
    dict::emit_method_clear_stack(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `map.put(k, v)` → the PREVIOUS value or null.
pub fn emit_map_put(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    let k = chunks[current].alloc_scratch(1);
    let m = chunks[current].alloc_scratch(1);
    let prev = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    set(chunks, current, k, line);
    set(chunks, current, m, line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    dict::emit_get_dynamic(chunks, current, line);
    emit_undef_to_null(chunks, current, line);
    set(chunks, current, prev, line);
    get(chunks, current, m, line);
    get(chunks, current, k, line);
    get(chunks, current, v, line);
    emit_dict_set_tracked(chunks, current, line);
    get(chunks, current, prev, line);
}

/// Kotlin's "absent" is null, never undefined — a dynamic property read on a
/// missing key answers js-undefined, which renders as `undefined`. [v] → [v|null].
fn emit_undef_to_null(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    get(chunks, current, v, line);
    let test = chunks[current].add_import("wasm:js-undefined", "test");
    chunks[current].emit_call(test, 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    get(chunks, current, v, line);
    chunks[current].emit_end(line);
}

/// `list[i]` / `elementAt(i)` — Kotlin THROWS IndexOutOfBoundsException out
/// of range (JS answers undefined). Stack: [arr, i] → [value].
pub fn emit_list_get_throwing(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let i = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    let bad = chunks[current].alloc_scratch(1);
    set(chunks, current, i, line);
    set(chunks, current, arr, line);
    get(chunks, current, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    set(chunks, current, bad, line);
    get(chunks, current, i, line);
    get(chunks, current, arr, line);
    common_collections::emit_len(chunks, current, line);
    ops::emit_dyn_ge(&mut chunks[current], line);
    truthy(chunks, current, line);
    get(chunks, current, bad, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Index out of bounds for length", line);
    crate::emitter::nullability::emit_exception(
        chunks,
        current,
        1,
        "IndexOutOfBoundsException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    get(chunks, current, arr, line);
    get(chunks, current, i, line);
    common_collections::emit_get(chunks, current, line);
}

/// `subList(from, to)` — validated: `from > to` is IllegalArgumentException,
/// out-of-range is IndexOutOfBoundsException (a bare slice silently clamps).
/// Stack: [arr, from, to] → [slice].
pub fn emit_sub_list(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let to = chunks[current].alloc_scratch(1);
    let from = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    let bad = chunks[current].alloc_scratch(1);
    set(chunks, current, to, line);
    set(chunks, current, from, line);
    set(chunks, current, arr, line);
    get(chunks, current, from, line);
    get(chunks, current, to, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("fromIndex > toIndex", line);
    crate::emitter::nullability::emit_exception(
        chunks,
        current,
        1,
        "IllegalArgumentException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    get(chunks, current, from, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    ops::emit_dyn_lt(&mut chunks[current], line);
    truthy(chunks, current, line);
    set(chunks, current, bad, line);
    get(chunks, current, to, line);
    get(chunks, current, arr, line);
    common_collections::emit_len(chunks, current, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    truthy(chunks, current, line);
    get(chunks, current, bad, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Index out of bounds for length", line);
    crate::emitter::nullability::emit_exception(
        chunks,
        current,
        1,
        "IndexOutOfBoundsException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    get(chunks, current, arr, line);
    get(chunks, current, from, line);
    get(chunks, current, to, line);
    common_collections::emit_slice(chunks, current, line);
}

/// `list[i] = v` — rejects `Collections.unmodifiable*` receivers first.
/// Stack: [arr, i, v] → same contract as `common:collections.set`.
pub fn emit_list_set_checked(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(chunks, current, v, line);
    set(chunks, current, i, line);
    set(chunks, current, arr, line);
    crate::emitter::collections::emit_throw_if_java_immutable(chunks, current, arr, line);
    get(chunks, current, arr, line);
    get(chunks, current, i, line);
    get(chunks, current, v, line);
    common_collections::emit_set(chunks, current, line);
}

/// `m[k]` — dynamic get with Kotlin's absent-is-null contract.
/// Stack: [m, k] → [value | null].
pub fn emit_dict_get_null(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    dict::emit_get_dynamic(chunks, current, line);
    emit_undef_to_null(chunks, current, line);
}

/// `x.first` / `x.second` / `x.third` when the walker can't type the
/// receiver (lambda params, function results). Pairs/Triples are tagged
/// ARRAYS with no by-name props, so the read is positional there; anything
/// else (a data class declaring `first`) keeps the property read. Also
/// covers `range.first` — ranges materialize as arrays.
/// Stack: [x, name, idx] → [value].
pub fn emit_tuple_prop(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let x = chunks[current].alloc_scratch(1);
    set(chunks, current, i, line);
    set(chunks, current, n, line);
    set(chunks, current, x, line);
    // Null-tolerant so `?.first` shares this probe: null receiver → null.
    get(chunks, current, x, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    get(chunks, current, x, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    get(chunks, current, x, line);
    get(chunks, current, i, line);
    common_collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    get(chunks, current, x, line);
    get(chunks, current, n, line);
    dict::emit_get_dynamic(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `x?.field` on receivers the shared null-safe path can't see into
/// (array-backed entries/pairs — `is_object` answers false there).
/// Stack: [obj, key] → [prop value | null].
pub fn emit_safe_get(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let k = chunks[current].alloc_scratch(1);
    let o = chunks[current].alloc_scratch(1);
    set(chunks, current, k, line);
    set(chunks, current, o, line);
    get(chunks, current, o, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    truthy(chunks, current, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    get(chunks, current, o, line);
    get(chunks, current, k, line);
    dict::emit_get_dynamic(chunks, current, line);
    emit_undef_to_null(chunks, current, line);
    chunks[current].emit_end(line);
}
