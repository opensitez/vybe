//! Shared sorted set / sorted map primitives (comparison-ordered collections).
//!
//! Sibling to [`crate::compiler::heap`]. The representations are ECMA-native:
//!
//! * A sorted **set** is an `ecma` array kept ordered by [`emit_sorted_add`]
//!   (dedupe-via-contains, push, then re-sort). An optional comparator function
//!   lives under [`COMPARATOR_KEY`]; absent it the type-aware default ordering
//!   (`__vybe_sort_in_place`) applies.
//! * A sorted **map** is an `ecma:map` whose key / value / entry views sort at
//!   read time (the map itself keeps insertion order; ordering is imposed on
//!   enumeration).
//!
//! Consumed by Java (`TreeSet`/`TreeMap`) and dotnet (`SortedSet`/
//! `SortedDictionary`); available to any comparison-ordered surface. This is the
//! single definition site for the comparator/flag prop keys so a shared writer
//! and a language-local reader cannot drift apart.
//!
//! NOTE: the prop-key string *values* retain a `__java_` prefix so the initial
//! Java lift stays byte-identical. Renaming the values is a separate verified
//! pass — a value change silently breaks the key agreement between writers and
//! readers, so it must not be bundled with the lift.

use crate::compiler::collections;
use crate::compiler::instructions::{core_wasm, host};
use crate::compiler::ops;
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Property key holding the (optional) comparator function on a sorted set/map.
pub const COMPARATOR_KEY: &str = "__java_comparator";
/// Marks an array as set-shaped (dedupe semantics) rather than a plain list.
pub const SET_COLLECTION_KEY: &str = "__java_set_collection";
/// Marks a sorted set whose stored order is descending.
pub const DESCENDING_SET_KEY: &str = "__java_descending_set";
/// Marks a sorted map whose enumeration order is descending.
pub const DESCENDING_MAP_KEY: &str = "__java_descending_map";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Push the comparator function stored on `value` (or null). Stack: `-> [cmp]`.
pub fn emit_comparator(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    get(&mut chunks[current], value, line);
    chunks[current].emit_string_const(COMPARATOR_KEY, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
}

/// Sort the array in `value` in place according to its comparator (or the
/// type-aware default), honoring the descending flag.
pub fn emit_sort_if_ordered(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    let comparator = chunks[current].alloc_scratch(1);
    emit_comparator(chunks, current, value, line);
    set(&mut chunks[current], comparator, line);
    get(&mut chunks[current], comparator, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("function", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    collections::emit_reverse(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

/// Tag the array on the stack as set-shaped. Stack: `[list] -> [list]`.
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

/// Build a new sorted collection. When `map` is false the backing is an array
/// tagged as a set; when true it is an `ecma:map`. A comparator argument (or
/// null when `argc == 0`) is stashed under [`COMPARATOR_KEY`].
/// Stack in: `[comparator?]` -> `[collection]`.
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

/// Add `value` to a sorted set, keeping it ordered and deduped. Pushes the
/// standard set-`add` boolean (true if inserted, false if already present).
/// Stack in: `[list, value]` -> `[bool]`.
pub fn emit_sorted_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let list = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], list, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], value, line);
    collections::emit_contains(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
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

/// Return the minimum (`last == false`) or maximum (`last == true`) element of
/// a sorted set. Stack in: `[list]` -> `[element]`.
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

/// Comparison for the `ceiling`/`floor`/`higher`/`lower` family:
/// mode 0 = `key >= bound`, 1 = `key <= bound`, 2 = `key > bound`, else `key < bound`.
pub fn emit_bound_condition(
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
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
        }
        1 => {
            ops::emit_dyn_gt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
        }
        2 => ops::emit_dyn_gt(&mut chunks[current], line),
        _ => ops::emit_dyn_lt(&mut chunks[current], line),
    }
}

fn emit_upper_condition(
    chunks: &mut [Chunk],
    current: usize,
    key: u16,
    upper: Option<u16>,
    upper_inclusive: bool,
    line: u32,
) {
    if let Some(upper) = upper {
        get(&mut chunks[current], key, line);
        get(&mut chunks[current], upper, line);
        if upper_inclusive {
            // key <= upper
            ops::emit_dyn_gt(&mut chunks[current], line);
            ops::emit_dyn_not(&mut chunks[current], line);
        } else {
            // key < upper
            ops::emit_dyn_lt(&mut chunks[current], line);
        }
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_bool_const(true, line);
    }
}

/// Range predicate: `lower <= key` (inclusive lower) combined with the upper
/// bound per `upper_inclusive`. Either bound may be absent.
pub fn emit_range_condition(
    chunks: &mut [Chunk],
    current: usize,
    key: u16,
    lower: Option<u16>,
    upper: Option<u16>,
    upper_inclusive: bool,
    line: u32,
) {
    if let Some(lower) = lower {
        get(&mut chunks[current], key, line);
        get(&mut chunks[current], lower, line);
        ops::emit_dyn_lt(&mut chunks[current], line);
        ops::emit_dyn_not(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        emit_upper_condition(chunks, current, key, upper, upper_inclusive, line);
        chunks[current].emit_else(line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_end(line);
    } else {
        emit_upper_condition(chunks, current, key, upper, upper_inclusive, line);
    }
}

/// Materialize a new set holding the elements of a sorted set within a range.
/// mode 0 = `[lower, upper]`, 1 = upper bound only, 2 = lower bound only.
/// `upper_inclusive` selects `<=` (dotnet `GetViewBetween`) vs `<` (Java
/// `subSet`/`headSet`) for the upper bound.
/// Stack in (mode 0): `[list, lower, upper]` -> `[view]`.
pub fn emit_sorted_set_range_view(
    chunks: &mut [Chunk],
    current: usize,
    mode: u8,
    upper_inclusive: bool,
    line: u32,
) {
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
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], list, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], value, line);
    emit_range_condition(chunks, current, value, lower, upper, upper_inclusive, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
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
    get(&mut chunks[current], out, line);
    emit_mark_set_collection(chunks, current, line);
}

/// Return the map's keys as an array in comparator order (honoring the
/// descending-map flag). Stack in: `[map]` -> `[keys_array]`.
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
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], keys, line);
    collections::emit_reverse(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], keys, line);
}

/// Return the map's values as an array ordered by their keys' comparator order.
/// Stack in: `[map]` -> `[values_array]`.
pub fn emit_sorted_map_values(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    emit_sorted_map_key_set(chunks, current, line);
    set(&mut chunks[current], keys, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], keys, line);
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
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], keys, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    collections::emit_push(chunks, current, line);
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
    get(&mut chunks[current], out, line);
}

/// Return the map's entries as an array of `[key, value]` pairs in comparator
/// key order. Stack in: `[map]` -> `[entries_array]`.
pub fn emit_sorted_map_entries(chunks: &mut [Chunk], current: usize, line: u32) {
    let map = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], map, line);
    get(&mut chunks[current], map, line);
    emit_sorted_map_key_set(chunks, current, line);
    set(&mut chunks[current], keys, line);
    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], keys, line);
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
    get(&mut chunks[current], keys, line);
    get(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], key, line);
    get(&mut chunks[current], out, line);
    // pair = [key, value]
    get(&mut chunks[current], key, line);
    get(&mut chunks[current], map, line);
    get(&mut chunks[current], key, line);
    host::emit(&mut chunks[current], "ecma:map", "get", 2, line);
    collections::emit_array_new(chunks, current, 2, line);
    collections::emit_push(chunks, current, line);
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
    get(&mut chunks[current], out, line);
}
