//! Kotlin collection adapters that need Kotlin-specific return contracts.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ValueSource,
};
use vybe_compiler::primitives::{
    collections as common_collections, dict, instructions::host, loops, ops, sets as common_sets,
};
use vybe_runtime::Chunk;
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

const MUTABLE_LIST_MARKER: &str = "__kt_mutable_list";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

pub(crate) fn emit_is_ecma_set(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:object", "toStringTag", 1, line);
    chunks[current].emit_string_const("[object Set]", line);
    host::emit(&mut chunks[current], "wasm:js-string", "equals", 2, line);
}

/// Kotlin `MutableCollection.add(value)`.
///
/// Arrays/lists append and return `true`; Kotlin's dict-backed sets use
/// `MutableSet.add` duplicate semantics and return whether the set changed.
/// Throw UnsupportedOperationException when `slot` carries a
/// `Collections.unmodifiable*`/`singleton` marker (the jvm platform stamps
/// `__java_immutable_list` on lists AND sets, `__java_immutable_map` on maps).
pub fn emit_throw_if_java_immutable(chunks: &mut Vec<Chunk>, current: usize, slot: u16, line: u32) {
    for key in ["__java_immutable_list", "__java_immutable_map"] {
        get(&mut chunks[current], slot, line);
        chunks[current].emit_string_const(key, line);
        host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        chunks[current].emit_string_const("collection is unmodifiable", line);
        crate::emitter::nullability::emit_exception(
            chunks,
            current,
            1,
            "UnsupportedOperationException",
            line,
        );
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
    }
}

pub fn emit_file_walk_max_depth(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let max_depth = chunks[current].alloc_scratch(1);
    let walk = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        set(&mut chunks[current], max_depth, line);
    } else {
        chunks[current].emit_i32_const(i32::MAX, line);
        set(&mut chunks[current], max_depth, line);
    }
    set(&mut chunks[current], walk, line);
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }

    common_collections::emit_array_new(chunks, current, 0, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    let idx = chunks[current].alloc_scratch(1);
    let state = loops::emit_for_in_start(chunks, current, walk, idx, line);
    let item = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], item, line);
    get(&mut chunks[current], item, line);
    chunks[current].emit_string_const("__java_io_depth", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    let depth = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], depth, line);
    get(&mut chunks[current], max_depth, line);
    get(&mut chunks[current], depth, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], item, line);
    common_collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx, state, line);
    get(&mut chunks[current], out, line);
}

/// Variadic `kotlin.collections.setOf(...)` / `mutableSetOf(...)`.
///
/// Tree-resolved calls arrive as ordinary variadic calls; lower them to the
/// shared set literal primitive.
pub fn emit_set_literal(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // A Kotlin set IS a `java.util.LinkedHashSet`-shaped JVM set: construct
    // through the platform so every set in the program shares ONE
    // representation. The old dict-backed shape (render-string keys) is gone —
    // it was a second implementation of the same Java class.
    //
    // The ELEMENTS are packed into one array first: `setOf(x)` is a
    // single-element set, while the platform ctor's 1-arg form is Java's
    // `HashSet(Collection)` and would read `x` as an iterable.
    let base = if argc == 0 {
        0
    } else {
        chunks[current].alloc_scratch(argc as u16)
    };
    common_collections::emit_pack_n(chunks, current, argc as u16, base, line);
    vybe_platform_jvm::emitter::dispatch::dispatch(
        "jvm.java.hash_set_new",
        chunks,
        current,
        1,
        line,
    );
}

/// Variadic `kotlin.collections.hashSetOf(...)`.
pub fn emit_hash_set_literal(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_set_literal(chunks, current, argc, line);
}

/// Kotlin dict-backed set `size`. The marker property is implementation detail.
pub fn emit_set_size(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    // `java.util.HashSet` (and other jvm sets) are ARRAY-backed — the dict
    // size probe threw on them.
    let v = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    get(&mut chunks[current], v, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], v, line);
    common_collections::emit_len(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], v, line);
    emit_is_ecma_set(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], v, line);
    common_sets::emit_size(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], v, line);
    dict::emit_method_size(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Kotlin set membership. Stack in: `[set, value]`; stack out: `[bool]`.
pub fn emit_set_has(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], set_slot, line);

    get(&mut chunks[current], set_slot, line);
    emit_is_ecma_set(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    // A Kotlin set IS a JVM set — membership is the snapshot-keyed
    // (hashCode/equals) contract the platform's sets declare.
    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], value, line);
    common_sets::emit_has_snapshot(chunks, current, line);
    chunks[current].emit_else(line);

    get(&mut chunks[current], set_slot, line);
    get(&mut chunks[current], value, line);
    crate::emitter::tostring::emit_to_string(chunks, current, line);
    dict::emit_method_has(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `CharArray.concatToString()` — the chars, joined bare.
pub fn emit_concat_to_string(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    chunks[current].emit_string_const("", line);
    common_collections::emit_join(chunks, current, line);
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
    chunks[current].emit_string_const(crate::emitter::tostring::SET_MARKER, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_dict_backed_set_values_array(chunks, current, collection, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], collection, line);
    emit_is_ecma_set(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], collection, line);
    common_sets::emit_values_array(chunks, current, line);
    chunks[current].emit_else(line);
    emit_dict_backed_set_values_array(chunks, current, collection, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_dict_backed_set_values_array(
    chunks: &mut Vec<Chunk>,
    current: usize,
    collection: u16,
    line: u32,
) {
    // From `__keys`, NOT `ecma:object.entries`: JS enumerates integer-like
    // keys ASCENDING, so `linkedSetOf(4, 1, 3)` iterated sorted instead of
    // in insertion order. A set LITERAL's `__keys` also lists the
    // `__kt_set` marker (and `withDefault` maps their provider) — skipped.
    {
        let keys = chunks[current].alloc_scratch(1);
        let result = chunks[current].alloc_scratch(1);
        let idx = chunks[current].alloc_scratch(1);
        let key = chunks[current].alloc_scratch(1);
        get(&mut chunks[current], collection, line);
        dict::emit_keys(chunks, current, line);
        set(&mut chunks[current], keys, line);
        common_collections::emit_array_new(chunks, current, 0, line);
        set(&mut chunks[current], result, line);
        let state = loops::emit_for_in_start(chunks, current, keys, idx, line);
        set(&mut chunks[current], key, line);
        get(&mut chunks[current], key, line);
        chunks[current].emit_string_const(crate::emitter::tostring::SET_MARKER, line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        get(&mut chunks[current], key, line);
        chunks[current].emit_string_const("__kt_default", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_op(Op::I32_OR, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        get(&mut chunks[current], result, line);
        get(&mut chunks[current], collection, line);
        get(&mut chunks[current], key, line);
        dict::emit_get_dynamic(chunks, current, line);
        common_collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
        loops::emit_for_in_end(chunks, current, idx, state, line);
        get(&mut chunks[current], result, line);
    }
}

/// Kotlin `toList()` / `toTypedArray()` for arrays and dict-backed sets.
pub fn emit_to_list(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let collection = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], collection, line);
    emit_collection_values_array(chunks, current, collection, line);
}

/// Kotlin `mutableListOf` / `arrayListOf`: same array storage as List, stamped
/// so `is MutableList` and `as? MutableList` can distinguish readonly factories.
pub fn emit_to_mutable_list(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    emit_to_list(chunks, current, argc, line);
    let out = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], out, line);
    get(&mut chunks[current], out, line);
    let marker = class_slots::resolve_interned(&mut chunks[current], &ClassSlot::internal(MUTABLE_LIST_MARKER), &PlainNames);
    chunks[current].emit_bool_const(true, line);
    class_slots::emit_class_set(&mut chunks[current], ObjSource::Stack, &marker, ValueSource::Stack, line);
    get(&mut chunks[current], out, line);
}

pub fn emit_is_mutable_list(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    get(&mut chunks[current], value, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal((MUTABLE_LIST_MARKER).to_string()), &PlainNames);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

/// Kotlin `toSet()` / `toMutableSet()` for arrays and dict-backed sets.
pub fn emit_to_set(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_to_hash_set(chunks, current, 1, line);
}

/// Kotlin `toHashSet()` / `toSet()` — a real JVM set, built by the platform.
/// The operand may be any collection shape (`union` is declared on
/// `Iterable`), so it is flattened to a values array first.
pub fn emit_to_hash_set(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let collection = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], collection, line);
    emit_collection_values_array(chunks, current, collection, line);
    vybe_platform_jvm::emitter::dispatch::dispatch(
        "jvm.java.hash_set_new",
        chunks,
        current,
        1,
        line,
    );
}

/// `intersect` / `subtract`: keep each element of the LEFT collection whose
/// render key is (`keep_present`) or is not present on the right.
///
/// Stack: `[left, right] -> [kotlin set]`. Iterating the left operand is what
/// gives Kotlin's documented ordering — the result follows the receiver.
///
/// This walks the dict backing on both sides rather than handing the operands
/// to `common_sets`, because a Kotlin set is NOT an ECMA Set: `emit_to_hash_set`
/// builds a `dict` keyed by the Kotlin render string and stamped with
/// `SET_MARKER`. Calling `ecma:set.intersection` on one answers an empty Set —
/// silently, since a plain object is a legal argument.
fn emit_set_from_filter(chunks: &mut Vec<Chunk>, current: usize, keep_present: bool, line: u32) {
    let right = chunks[current].alloc_scratch(1);
    let left = chunks[current].alloc_scratch(1);
    let right_values = chunks[current].alloc_scratch(1);
    let values = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    set(&mut chunks[current], right, line);
    set(&mut chunks[current], left, line);
    emit_collection_values_array(chunks, current, right, line);
    set(&mut chunks[current], right_values, line);
    emit_collection_values_array(chunks, current, left, line);
    set(&mut chunks[current], values, line);
    common_collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);
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
    get(&mut chunks[current], right_values, line);
    get(&mut chunks[current], value, line);
    common_collections::emit_contains(chunks, current, line);
    if !keep_present {
        ops::emit_dyn_not(&mut chunks[current], line);
    }
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], out, line);
    get(&mut chunks[current], value, line);
    common_collections::emit_push(chunks, current, line);
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
    // The filtered values, as ONE JVM set (platform representation).
    get(&mut chunks[current], out, line);
    vybe_platform_jvm::emitter::dispatch::dispatch(
        "jvm.java.hash_set_new",
        chunks,
        current,
        1,
        line,
    );
}

/// Kotlin `union`: every element of the left operand in order, then each
/// element of the right that is not already there.
///
/// Concatenating the two value arrays and re-running `emit_to_hash_set` gives
/// exactly that — it already dedupes by the Kotlin render key, keeps insertion
/// order, and stamps `SET_MARKER` — so the one representation of a Kotlin set
/// stays in one place. The operands may be lists as well as sets (Kotlin
/// declares `union` on `Iterable`), which `emit_collection_values_array`
/// already normalizes.
pub fn emit_set_union(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let right = chunks[current].alloc_scratch(1);
    let left = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], right, line);
    set(&mut chunks[current], left, line);
    emit_collection_values_array(chunks, current, left, line);
    emit_collection_values_array(chunks, current, right, line);
    common_collections::emit_concat(chunks, current, line);
    emit_to_hash_set(chunks, current, 1, line);
}

pub fn emit_set_intersect(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_set_from_filter(chunks, current, true, line);
}

pub fn emit_set_subtract(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    emit_set_from_filter(chunks, current, false, line);
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
    emit_is_ecma_set(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], value, line);
    common_sets::emit_size(chunks, current, line);
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
    // StringBuilder: empty when its `__buffer` is.
    get(&mut chunks[current], value, line);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal(("__buffer").to_string()), &PlainNames);
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &cs_slot, Dest::Stack, line);
    let buf = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], buf, line);
    get(&mut chunks[current], buf, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], buf, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
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
        let cs_slot = class_slots::resolve(&ClassSlot::Internal((prop).to_string()), &PlainNames);
        class_slots::emit_class_set(&mut chunks[current], ObjSource::Stack, &cs_slot, ValueSource::Stack, line);
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
    // CLONE, not the array itself: `arr.toList()` is a snapshot in Kotlin —
    // mutating the source must not show through.
    get(&mut chunks[current], v, line);
    common_collections::emit_clone(chunks, current, line);
    chunks[current].emit_else(line);

    // A Set is its values; a Map is its entries.
    get(&mut chunks[current], v, line);
    emit_is_ecma_set(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], v, line);
    common_sets::emit_values_array(chunks, current, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], v, line);
    let marker = class_slots::resolve_interned(
        &mut chunks[current],
        &ClassSlot::internal(crate::emitter::tostring::SET_MARKER),
        &PlainNames,
    );
    class_slots::emit_class_get(&mut chunks[current], ObjSource::Stack, &marker, Dest::Stack, line);
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
    let state =
        vybe_compiler::primitives::loops::emit_for_in_start(chunks, current, keys, idx, line);
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

/// `reverse()` — IN PLACE, answers Unit (`reversed()`/`reversedArray()` are
/// the copying forms).
pub fn emit_reverse_in_place(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    common_collections::emit_reverse(chunks, current, line);
}

/// `sortedArray()` — sorted COPY.
pub fn emit_sorted_copy(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    common_collections::emit_clone(chunks, current, line);
    common_collections::emit_sort(chunks, current, line);
}

/// `sortedArrayDescending()` — descending sorted COPY.
pub fn emit_sorted_desc_copy(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    common_collections::emit_clone(chunks, current, line);
    common_collections::emit_sort(chunks, current, line);
    common_collections::emit_reverse(chunks, current, line);
}

/// `sort(fromIndex, toIndex)` — sort the RANGE in place.
/// Stack: [arr, from, to] → [arr].
pub fn emit_sort_range(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let to = chunks[current].alloc_scratch(1);
    let from = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    let piece = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], to, line);
    set(&mut chunks[current], from, line);
    set(&mut chunks[current], arr, line);
    get(&mut chunks[current], arr, line);
    get(&mut chunks[current], from, line);
    get(&mut chunks[current], to, line);
    common_collections::emit_slice(chunks, current, line);
    common_collections::emit_sort(chunks, current, line);
    set(&mut chunks[current], piece, line);
    // write the sorted run back: arr[from + i] = piece[i]
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], i, line);
    let _block = chunks[current].emit_block(line);
    let (_loop, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], piece, line);
    common_collections::emit_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    get(&mut chunks[current], arr, line);
    get(&mut chunks[current], from, line);
    get(&mut chunks[current], i, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    get(&mut chunks[current], piece, line);
    get(&mut chunks[current], i, line);
    common_collections::emit_get(chunks, current, line);
    common_collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    get(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], arr, line);
}

/// `binarySearch(element, fromIndex, toIndex)` — Kotlin's member order vs
/// Java's static `(a, from, to, key)`; reorder and delegate.
/// Stack: [arr, key, from, to] → [index | -(insertion)-1].
pub fn emit_binary_search_range(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 4 {
        // 1-arg form: [arr, key] — Java's order already.
        vybe_platform_jvm::emitter::arrays_adapter::emit_binary_search(chunks, current, argc, line);
        return;
    }
    let to = chunks[current].alloc_scratch(1);
    let from = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let arr = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], to, line);
    set(&mut chunks[current], from, line);
    set(&mut chunks[current], key, line);
    set(&mut chunks[current], arr, line);
    get(&mut chunks[current], arr, line);
    get(&mut chunks[current], from, line);
    get(&mut chunks[current], to, line);
    get(&mut chunks[current], key, line);
    vybe_platform_jvm::emitter::arrays_adapter::emit_binary_search(chunks, current, 4, line);
}

/// `toByteArray()` for ANY receiver — strings via the jvm encoder, arrays as
/// a plain copy.
pub fn emit_to_byte_array(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    get(&mut chunks[current], v, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], v, line);
    vybe_platform_jvm::emitter::io_adapter::emit_string_to_byte_array(chunks, current, argc, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], v, line);
    common_collections::emit_clone(chunks, current, line);
    chunks[current].emit_end(line);
}
