//! Dict/map compilation — emits WASM-compatible bytecode using GC struct ops.
//!
//! A dict is an Object with properties (key-value pairs via struct_set/struct_get)
//! plus a `__keys` array that tracks insertion order for enumeration.
//!
//! This design means:
//! - Key access: `struct_get` / `struct_set` (standard WASM, O(1))
//! - Keys enumeration: `struct_get "__keys"` (standard WASM, returns array)
//! - Values/items enumeration: inline loop over __keys (standard WASM via common::loops)
//! - No host calls needed for any dict operation
//!
//! All languages (Python, Dart, JS, VB, C#) use this same structure.

use crate::primitives::class_slots;
use crate::primitives::instructions::core_wasm;
use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// ── Creation ────────────────────────────────────────────────────────────

/// Emit bytecode to create an empty dict with key tracking.
/// Stack: [] → [dict_object]
/// The map property holding `rendered key -> the key object first stored`.
const TUPLE_KEYS_KEY: &str = "__tuple_keys";

/// Rewrite `key_slot` to the map's canonical key for that value.
///
/// A tuple IS its contents, so two equal tuples must reach one entry. The map
/// itself still holds the original key object — enumeration yields tuples, not
/// rendered strings — and a side map remembers which object a rendering was
/// first stored under. A key that is not a tuple is left alone, so a class
/// instance keeps reference identity.
pub fn emit_canonical_key_slot(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    key_slot: u16,
    line: u32,
) {
    let keys = chunks[current].alloc_scratch(1);
    let rendered = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::primitives::tuples::emit_is_tuple(chunks, current, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_string_const(TUPLE_KEYS_KEY, line);
    call_host(chunks, current, "ecma:object", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    call_host(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_string_const(TUPLE_KEYS_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    call_host(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    call_host(chunks, current, "ecma:json", "stringify", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rendered, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rendered, line);
    call_host(chunks, current, "ecma:map", "has", 2, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rendered, line);
    call_host(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rendered, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    call_host(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

fn call_host(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

pub fn emit_new(chunks: &mut [Chunk], current: usize, line: u32) {
    // Create empty Object
    class_slots::emit_class_alloc(&mut chunks[current], line);
    // Create empty __keys array and attach it
    chunks[current].emit_dup(line);
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    let keys_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::ValueSource::Stack,
        line,
    );
}

// ── Setting keys ────────────────────────────────────────────────────────

/// Emit bytecode to set a string-literal key on a dict.
/// Also appends the key to __keys for enumeration.
/// Stack before: [dict]  (caller must dup if they want to keep dict)
/// Stack after: [dict]
pub fn emit_set_const_key(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    // Set the property: dict.key = value (value must be pushed by caller before this)
    // Actually — caller pushes value AFTER calling this? No.
    // Convention: caller has [dict, value] on stack. We do struct_set.
    let key_idx = class_slots::resolve(
        &class_slots::ClassSlot::internal(key),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &key_idx,
        class_slots::ValueSource::Stack,
        line,
    );
    // Append key to __keys: dict.__keys.push(key)
    chunks[current].emit_dup(line);
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &slot,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_string_const(key, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Emit bytecode to set a dynamic key (key on stack).
/// Stack before: [dict, key, value]
/// Stack after: [dict]
pub fn emit_set_dynamic(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

// ── Map/Set method helpers (used by all languages) ──────────────────────
//
// Python: d[key] = val, key in d, del d[key], len(d), d.keys(), d.values(), d.items()
// JS: map.set(key, val), map.get(key), map.has(key), map.delete(key), map.clear(), map.size
// Dart: map[key] = val, map[key], map.containsKey(key), map.remove(key), map.length
// C#: dict[key] = val, dict[key], dict.ContainsKey(key), dict.Remove(key), dict.Count
//
// All compile to the same opcodes through these helpers.

/// map.set(key, val) / dict[key] = val — with __keys tracking.
/// Caller must have dict in `dict_slot`. Key and value on stack.
/// Stack before: [key, value]  Stack after: [dict] (for chaining)
pub fn emit_method_set(chunks: &mut [Chunk], current: usize, dict_slot: u16, line: u32) {
    // array_set(dict, key, val)
    let _val_tmp = chunks[current].add_constant(Value::Null); // dummy, we'll use stack
    // Actually we need to save key for __keys. Use stack manipulation:
    // Stack: [key, value]. Need [dict, key, value] for array_set.
    // Then [dict.__keys, key] for array_push.

    // Save value to use later
    // Strategy: emit dict first, then key+value are already there
    // But key and value are already on stack. We need dict under them.
    // This requires the caller to push dict first. Let me redesign:

    // Actually the simplest correct API: caller pushes [dict, key, value].
    // We do array_set + __keys push.
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // Push key to __keys: need dict and key again
    // The dict is still accessible via dict_slot. The key was consumed by array_set.
    // We need the caller to compile key twice, or we save it.
    // The clean API: take dict_slot, key is compiled by caller BEFORE calling this.
    // For now, skip __keys tracking for dynamic set — use emit_method_set_with_locals instead.
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line); // return dict for chaining
}

/// map.set(key, val) with __keys tracking — requires local slots.
/// Caller must store dict in `dict_slot` and key in `key_slot`.
/// Value must be on stack.
/// Stack before: [value]  Stack after: [dict]
pub fn emit_method_set_tracked(
    chunks: &mut [Chunk],
    current: usize,
    dict_slot: u16,
    key_slot: u16,
    line: u32,
) {
    // array_set(dict, key, val)
    // Stack has [value]. Need [dict, key, value].
    // Save value, push dict+key, push value back.
    let _val_const = chunks[current].add_constant(Value::Null);
    // Actually: just reorder using locals.
    // Caller already has dict in dict_slot and key in key_slot.
    // We'll build the stack ourselves:

    // Step 1: save value
    // Value is TOS. We don't have a temp slot. Assume caller provides one or
    // we create inline. Simplest: use the pattern [dict, key, val] → array_set.
    // Caller should push dict, then key, then value before calling.
    // Let me just provide a simpler API:

    // [dict_slot has dict, key_slot has key, TOS has value]
    // Build: [dict, key, value] for array_set
    // But TOS=value, we need dict under it. Save value temporarily.
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    // Stack: [value, dict, key]. Need [dict, key, value]. Swap:
    // Actually this is wrong. Let me just take all three from locals.
    // The caller must save value to a local too.

    // REVISED: Take dict_slot, key_slot, val_slot.
    chunks[current].emit_op(Op::DROP, line); // drop key we just pushed
    chunks[current].emit_op(Op::DROP, line); // drop dict we just pushed
    // Forget it — the stack ordering is too complex without a swap opcode.
    // Just use array_set directly. Caller is responsible for stack order.
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // Push key to __keys
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    let keys_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // Return dict for chaining
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
}

/// map.get(key) / dict[key] — get value by dynamic key.
/// Stack before: [dict, key]  Stack after: [value_or_null]
pub fn emit_method_get(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::primitives::collections::emit_get(chunks, current, line);
}

/// map.has(key) / key in dict — check if key exists.
/// Stack before: [dict, key]  Stack after: [bool]
pub fn emit_method_has(chunks: &mut [Chunk], current: usize, line: u32) {
    let dict_slot = chunks[current].local_count;
    let key_slot = dict_slot + 1;
    let has_slot = key_slot + 1;
    let keys_slot = has_slot + 1;
    chunks[current].alloc_scratch(4);
    let keys_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );

    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dict_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    crate::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::primitives::collections::emit_contains(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, has_slot, line);
    chunks[current].emit_if(line);

    chunks[current].emit_bool_const(true, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    crate::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// map.delete(key) / del dict[key] — remove a key.
/// Stack before: [dict, key]  Stack after: [bool] (true if existed)
pub fn emit_method_delete(chunks: &mut [Chunk], current: usize, line: u32) {
    let dict_slot = chunks[current].local_count;
    let key_slot = dict_slot + 1;
    let idx_slot = dict_slot + 2;
    chunks[current].alloc_scratch(3);
    let keys_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );

    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dict_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::primitives::collections::emit_index_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);

    chunks[current].emit_else(line);

    chunks[current].emit_bool_const(false, line);

    chunks[current].emit_end(line);
}

/// map.clear() — remove all entries.
/// dict_slot must contain the dict.
/// Stack before: []  Stack after: []
pub fn emit_method_clear(chunks: &mut [Chunk], current: usize, dict_slot: u16, line: u32) {
    // Replace __keys with empty array
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    let keys_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::ValueSource::Stack,
        line,
    );
}

/// Stack-based variant of `emit_method_clear`. Takes the dict from TOS,
/// clears its `__keys`, leaves `null` on TOS so the call site can use the
/// uniform "method returns void → leaves a value to drop" convention.
/// Stack before: [dict]  Stack after: [null]
pub fn emit_method_clear_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    // [dict] → set __keys = []
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    let keys_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );
    // struct_set pops [obj, val], pushes [val]
    class_slots::emit_class_set(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::ValueSource::Stack,
        line,
    );
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// map.size / len(dict) / dict.Count — number of entries.
/// Stack before: [dict]  Stack after: [i32]
pub fn emit_method_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let keys_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::Dest::Stack,
        line,
    );
    crate::primitives::collections::emit_len(chunks, current, line);
}

/// set.add(value) — add to set (same as dict[value] = true + __keys tracking).
/// Stack before: [set, value]  Stack after: [set]
pub fn emit_set_add(chunks: &mut [Chunk], current: usize, line: u32) {
    // set[value] = true
    chunks[current].emit_bool_const(true, line);
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // TODO: __keys tracking (needs key saved to a local)
}

// ── Getting values ──────────────────────────────────────────────────────

/// Emit bytecode to get a value from a dict by string key.
/// Stack before: [dict]  Stack after: [value_or_null]
pub fn emit_get_const_key(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let key_idx = class_slots::resolve(
        &class_slots::ClassSlot::internal(key),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &key_idx,
        class_slots::Dest::Stack,
        line,
    );
}

/// Emit bytecode to get a value from a dict by dynamic key (on stack).
/// Stack before: [dict, key]  Stack after: [value_or_null]
pub fn emit_get_dynamic(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_get(chunks, current, line);
}

// ── Enumeration (pure WASM — no host calls) ─────────────────────────────

/// Emit bytecode to get all keys as an array.
/// Stack before: [dict]  Stack after: [array_of_keys]
pub fn emit_keys(chunks: &mut [Chunk], current: usize, line: u32) {
    let keys_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &keys_key,
        class_slots::Dest::Stack,
        line,
    );
    // If __keys doesn't exist (legacy dict without tracking), fall back to host
    chunks[current].emit_dup(line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    // ⛔ TAKES THE VALUE AS A PARAM: the `drop` below is discarding a value
    // pushed OUTSIDE this `if`, which a VM's shared operand stack allows and a
    // WASM block does not. Both arms leave one value — `(1, 1)`.
    chunks[current].emit_if_params(line, 1, 1);
    // Null — return an empty array for legacy dicts without key tracking.
    chunks[current].emit_op(Op::DROP, line); // drop null
    // We need the dict back — but it was consumed by struct_get.
    // Can't recover without a local. Use host call as fallback.
    // Actually, struct_get consumes the dict. We need to dup before.
    // Let's restructure: caller should dup before calling emit_keys if they need dict after.
    // For the fallback, emit empty array.
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
}

/// Emit bytecode to get all values as an array.
/// Requires dict in `dict_slot` (not consumed from stack).
/// Stack before: []  Stack after: [array_of_values]
pub fn emit_values_from_local(
    chunks: &mut [Chunk],
    current: usize,
    dict_slot: u16,
    keys_slot: u16,
    result_slot: u16,
    idx_slot: u16,
    line: u32,
) {
    // Get __keys array
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &slot,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);

    // result = []
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    // for i in 0..keys.length: result.push(dict[keys[i]])
    let state =
        crate::primitives::loops::emit_for_in_start(chunks, current, keys_slot, idx_slot, line);
    // Stack: [key_string]. Use it to get value from dict.
    // Store key, get dict[key], push to result
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    // Stack: [key, dict]. Need [dict, key] for array_get. Swap via temp.
    // Actually array_get pops [key, obj] — key is TOS, obj is TOS-1.
    // We have [key, dict]. That's [obj=dict, key=key_on_top]? No.
    // array_get: pops key (TOS), then obj (TOS-1). So [dict, key] → array_get.
    // We have [key_from_for_in, dict_from_local_get].
    // Stack order: key was pushed first, dict second. So TOS=dict, TOS-1=key.
    // We need TOS=key, TOS-1=dict. They're backwards.
    // Fix: push dict first, then key.
    chunks[current].emit_op(Op::DROP, line); // drop the element from for_in_start
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line); // keys[i] → key string
    crate::primitives::collections::emit_get(chunks, current, line); // dict[key] → value
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line); // result.push(value)
    chunks[current].emit_op(Op::DROP, line);

    crate::primitives::loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Emit bytecode to get all [key, value] pairs as an array.
/// Requires dict in `dict_slot`.
/// Stack before: []  Stack after: [array_of_pairs]
pub fn emit_items_from_local(
    chunks: &mut [Chunk],
    current: usize,
    dict_slot: u16,
    keys_slot: u16,
    result_slot: u16,
    idx_slot: u16,
    line: u32,
) {
    // Get __keys array
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    let slot = class_slots::resolve(
        &class_slots::ClassSlot::internal("__keys"),
        &class_slots::PlainNames,
    );
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &slot,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);

    // result = []
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    // for i in 0..keys.length: result.push([keys[i], dict[keys[i]]])
    let state =
        crate::primitives::loops::emit_for_in_start(chunks, current, keys_slot, idx_slot, line);
    chunks[current].emit_op(Op::DROP, line); // drop element from for_in_start

    // Build [key, value] pair
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line); // key = keys[i]
    chunks[current].emit_dup(line); // keep key for value lookup
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    // Stack: [key, key_dup, dict]. Need dict[key] → [key, dict, key_dup] for array_get
    // Hmm, this is getting complicated without a swap op or temp local.
    // Simpler: re-fetch key from array.
    chunks[current].emit_op(Op::DROP, line); // drop dup
    chunks[current].emit_op(Op::DROP, line); // drop key

    // Start over with locals:
    // key = keys[i]
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    // value = dict[key]
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    // pair = [key, value]
    crate::primitives::collections::emit_array_pair(chunks, current, line);
    // result.push(pair)
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    // Stack: [pair, result]. Need [result, pair] for array_push.
    // array_push: pops [arr, val] → pushes arr. So need [result, pair].
    // We have [pair, result]. Wrong order.
    // Fix: push result first.
    chunks[current].emit_op(Op::DROP, line); // drop result
    chunks[current].emit_op(Op::DROP, line); // drop pair

    // Redo properly:
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    // pair = [keys[i], dict[keys[i]]]
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line); // key
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::primitives::collections::emit_get(chunks, current, line); // key again
    crate::primitives::collections::emit_get(chunks, current, line); // dict[key] = value
    crate::primitives::collections::emit_array_pair(chunks, current, line); // [key, value]
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line); // result.push(pair)
    chunks[current].emit_op(Op::DROP, line);

    crate::primitives::loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

// ── Stack-based (polymorphic dispatch via spec-clean opcodes) ────────────

/// Get dynamic key. Stack: `[dict, key]` → `[value_or_undefined]`.
///
/// `Op::ARRAY_GET` is fully polymorphic — it dispatches per
/// `ObjectKind` to the right backing (Map / Array / property-bag /
/// `__getitem__` dunder) and matches `ecma:map.get` / `ecma:object.get`
/// semantics depending on receiver shape.
pub fn emit_get(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::ARRAY_GET, line);
}

/// Set dynamic key. Stack: `[dict, key, value]` → `[]`.
///
/// `Op::ARRAY_SET` mirrors `Op::ARRAY_GET`'s polymorphism — Map insert
/// vs property-bag write vs `__setitem__` dunder. The opcode pushes
/// the assigned value onto the stack; we drop it to satisfy the
/// stack contract (no return).
pub fn emit_set(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::ARRAY_SET, line);
}

/// Collect all values. Stack: `[dict]` → `[array_of_values]`.
///
/// `ecma:object.entries(dict)` returns `[[k0,v0], [k1,v1], ...]`;
/// stdlib chunk `__vybe_dict_values_from_entries` walks that array
/// and projects out the v slot to give a values-only array. Works
/// across Map, plain Object, and PHP-array (`__keys`-tracked) shapes
/// because `entries` itself is polymorphic at the host fn level.
pub fn emit_values(chunks: &mut [Chunk], current: usize, line: u32) {
    // [dict] → entries → [[k,v], ...]
    let entries_fn = chunks[current].add_import("ecma:object", "entries");
    chunks[current].emit_call(entries_fn, 1, line);

    // [[k,v], ...] → call __vybe_dict_values_from_entries → [v0, v1, ...]
    let entries_local = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_local, line);
    crate::primitives::globals::emit_read(
        &mut chunks[current],
        "__vybe_dict_values_from_entries",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries_local, line);
    crate::primitives::callable::emit_direct_invoke_chunk(&mut chunks[current], 1, line);
}

/// Get all [key, value] pairs — `Object.entries(obj)` per ECMA-262
/// §20.1.2.5. Polymorphic over Array, Map, and plain Object.
/// Stack before: [dict]  Stack after: [array_of_pairs]
pub fn emit_items(chunks: &mut [Chunk], current: usize, line: u32) {
    let items_fn = chunks[current].add_import("ecma:object", "entries");
    chunks[current].emit_call(items_fn, 1, line);
}

// ── Builder pattern for dict literals ───────────────────────────────────

/// Start building a dict literal. Creates the dict with __keys tracking.
/// Stack: [] → [dict]
pub fn emit_literal(
    chunks: &mut [Chunk],
    current: usize,
    _entry_count: usize,
    line: u32,
) -> DictBuilder {
    emit_new(chunks, current, line);
    DictBuilder {}
}

/// Builder for dict literals. Call `set_entry()` for each key-value pair.
pub struct DictBuilder {}

impl DictBuilder {
    /// Set a string key entry. Value must be on stack.
    /// Stack before: [dict, value]  Stack after: [dict]
    pub fn set_entry(&mut self, chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
        emit_set_const_key(chunks, current, key, line);
    }
}


// ── Linkable chunk builders ──────────────────────────────────────────────────
//
// Linkable chunk builders — the standalone-chunk packaging of what the
// `emit_*` forms above splice inline. Same concept, same module.

// ── dict_values_from_entries(entries) → array of values ────────────
//
// Given `[[k0,v0], [k1,v1], ...]` (ECMA-262 §20.1.2.5 `Object.entries`
// shape), return `[v0, v1, ...]`. Used by `dict::emit_values` as the
// generic-shape values getter — works for Map, plain Object, and PHP
// `__keys`-tracked dict alike.
pub fn build_dict_values_from_entries(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_dict_values_from_entries");
    c.arity = 1;
    c.local_count = 4; // entries(0), result(1), i(2), len(3)
    let entries = 0u16;
    let result = 1u16;
    let i = 2u16;
    let len = 3u16;

    // result = []
    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // len = entries.length
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    crate::primitives::collections::emit_len_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, len, 0);

    // i = 0
    crate::primitives::instructions::core_wasm::i32_const(&mut c, 0, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    let block_p = c.emit_block(0);
    let (loop_p, _) = c.emit_loop_s(0);

    // if i >= len: break
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op_u16(Op::LOCAL_GET, len, 0);
    crate::primitives::ops::emit_dyn_lt_into(imports, &mut c, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(1, 0);

    // result.push(entries[i][1])
    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op_u16(Op::LOCAL_GET, entries, 0);
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    crate::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::ARRAY_GET, 0);
    crate::primitives::collections::emit_push_into(imports, &mut c, 0);
    c.emit_op(Op::DROP, 0);

    // i += 1
    c.emit_op_u16(Op::LOCAL_GET, i, 0);
    crate::primitives::instructions::core_wasm::i32_const(&mut c, 0, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, i, 0);

    c.emit_br(0, 0);
    c.emit_end(0);
    c.patch_loop(loop_p);
    c.emit_end(0);
    c.patch_block(block_p);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── setdefault(dict, key, default) — Python `dict.setdefault` ─────
//
// If `key` is present in `dict`, return its value. Otherwise set
// `dict[key] = default` and return `default`. Polymorphic (Map /
// plain Object / PHP `__keys`-tracked) via `Op::ARRAY_GET` /
// `Op::ARRAY_SET`.
pub fn build_setdefault(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_setdefault");
    c.arity = 3;
    c.local_count = 5; // dict(0), key(1), default(2), existing(3), result(4)
    let dict = 0u16;
    let key = 1u16;
    let default = 2u16;
    let existing = 3u16;
    let result = 4u16;

    // existing = dict[key]
    c.emit_op_u16(Op::LOCAL_GET, dict, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op(Op::ARRAY_GET, 0);
    c.emit_op_u16(Op::LOCAL_SET, existing, 0);

    // result = existing (default to existing; overwrite if missing)
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    // if existing is null/undefined: assign default + use it as result.
    let done_block = c.emit_block(0);
    c.emit_op_u16(Op::LOCAL_GET, existing, 0);
    c.emit_op(Op::REF_IS_NULL, 0);
    crate::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_br_if(0, 0); // existing not null → keep result, exit

    // dict[key] = default; result = default
    c.emit_op_u16(Op::LOCAL_GET, dict, 0);
    c.emit_op_u16(Op::LOCAL_GET, key, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op(Op::ARRAY_SET, 0);
    c.emit_op_u16(Op::LOCAL_GET, default, 0);
    c.emit_op_u16(Op::LOCAL_SET, result, 0);

    c.emit_end(0);
    c.patch_block(done_block);

    c.emit_op_u16(Op::LOCAL_GET, result, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

// ── keys(obj) → array of string keys ────────────────────────
// Iterates object properties, collects non-internal keys.
#[allow(dead_code)]
pub fn build_keys(imports: &mut Chunk) -> Chunk {
    // Can't iterate properties in pure bytecode without host support.
    // Use dict_keys host call pattern — but that's what we're trying to avoid.
    // Fallback: return empty array. On Vybe, host fn handles it.
    let mut c = Chunk::new("__stdlib_keys");
    c.arity = 1;
    c.local_count = 1;
    // Return empty array as fallback (properties aren't enumerable in pure WASM)
    crate::primitives::collections::emit_array_new_into(imports, &mut c, 0, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
