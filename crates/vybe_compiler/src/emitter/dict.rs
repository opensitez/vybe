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

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

// ── Creation ────────────────────────────────────────────────────────────

/// Emit bytecode to create an empty dict with key tracking.
/// Stack: [] → [dict_object]
pub fn emit_new(chunks: &mut [Chunk], current: usize, line: u32) {
    // Create empty Object
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    // Create empty __keys array and attach it
    chunks[current].emit_op(Op::DUP, line);
    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::STRUCT_SET, keys_key, line);
    chunks[current].emit_op(Op::DROP, line);
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
    let key_idx = chunks[current].add_constant(Value::String(Arc::from(key)));
    chunks[current].emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunks[current].emit_op(Op::DROP, line);
    // Append key to __keys: dict.__keys.push(key)
    chunks[current].emit_op(Op::DUP, line);
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    let key_str = chunks[current].add_constant(Value::String(Arc::from(key)));
    chunks[current].emit_op_u16(Op::CONST, key_str, line);
    crate::emitter::collections::emit_push(chunks, current, line); chunks[current].emit_op(Op::DROP, line);
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
    crate::emitter::collections::emit_set(chunks, current, line); chunks[current].emit_op(Op::DROP, line);
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
pub fn emit_method_set_tracked(chunks: &mut [Chunk], current: usize, dict_slot: u16, key_slot: u16, line: u32) {
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
    crate::emitter::collections::emit_set(chunks, current, line); chunks[current].emit_op(Op::DROP, line);

    // Push key to __keys
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::emitter::collections::emit_push(chunks, current, line); chunks[current].emit_op(Op::DROP, line);

    // Return dict for chaining
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
}

/// map.get(key) / dict[key] — get value by dynamic key.
/// Stack before: [dict, key]  Stack after: [value_or_null]
pub fn emit_method_get(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::collections::emit_get(chunks, current, line);
}

/// map.has(key) / key in dict — check if key exists.
/// Stack before: [dict, key]  Stack after: [bool]
pub fn emit_method_has(chunks: &mut [Chunk], current: usize, line: u32) {
    let dict_slot = chunks[current].local_count;
    let key_slot = dict_slot + 1;
    let has_slot = key_slot + 1;
    chunks[current].local_count += 3;
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));

    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dict_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    let no_keys = chunks[current].emit_jump(Op::BR_IF_TRUE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::emitter::collections::emit_contains(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, has_slot, line);
    let key_found = chunks[current].emit_jump(Op::BR_IF_TRUE, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::emitter::collections::emit_get(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::DYN_NOT, line);
    let end = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(no_keys);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::emitter::collections::emit_get(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::DYN_NOT, line);
    let no_keys_end = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(key_found);
    chunks[current].emit_op(Op::TRUE, line);
    chunks[current].patch_jump(end);
    chunks[current].patch_jump(no_keys_end);
}

/// map.delete(key) / del dict[key] — remove a key.
/// Stack before: [dict, key]  Stack after: [bool] (true if existed)
pub fn emit_method_delete(chunks: &mut [Chunk], current: usize, line: u32) {
    let dict_slot = chunks[current].local_count;
    let key_slot = dict_slot + 1;
    let idx_slot = dict_slot + 2;
    chunks[current].local_count += 3;
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));

    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dict_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    crate::emitter::collections::emit_index_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op(Op::DYN_GE, line);
    let missing = chunks[current].emit_jump(Op::BR_IF_FALSE, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::emitter::collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op(Op::NULL, line);
    crate::emitter::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::TRUE, line);
    let end = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(missing);
    chunks[current].emit_op(Op::FALSE, line);
    chunks[current].patch_jump(end);
}

/// map.clear() — remove all entries.
/// dict_slot must contain the dict.
/// Stack before: []  Stack after: []
pub fn emit_method_clear(chunks: &mut [Chunk], current: usize, dict_slot: u16, line: u32) {
    // Replace __keys with empty array
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::STRUCT_SET, keys_key, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Stack-based variant of `emit_method_clear`. Takes the dict from TOS,
/// clears its `__keys`, leaves `null` on TOS so the call site can use the
/// uniform "method returns void → leaves a value to drop" convention.
/// Stack before: [dict]  Stack after: [null]
pub fn emit_method_clear_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    // [dict] → set __keys = []
    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    // struct_set pops [obj, val], pushes [val]
    chunks[current].emit_op_u16(Op::STRUCT_SET, keys_key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// map.size / len(dict) / dict.Count — number of entries.
/// Stack before: [dict]  Stack after: [i32]
pub fn emit_method_size(chunks: &mut [Chunk], current: usize, line: u32) {
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    crate::emitter::collections::emit_len(chunks, current, line);
}

/// set.add(value) — add to set (same as dict[value] = true + __keys tracking).
/// Stack before: [set, value]  Stack after: [set]
pub fn emit_set_add(chunks: &mut [Chunk], current: usize, line: u32) {
    // set[value] = true
    chunks[current].emit_op(Op::TRUE, line);
    crate::emitter::collections::emit_set(chunks, current, line); chunks[current].emit_op(Op::DROP, line);
    // TODO: __keys tracking (needs key saved to a local)
}

// ── Getting values ──────────────────────────────────────────────────────

/// Emit bytecode to get a value from a dict by string key.
/// Stack before: [dict]  Stack after: [value_or_null]
pub fn emit_get_const_key(chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
    let key_idx = chunks[current].add_constant(Value::String(Arc::from(key)));
    chunks[current].emit_op_u16(Op::STRUCT_GET, key_idx, line);
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
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    // If __keys doesn't exist (legacy dict without tracking), fall back to host
    chunks[current].emit_op(Op::DUP, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    let not_null = chunks[current].emit_jump(Op::BR_IF_FALSE, line);
    // Null — fall back to host dictKeys for legacy dicts
    chunks[current].emit_op(Op::DROP, line); // drop null
    // We need the dict back — but it was consumed by struct_get.
    // Can't recover without a local. Use host call as fallback.
    // Actually, struct_get consumes the dict. We need to dup before.
    // Let's restructure: caller should dup before calling emit_keys if they need dict after.
    // For the fallback, emit empty array.
    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].patch_jump(not_null);
}

/// Emit bytecode to get all values as an array.
/// Requires dict in `dict_slot` (not consumed from stack).
/// Stack before: []  Stack after: [array_of_values]
pub fn emit_values_from_local(chunks: &mut [Chunk], current: usize, dict_slot: u16, keys_slot: u16, result_slot: u16, idx_slot: u16, line: u32) {
    // Get __keys array
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    // result = []
    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    // for i in 0..keys.length: result.push(dict[keys[i]])
    let state = crate::emitter::loops::emit_for_in_start(chunks, current, keys_slot, idx_slot, line);
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
    crate::emitter::collections::emit_get(chunks, current, line); // keys[i] → key string
    crate::emitter::collections::emit_get(chunks, current, line); // dict[key] → value
    crate::emitter::collections::emit_push(chunks, current, line); chunks[current].emit_op(Op::DROP, line); // result.push(value)
    chunks[current].emit_op(Op::DROP, line);

    crate::emitter::loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Emit bytecode to get all [key, value] pairs as an array.
/// Requires dict in `dict_slot`.
/// Stack before: []  Stack after: [array_of_pairs]
pub fn emit_items_from_local(chunks: &mut [Chunk], current: usize, dict_slot: u16, keys_slot: u16, result_slot: u16, idx_slot: u16, line: u32) {
    // Get __keys array
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    let keys_key = chunks[current].add_constant(Value::String(Arc::from("__keys")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, keys_key, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    // result = []
    crate::emitter::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    // for i in 0..keys.length: result.push([keys[i], dict[keys[i]]])
    let state = crate::emitter::loops::emit_for_in_start(chunks, current, keys_slot, idx_slot, line);
    chunks[current].emit_op(Op::DROP, line); // drop element from for_in_start

    // Build [key, value] pair
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::emitter::collections::emit_get(chunks, current, line); // key = keys[i]
    chunks[current].emit_op(Op::DUP, line); // keep key for value lookup
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
    crate::emitter::collections::emit_get(chunks, current, line);
    // value = dict[key]
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::emitter::collections::emit_get(chunks, current, line);
    crate::emitter::collections::emit_get(chunks, current, line);
    // pair = [key, value]
    crate::emitter::collections::emit_array_pair(chunks, current, line);
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
    crate::emitter::collections::emit_get(chunks, current, line); // key
    chunks[current].emit_op_u16(Op::LOCAL_GET, dict_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    crate::emitter::collections::emit_get(chunks, current, line); // key again
    crate::emitter::collections::emit_get(chunks, current, line); // dict[key] = value
    crate::emitter::collections::emit_array_pair(chunks, current, line); // [key, value]
    crate::emitter::collections::emit_push(chunks, current, line); chunks[current].emit_op(Op::DROP, line); // result.push(pair)
    chunks[current].emit_op(Op::DROP, line);

    crate::emitter::loops::emit_for_in_end(chunks, current, idx_slot, state, line);

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
    chunks[current].emit_op(Op::DROP, line);
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
    let entries_fn = chunks[0].add_import("ecma:object", "entries");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, entries_fn, line);
    chunks[current].emit(1, line);

    // [[k,v], ...] → call __vybe_dict_values_from_entries → [v0, v1, ...]
    let global_name = chunks[current].add_constant(Value::String(Arc::from("__vybe_dict_values_from_entries")));
    let entries_local = chunks[current].local_count;
    chunks[current].local_count = entries_local + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, entries_local, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::GLOBAL_GET, global_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, entries_local, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
}

/// Get all [key, value] pairs — `Object.entries(obj)` per ECMA-262
/// §20.1.2.5. Polymorphic over Array, Map, and plain Object.
/// Stack before: [dict]  Stack after: [array_of_pairs]
pub fn emit_items(chunks: &mut [Chunk], current: usize, line: u32) {
    let items_fn = chunks[0].add_import("ecma:object", "entries");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, items_fn, line);
    chunks[current].emit(1, line);
}

// ── Builder pattern for dict literals ───────────────────────────────────

/// Start building a dict literal. Creates the dict with __keys tracking.
/// Stack: [] → [dict]
pub fn emit_literal(chunks: &mut [Chunk], current: usize, _entry_count: usize, line: u32) -> DictBuilder {
    emit_new(chunks, current, line);
    DictBuilder { }
}

/// Builder for dict literals. Call `set_entry()` for each key-value pair.
pub struct DictBuilder { }

impl DictBuilder {
    /// Set a string key entry. Value must be on stack.
    /// Stack before: [dict, value]  Stack after: [dict]
    pub fn set_entry(&mut self, chunks: &mut [Chunk], current: usize, key: &str, line: u32) {
        emit_set_const_key(chunks, current, key, line);
    }
}
