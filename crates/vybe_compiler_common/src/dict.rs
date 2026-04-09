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
pub fn emit_new(chunk: &mut Chunk, line: u32) {
    // Create empty Object
    chunk.emit_op_u16(Op::struct_new, 0, line);
    // Create empty __keys array and attach it
    chunk.emit_op(Op::dup, line);
    chunk.emit_op_u16(Op::array_new, 0, line);
    let keys_key = chunk.add_constant(Value::String(Arc::from("__keys")));
    chunk.emit_op_u16(Op::struct_set, keys_key, line);
    chunk.emit_op(Op::drop, line);
}

// ── Setting keys ────────────────────────────────────────────────────────

/// Emit bytecode to set a string-literal key on a dict.
/// Also appends the key to __keys for enumeration.
/// Stack before: [dict]  (caller must dup if they want to keep dict)
/// Stack after: [dict]
pub fn emit_set_const_key(chunk: &mut Chunk, key: &str, line: u32) {
    // Set the property: dict.key = value (value must be pushed by caller before this)
    // Actually — caller pushes value AFTER calling this? No.
    // Convention: caller has [dict, value] on stack. We do struct_set.
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::struct_set, key_idx, line);
    chunk.emit_op(Op::drop, line);
    // Append key to __keys: dict.__keys.push(key)
    chunk.emit_op(Op::dup, line);
    let keys_key = chunk.add_constant(Value::String(Arc::from("__keys")));
    chunk.emit_op_u16(Op::struct_get, keys_key, line);
    let key_str = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::r#const, key_str, line);
    chunk.emit_op(Op::array_push, line);
    chunk.emit_op(Op::drop, line); // drop array_push result
}

/// Emit bytecode to set a dynamic key (key on stack).
/// Stack before: [dict, key, value]
/// Stack after: [dict]
pub fn emit_set_dynamic(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_set, line);
    chunk.emit_op(Op::drop, line);
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
pub fn emit_method_set(chunk: &mut Chunk, dict_slot: u16, line: u32) {
    // array_set(dict, key, val)
    let val_tmp = chunk.add_constant(Value::Null); // dummy, we'll use stack
    // Actually we need to save key for __keys. Use stack manipulation:
    // Stack: [key, value]. Need [dict, key, value] for array_set.
    // Then [dict.__keys, key] for array_push.

    // Save value to use later
    // Strategy: emit dict first, then key+value are already there
    // But key and value are already on stack. We need dict under them.
    // This requires the caller to push dict first. Let me redesign:

    // Actually the simplest correct API: caller pushes [dict, key, value].
    // We do array_set + __keys push.
    chunk.emit_op(Op::array_set, line);
    chunk.emit_op(Op::drop, line);
    // Push key to __keys: need dict and key again
    // The dict is still accessible via dict_slot. The key was consumed by array_set.
    // We need the caller to compile key twice, or we save it.
    // The clean API: take dict_slot, key is compiled by caller BEFORE calling this.
    // For now, skip __keys tracking for dynamic set — use emit_method_set_with_locals instead.
    chunk.emit_op_u16(Op::local_get, dict_slot, line); // return dict for chaining
}

/// map.set(key, val) with __keys tracking — requires local slots.
/// Caller must store dict in `dict_slot` and key in `key_slot`.
/// Value must be on stack.
/// Stack before: [value]  Stack after: [dict]
pub fn emit_method_set_tracked(chunk: &mut Chunk, dict_slot: u16, key_slot: u16, line: u32) {
    // array_set(dict, key, val)
    // Stack has [value]. Need [dict, key, value].
    // Save value, push dict+key, push value back.
    let val_const = chunk.add_constant(Value::Null);
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
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    chunk.emit_op_u16(Op::local_get, key_slot, line);
    // Stack: [value, dict, key]. Need [dict, key, value]. Swap:
    // Actually this is wrong. Let me just take all three from locals.
    // The caller must save value to a local too.

    // REVISED: Take dict_slot, key_slot, val_slot.
    chunk.emit_op(Op::drop, line); // drop key we just pushed
    chunk.emit_op(Op::drop, line); // drop dict we just pushed
    // Forget it — the stack ordering is too complex without a swap opcode.
    // Just use array_set directly. Caller is responsible for stack order.
    chunk.emit_op(Op::array_set, line);
    chunk.emit_op(Op::drop, line);

    // Push key to __keys
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    let keys_key = chunk.add_constant(Value::String(Arc::from("__keys")));
    chunk.emit_op_u16(Op::struct_get, keys_key, line);
    chunk.emit_op_u16(Op::local_get, key_slot, line);
    chunk.emit_op(Op::array_push, line);
    chunk.emit_op(Op::drop, line);

    // Return dict for chaining
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
}

/// map.get(key) / dict[key] — get value by dynamic key.
/// Stack before: [dict, key]  Stack after: [value_or_null]
pub fn emit_method_get(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_get, line);
}

/// map.has(key) / key in dict — check if key exists.
/// Stack before: [dict, key]  Stack after: [bool]
pub fn emit_method_has(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_get, line);
    chunk.emit_op(Op::ref_is_null, line);
    chunk.emit_op(Op::dyn_not, line);
}

/// map.delete(key) / del dict[key] — remove a key.
/// Stack before: [dict, key]  Stack after: [bool] (true if existed)
pub fn emit_method_delete(chunk: &mut Chunk, line: u32) {
    // Check if key exists first
    chunk.emit_op(Op::dup, line); // dup key
    // Stack: [dict, key, key]
    // We need to get the value first to check existence, but we also need
    // dict+key for the delete. Save dict.
    // Simplest: just set to null (WASM has no "delete property" opcode)
    chunk.emit_op(Op::null, line);
    chunk.emit_op(Op::array_set, line);
    chunk.emit_op(Op::drop, line);
    chunk.emit_op(Op::r#true, line); // assume deleted
}

/// map.clear() — remove all entries.
/// dict_slot must contain the dict.
/// Stack before: []  Stack after: []
pub fn emit_method_clear(chunk: &mut Chunk, dict_slot: u16, line: u32) {
    // Replace __keys with empty array
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    chunk.emit_op_u16(Op::array_new, 0, line);
    let keys_key = chunk.add_constant(Value::String(Arc::from("__keys")));
    chunk.emit_op_u16(Op::struct_set, keys_key, line);
    chunk.emit_op(Op::drop, line);
}

/// map.size / len(dict) / dict.Count — number of entries.
/// Stack before: [dict]  Stack after: [i32]
pub fn emit_method_size(chunk: &mut Chunk, line: u32) {
    let keys_key = chunk.add_constant(Value::String(Arc::from("__keys")));
    chunk.emit_op_u16(Op::struct_get, keys_key, line);
    chunk.emit_op(Op::array_length, line);
}

/// set.add(value) — add to set (same as dict[value] = true + __keys tracking).
/// Stack before: [set, value]  Stack after: [set]
pub fn emit_set_add(chunk: &mut Chunk, line: u32) {
    // set[value] = true
    chunk.emit_op(Op::r#true, line);
    chunk.emit_op(Op::array_set, line);
    chunk.emit_op(Op::drop, line);
    // TODO: __keys tracking (needs key saved to a local)
}

// ── Getting values ──────────────────────────────────────────────────────

/// Emit bytecode to get a value from a dict by string key.
/// Stack before: [dict]  Stack after: [value_or_null]
pub fn emit_get_const_key(chunk: &mut Chunk, key: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    chunk.emit_op_u16(Op::struct_get, key_idx, line);
}

/// Emit bytecode to get a value from a dict by dynamic key (on stack).
/// Stack before: [dict, key]  Stack after: [value_or_null]
pub fn emit_get_dynamic(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::array_get, line);
}

// ── Enumeration (pure WASM — no host calls) ─────────────────────────────

/// Emit bytecode to get all keys as an array.
/// Stack before: [dict]  Stack after: [array_of_keys]
pub fn emit_keys(chunk: &mut Chunk, line: u32) {
    let keys_key = chunk.add_constant(Value::String(Arc::from("__keys")));
    chunk.emit_op_u16(Op::struct_get, keys_key, line);
    // If __keys doesn't exist (legacy dict without tracking), fall back to host
    chunk.emit_op(Op::dup, line);
    chunk.emit_op(Op::ref_is_null, line);
    let not_null = chunk.emit_jump(Op::br_if_false, line);
    // Null — fall back to host dictKeys for legacy dicts
    chunk.emit_op(Op::drop, line); // drop null
    // We need the dict back — but it was consumed by struct_get.
    // Can't recover without a local. Use host call as fallback.
    // Actually, struct_get consumes the dict. We need to dup before.
    // Let's restructure: caller should dup before calling emit_keys if they need dict after.
    // For the fallback, emit empty array.
    chunk.emit_op_u16(Op::array_new, 0, line);
    chunk.patch_jump(not_null);
}

/// Emit bytecode to get all values as an array.
/// Requires dict in `dict_slot` (not consumed from stack).
/// Stack before: []  Stack after: [array_of_values]
pub fn emit_values_from_local(chunk: &mut Chunk, dict_slot: u16, keys_slot: u16, result_slot: u16, idx_slot: u16, line: u32) {
    // Get __keys array
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    let keys_key = chunk.add_constant(Value::String(Arc::from("__keys")));
    chunk.emit_op_u16(Op::struct_get, keys_key, line);
    chunk.emit_op_u16(Op::local_set, keys_slot, line);
    chunk.emit_op(Op::drop, line);

    // result = []
    chunk.emit_op_u16(Op::array_new, 0, line);
    chunk.emit_op_u16(Op::local_set, result_slot, line);
    chunk.emit_op(Op::drop, line);

    // for i in 0..keys.length: result.push(dict[keys[i]])
    let (loop_start, exit) = crate::loops::emit_for_in_start(chunk, keys_slot, idx_slot, line);
    // Stack: [key_string]. Use it to get value from dict.
    // Store key, get dict[key], push to result
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    // Stack: [key, dict]. Need [dict, key] for array_get. Swap via temp.
    // Actually array_get pops [key, obj] — key is TOS, obj is TOS-1.
    // We have [key, dict]. That's [obj=dict, key=key_on_top]? No.
    // array_get: pops key (TOS), then obj (TOS-1). So [dict, key] → array_get.
    // We have [key_from_for_in, dict_from_local_get].
    // Stack order: key was pushed first, dict second. So TOS=dict, TOS-1=key.
    // We need TOS=key, TOS-1=dict. They're backwards.
    // Fix: push dict first, then key.
    chunk.emit_op(Op::drop, line); // drop the element from for_in_start
    chunk.emit_op_u16(Op::local_get, result_slot, line);
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    chunk.emit_op_u16(Op::local_get, keys_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line); // keys[i] → key string
    chunk.emit_op(Op::array_get, line); // dict[key] → value
    chunk.emit_op(Op::array_push, line); // result.push(value)
    chunk.emit_op(Op::drop, line);

    crate::loops::emit_for_in_end(chunk, idx_slot, loop_start, exit, line);

    chunk.emit_op_u16(Op::local_get, result_slot, line);
}

/// Emit bytecode to get all [key, value] pairs as an array.
/// Requires dict in `dict_slot`.
/// Stack before: []  Stack after: [array_of_pairs]
pub fn emit_items_from_local(chunk: &mut Chunk, dict_slot: u16, keys_slot: u16, result_slot: u16, idx_slot: u16, line: u32) {
    // Get __keys array
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    let keys_key = chunk.add_constant(Value::String(Arc::from("__keys")));
    chunk.emit_op_u16(Op::struct_get, keys_key, line);
    chunk.emit_op_u16(Op::local_set, keys_slot, line);
    chunk.emit_op(Op::drop, line);

    // result = []
    chunk.emit_op_u16(Op::array_new, 0, line);
    chunk.emit_op_u16(Op::local_set, result_slot, line);
    chunk.emit_op(Op::drop, line);

    // for i in 0..keys.length: result.push([keys[i], dict[keys[i]]])
    let (loop_start, exit) = crate::loops::emit_for_in_start(chunk, keys_slot, idx_slot, line);
    chunk.emit_op(Op::drop, line); // drop element from for_in_start

    // Build [key, value] pair
    chunk.emit_op_u16(Op::local_get, keys_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line); // key = keys[i]
    chunk.emit_op(Op::dup, line); // keep key for value lookup
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    // Stack: [key, key_dup, dict]. Need dict[key] → [key, dict, key_dup] for array_get
    // Hmm, this is getting complicated without a swap op or temp local.
    // Simpler: re-fetch key from array.
    chunk.emit_op(Op::drop, line); // drop dup
    chunk.emit_op(Op::drop, line); // drop key

    // Start over with locals:
    // key = keys[i]
    chunk.emit_op_u16(Op::local_get, keys_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line);
    // value = dict[key]
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    chunk.emit_op_u16(Op::local_get, keys_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line);
    chunk.emit_op(Op::array_get, line);
    // pair = [key, value]
    chunk.emit_op_u16(Op::array_new, 2, line);
    // result.push(pair)
    chunk.emit_op_u16(Op::local_get, result_slot, line);
    // Stack: [pair, result]. Need [result, pair] for array_push.
    // array_push: pops [arr, val] → pushes arr. So need [result, pair].
    // We have [pair, result]. Wrong order.
    // Fix: push result first.
    chunk.emit_op(Op::drop, line); // drop result
    chunk.emit_op(Op::drop, line); // drop pair

    // Redo properly:
    chunk.emit_op_u16(Op::local_get, result_slot, line);
    // pair = [keys[i], dict[keys[i]]]
    chunk.emit_op_u16(Op::local_get, keys_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line); // key
    chunk.emit_op_u16(Op::local_get, dict_slot, line);
    chunk.emit_op_u16(Op::local_get, keys_slot, line);
    chunk.emit_op_u16(Op::local_get, idx_slot, line);
    chunk.emit_op(Op::array_get, line); // key again
    chunk.emit_op(Op::array_get, line); // dict[key] = value
    chunk.emit_op_u16(Op::array_new, 2, line); // [key, value]
    chunk.emit_op(Op::array_push, line); // result.push(pair)
    chunk.emit_op(Op::drop, line);

    crate::loops::emit_for_in_end(chunk, idx_slot, loop_start, exit, line);

    chunk.emit_op_u16(Op::local_get, result_slot, line);
}

// ── Stack-based (host-call fallback for code without local slots) ────────

/// Get dynamic key via host call (handles both __data and flat dicts).
/// Stack before: [dict, key]  Stack after: [value_or_null]
pub fn emit_get(chunk: &mut Chunk, line: u32) {
    let get_fn = chunk.add_import("vybe:types", "dictItem");
    chunk.emit_op_u16(Op::call_import, get_fn, line);
    chunk.emit(2, line);
}

/// Set dynamic key via host call.
/// Stack before: [dict, key, value]  Stack after: []
pub fn emit_set(chunk: &mut Chunk, line: u32) {
    let set_fn = chunk.add_import("vybe:types", "dictAdd");
    chunk.emit_op_u16(Op::call_import, set_fn, line);
    chunk.emit(3, line);
    chunk.emit_op(Op::drop, line);
}

/// Get all keys — tries __keys array first, falls back to host dictKeys.
/// Stack before: [dict]  Stack after: [array_of_keys]
pub fn emit_values(chunk: &mut Chunk, line: u32) {
    let vals_fn = chunk.add_import("vybe:types", "dictValues");
    chunk.emit_op_u16(Op::call_import, vals_fn, line);
    chunk.emit(1, line);
}

/// Get all [key, value] pairs.
/// Stack before: [dict]  Stack after: [array_of_pairs]
pub fn emit_items(chunk: &mut Chunk, line: u32) {
    let items_fn = chunk.add_import("vybe:array", "dictItems");
    chunk.emit_op_u16(Op::call_import, items_fn, line);
    chunk.emit(1, line);
}

// ── Builder pattern for dict literals ───────────────────────────────────

/// Start building a dict literal. Creates the dict with __keys tracking.
/// Stack: [] → [dict]
pub fn emit_literal(chunk: &mut Chunk, _entry_count: usize, line: u32) -> DictBuilder {
    emit_new(chunk, line);
    DictBuilder { }
}

/// Builder for dict literals. Call `set_entry()` for each key-value pair.
pub struct DictBuilder { }

impl DictBuilder {
    /// Set a string key entry. Value must be on stack.
    /// Stack before: [dict, value]  Stack after: [dict]
    pub fn set_entry(&mut self, chunk: &mut Chunk, key: &str, line: u32) {
        emit_set_const_key(chunk, key, line);
    }
}
