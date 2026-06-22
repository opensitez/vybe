//! PHP array helpers — Rust inline opcode emitters.
//!
//! Mirrors the inline-emit shape in `datetime_adapter.rs`: each
//! `emit_*(chunks, current, argc, line)` writes WASM opcodes directly
//! into `chunks[current]`. Composes only WASM ops + `ecma:array.*` /
//! `ecma:object.*` host imports — no PHP-specific host fns; no JS
//! polyfills. PHP `array` ≡ JS `Map` (assoc) or `Array` (sequential)
//! per the cross-language type model.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        
        _ => {
            let idx = chunk.add_constant(val);
            chunk.emit_op_u16(Op::CONST, idx, line);
        }
    }
}
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

/// Emit `wasm:js-boolean.test(val)` → i32 (1 if boolean). Value must be on stack.
pub fn emit_test_bool(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(idx, 1, line);
}
/// Emit `wasm:js-number.test(val)` → i32. Value must be on stack.
pub fn emit_test_number(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(idx, 1, line);
}
/// Emit `wasm:js-string.test(val)` → i32. Value must be on stack.
pub fn emit_test_string(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(idx, 1, line);
}
/// Emit `wasm:js-bigint.test(val)` → i32. Value must be on stack.
pub fn emit_test_bigint(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasm:js-bigint", "test");
    chunk.emit_call(idx, 1, line);
}
/// Test if value is "object-like" (not null, not a primitive).
/// Stack: [val] → i32.
pub fn emit_test_object(chunk: &mut Chunk, line: u32) {
    let slot = alloc_local(chunk);
    lset(chunk, slot, line);
    // not null AND not number AND not string AND not boolean AND not bigint AND not undefined
    lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line); // not null
    lget(chunk, slot, line);
    emit_test_number(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line); // not number
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, slot, line);
    emit_test_string(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line); // not string
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, slot, line);
    emit_test_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line); // not boolean
    chunk.emit_op(Op::I32_AND, line);
}
/// Test if value is callable (function/closure). Not null, not a primitive type.
/// For now same as object test — functions are non-primitive non-null values.
/// Stack: [val] → i32.
pub fn emit_test_function(chunk: &mut Chunk, line: u32) {
    emit_test_object(chunk, line);
}

fn emit_is_array(chunks: &mut [Chunk], current: usize, arr_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn emit_json_stringify_slots(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    flags_slot: Option<u16>,
    depth_slot: Option<u16>,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    lget(chunk, value_slot, line);
    if let Some(slot) = flags_slot {
        lget(chunk, slot, line);
    }
    if let Some(slot) = depth_slot {
        lget(chunk, slot, line);
    }
    let _ = chunk;
    call_import(chunks, current, "ecma:json", "stringify", argc, line);
}

fn emit_php_key_list_from_slot(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
}

pub fn emit_php_array_keys(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let strict_slot = alloc_local(chunk);
    let search_slot = alloc_local(chunk);
    let value_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);

    if argc >= 3 {
        lset(chunk, strict_slot, line);
    } else {
        push_const(chunk, Value::Bool(false), line);
        lset(chunk, strict_slot, line);
    }
    if argc >= 2 {
        lset(chunk, search_slot, line);
    } else {
        chunk.emit_op(Op::NULL, line);
        lset(chunk, search_slot, line);
    }
    lset(chunk, value_slot, line);

    if argc < 2 {
        emit_php_key_list_from_slot(chunks, current, value_slot, line);
        return;
    }

    emit_php_key_list_from_slot(chunks, current, value_slot, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, out_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);

    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);

    lget(chunk, value_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lget(chunk, search_slot, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);

    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

/// PHP 8.1 `array_is_list($a)` — true iff the keys are exactly 0,1,…,n-1.
/// Works on Array (keys "0".."n-1") and Map (assoc → keys won't be sequential)
/// uniformly via `ecma:object.keys`.
pub fn emit_php_array_is_list(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let keys_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);

    let _ = chunk;
    // keys = ecma:object.keys(value)  (value is on TOS)
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);
    push_const(chunk, Value::Bool(true), line);
    lset(chunk, result_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let lp = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // if (+keys[i]) !== i  OR keys[i] isn't numeric → result = false.
    // Coerce key to a number with `+0` (PHP "0"->0); a non-numeric string
    // coerces to NaN which fails the strict-position check.
    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    // Compare the string key against the string form of the index: a list
    // has keys "0","1",… == "" + 0, "" + 1, …. (Using `+ 0` here would
    // CONCATENATE the string key — "0" + 0 → "00" — and never match.)
    push_const(chunk, Value::String(std::sync::Arc::from("")), line);
    lget(chunk, i_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line); // "" + i → "0"
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    push_const(chunk, Value::Bool(false), line);
    lset(chunk, result_slot, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, lp, line);
    let chunk = &mut chunks[current];
    lget(chunk, result_slot, line);
}

pub fn emit_php_array_values(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    // PHP `array_values` returns the entry VALUES re-indexed. `ecma:object.values`
    // yields values for Array / Map (m.values()) / Object — NOT the for-of pairs
    // a Map yields under `iterForOf`. (collections::emit_iter_values is the JS
    // for-of path and must stay that way.)
    call_import(chunks, current, "ecma:object", "values", 1, line);
}

/// PHP `array_fill(start, count, value)`.
///
/// `start == 0` keeps the result on the fast sequential-array path so
/// existing array consumers like `implode` keep their usual behavior.
/// Non-zero starts use a map so PHP's numeric keys are preserved.
pub fn emit_array_fill(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    let count_slot = alloc_local(chunk);
    let start_slot = alloc_local(chunk);
    let sequential_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);

    lset(chunk, value_slot, line);
    lset(chunk, count_slot, line);
    lset(chunk, start_slot, line);

    lget(chunk, start_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    lset(chunk, sequential_slot, line);

    lget(chunk, sequential_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunk.emit_else(line);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    lset(chunk, out_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, count_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, start_slot, line);
    lget(chunk, i_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    lset(chunk, key_slot, line);

    lget(chunk, sequential_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

pub fn emit_php_end(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let arr_slot = alloc_local(chunk);
    let values_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    lset(chunk, arr_slot, line);

    lget(chunk, arr_slot, line);
    crate::emitter::collections::emit_iter_values(chunks, current, line);
    let chunk = &mut chunks[current];
    lset(chunk, values_slot, line);

    lget(chunk, values_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    lget(chunk, values_slot, line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_end(line);
}

#[allow(dead_code)]
fn emit_object_from_keys(
    chunks: &mut [Chunk],
    current: usize,
    source_slot: u16,
    keys_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let entries_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);
    let value_slot = alloc_local(chunk);
    let pair_slot = alloc_local(chunk);

    let _ = chunk;
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, entries_slot, line);

    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);

    lget(chunk, source_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, value_slot, line);

    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, pair_slot, line);
    lget(chunk, pair_slot, line);
    lget(chunk, key_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    lget(chunk, pair_slot, line);
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    lget(chunk, entries_slot, line);
    lget(chunk, pair_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);

    let chunk = &mut chunks[current];
    lget(chunk, entries_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "fromEntries", 1, line);
}

pub fn emit_php_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mode_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let value_slot = alloc_local(chunk);
    let base_len_slot = alloc_local(chunk);
    let extra_len_slot = alloc_local(chunk);

    let method_slot = alloc_local(chunk);

    if let Some(slot) = mode_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, value_slot, line);

    // PHP `count()` on a `Countable` object calls its `->count()` method.
    // Probe for a callable `count` method on the value; if present, call it.
    lget(chunk, value_slot, line);
    let count_key = chunk.add_constant(Value::String(std::sync::Arc::from("count")));
    chunk.emit_op_u16(Op::STRUCT_GET, count_key, line);
    lset(chunk, method_slot, line);
    lget(chunk, method_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    // Countable: method(value)
    lget(chunk, method_slot, line);
    lget(chunk, value_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);

    emit_is_array(chunks, current, value_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_if_value(line);

    // Sequential array → element count.
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);

    chunk.emit_else(line);
    // Associative array is an ObjectKind::Map; its size (and insertion order)
    // are native. Use the collection emitter (ecma:map / ecma:array length) —
    // no `vybe$assoc_keys_csv` side-band.
    lget(chunk, value_slot, line);
    crate::emitter::collections::emit_len(chunks, current, line);
    chunks[current].emit_end(line); // close is_array if
    chunks[current].emit_end(line); // close Countable if
    let _ = (base_len_slot, extra_len_slot);
}

pub fn emit_php_json_encode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let depth_slot = if argc >= 3 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let flags_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let value_slot = alloc_local(chunk);
    let render_slot = alloc_local(chunk);
    let assoc_keys_slot = alloc_local(chunk);
    let object_keys_slot = alloc_local(chunk);

    if let Some(slot) = depth_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = flags_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, value_slot, line);

    // Normalize the whole value tree to a JSON-serializable shape: associative
    // arrays (ObjectKind::Map) → plain Objects, recursively, so the host
    // ecma:json.stringify (which renders a bare Map as `{}` per ECMA) sees real
    // properties in native key order. Sequential arrays stay arrays. No CSV.
    super::misc_adapter::emit_php_json_normalize(chunks, current, value_slot, line);
    let chunk = &mut chunks[current];
    lset(chunk, render_slot, line);
    let _ = (assoc_keys_slot, object_keys_slot);

    emit_json_stringify_slots(
        chunks,
        current,
        render_slot,
        flags_slot,
        depth_slot,
        argc,
        line,
    );
}

/// Emit a callable-aware dispatch: call `fn_slot` as a function, or as
/// an object's `__invoke` method if the value is a class instance with
/// that magic method (PHP 8 callable-object pattern). The user-supplied
/// `push_args` closure pushes user arguments onto the stack; `argc` is
/// the count of those user args (without `$this`).
///
/// Stack on exit: `[result]` — caller `lset`s into a target slot.
fn emit_call_via_invoke_dispatch<F>(
    chunks: &mut [Chunk],
    current: usize,
    fn_slot: u16,
    argc: u8,
    line: u32,
    mut push_args: F,
) where
    F: FnMut(&mut [Chunk], usize),
{
    let chunk = &mut chunks[current];
    lget(chunk, fn_slot, line);
    emit_test_function(chunk, line);
    chunk.emit_if(line);

    let chunk = &mut chunks[current];
    lget(chunk, fn_slot, line);
    push_args(chunks, current);
    let chunk = &mut chunks[current];
    chunk.emit_op_u8(Op::CALL_REF, argc, line);
    chunk.emit_else(line);

    // Object: call $obj->__invoke(args). PHP method ABI passes `$this`
    // as arg0, so push fn (the receiver) twice.
    lget(chunk, fn_slot, line);
    let invoke_key = chunk.add_constant(Value::String(Arc::from("__invoke")));
    chunk.emit_op_u16(Op::STRUCT_GET, invoke_key, line);
    lget(chunk, fn_slot, line);
    push_args(chunks, current);
    let chunk = &mut chunks[current];
    chunk.emit_op_u8(Op::CALL_REF, argc + 1, line);
    chunk.emit_end(line);
}

pub fn emit_array_map(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let arr_slot = alloc_local(chunk);
    let fn_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let is_array_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);
    let mapped_slot = alloc_local(chunk);

    lset(chunk, arr_slot, line);
    lset(chunk, fn_slot, line);

    emit_is_array(chunks, current, arr_slot, line);
    let chunk = &mut chunks[current];
    lset(chunk, is_array_slot, line);

    lget(chunk, is_array_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunk.emit_else(line);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    lset(chunk, out_slot, line);

    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);

    emit_call_via_invoke_dispatch(chunks, current, fn_slot, 1, line, |cs, c| {
        let ch = &mut cs[c];
        lget(ch, arr_slot, line);
        lget(ch, key_slot, line);
        ch.emit_op(Op::ARRAY_GET, line);
    });
    let chunk = &mut chunks[current];
    lset(chunk, mapped_slot, line);

    lget(chunk, is_array_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, mapped_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, mapped_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

pub fn emit_array_filter(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let flag_slot = alloc_local(chunk);
    let fn_slot = alloc_local(chunk);
    let arr_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let is_array_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);
    let value_slot = alloc_local(chunk);

    if argc >= 3 {
        lset(chunk, flag_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, flag_slot, line);
    }
    if argc >= 2 {
        lset(chunk, fn_slot, line);
    } else {
        chunk.emit_op(Op::NULL, line);
        lset(chunk, fn_slot, line);
    }
    lset(chunk, arr_slot, line);

    emit_is_array(chunks, current, arr_slot, line);
    let chunk = &mut chunks[current];
    lset(chunk, is_array_slot, line);

    // PHP array_filter always preserves keys → output is always a Map
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, out_slot, line);

    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);
    lget(chunk, arr_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, value_slot, line);

    lget(chunk, fn_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_else(line);

    lget(chunk, flag_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, fn_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_else(line);

    lget(chunk, flag_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, fn_slot, line);
    lget(chunk, value_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    chunk.emit_else(line);

    lget(chunk, fn_slot, line);
    lget(chunk, value_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_end(line);

    chunk.emit_if(line);
    // Always use Map set to preserve keys
    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], out_slot, line);
}

pub fn emit_array_walk_recursive(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let userdata_slot = alloc_local(chunk);
    let fn_slot = alloc_local(chunk);
    let arr_slot = alloc_local(chunk);
    let work_slot = alloc_local(chunk);
    let cur_slot = alloc_local(chunk);
    let _ty_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);
    let child_slot = alloc_local(chunk);

    if argc >= 3 {
        lset(chunk, userdata_slot, line);
    } else {
        chunk.emit_op(Op::NULL, line);
        lset(chunk, userdata_slot, line);
    }
    lset(chunk, fn_slot, line);
    lset(chunk, arr_slot, line);

    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, work_slot, line);
    lget(chunk, work_slot, line);
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, work_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, work_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "pop", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, cur_slot, line);

    lget(chunk, cur_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_br_if(0, line);

    emit_is_array(chunks, current, cur_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_if(line);

    lget(chunk, cur_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let array_loop = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);
    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);
    lget(chunk, cur_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, child_slot, line);
    lget(chunk, work_slot, line);
    lget(chunk, child_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, array_loop, line);
    chunks[current].emit_br(1, line);
    chunks[current].emit_end(line);

    let chunk = &mut chunks[current];
    lget(chunk, cur_slot, line);
    emit_test_object(chunk, line);
    chunk.emit_if(line);

    lget(chunk, cur_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let object_loop = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);
    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);
    lget(chunk, cur_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, child_slot, line);
    lget(chunk, work_slot, line);
    lget(chunk, child_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, object_loop, line);
    chunks[current].emit_br(1, line);
    chunks[current].emit_end(line);

    let callback_arity = if argc >= 3 { 2 } else { 1 };
    emit_call_via_invoke_dispatch(chunks, current, fn_slot, callback_arity, line, |cs, c| {
        let ch = &mut cs[c];
        lget(ch, cur_slot, line);
        if argc >= 3 {
            lget(ch, userdata_slot, line);
        }
    });
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];
    push_const(chunk, Value::Bool(true), line);
}

// ── array_pad ──────────────────────────────────────────────────────

/// PHP `array_pad(arr, size, value)`. abs(size) target length;
/// negative pads left.
pub fn emit_array_pad(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    let size_slot = alloc_local(chunk);
    let arr_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let target_slot = alloc_local(chunk);
    let diff_slot = alloc_local(chunk);
    let pad_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);

    lset(chunk, value_slot, line);
    lset(chunk, size_slot, line);
    lset(chunk, arr_slot, line);

    // len = arr.length
    lget(chunk, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    // target = abs(size)
    lget(chunk, size_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, size_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_else(line);
    lget(chunk, size_slot, line);
    chunk.emit_end(line);
    lset(chunk, target_slot, line);

    // if target <= len: return arr.slice() (just a clone)
    lget(chunk, target_slot, line);
    lget(chunk, len_slot, line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if_value(line);
    // diff = target - len
    lget(chunk, target_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, diff_slot, line);

    // pad = []
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, pad_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    // for i in 0..diff: pad.push(value)
    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, diff_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, pad_slot, line);
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    // result = size < 0 ? pad.concat(arr) : arr.concat(pad)
    lget(chunk, size_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if_value(line);
    // Pad-left: pad.concat(arr)
    lget(chunk, pad_slot, line);
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "concat", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    // Pad-right: arr.concat(pad)
    lget(chunk, arr_slot, line);
    lget(chunk, pad_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "concat", 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    // No pad: clone via ecma:array.slice(arr)
    lget(&mut chunks[current], arr_slot, line);
    call_import(chunks, current, "ecma:array", "slice", 1, line);
    chunks[current].emit_end(line);
}

// ── array_chunk ────────────────────────────────────────────────────

/// PHP `array_chunk(arr, size, preserve_keys?)` → array of chunks.
///
/// Iterates the input via `Object.keys` so it works for both Map-backed
/// PHP assoc arrays and sequential Arrays. When `preserve_keys` is true,
/// each chunk is a Map carrying the original keys; otherwise each chunk
/// is a sequential Array.
pub fn emit_array_chunk(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let preserve_slot = alloc_local(chunk);
    let size_slot = alloc_local(chunk);
    let arr_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let end_slot = alloc_local(chunk);
    let chunk_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);

    if argc >= 3 {
        crate::emitter::ops::emit_dyn_to_bool(chunk, line);
        lset(chunk, preserve_slot, line);
    } else {
        push_const(chunk, Value::Bool(false), line);
        lset(chunk, preserve_slot, line);
    }
    lset(chunk, size_slot, line);
    lset(chunk, arr_slot, line);

    // out = []
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, out_slot, line);

    // block $done { if size < 1 { br $done } ... }
    let done_block = chunk.emit_block(line);
    lget(chunk, size_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);

    // keys = Object.keys(arr)
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    // n = keys.length; i = 0
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    // Outer loop: walk keys in `size` strides.
    let outer_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // end = min(i + size, n)
    lget(chunk, i_slot, line);
    lget(chunk, size_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, end_slot, line);
    lget(chunk, end_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, n_slot, line);
    lset(chunk, end_slot, line);
    chunk.emit_end(line);

    // chunk_obj = preserve ? ecma:map.new() : []
    lget(chunk, preserve_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunk.emit_end(line);
    lset(chunk, chunk_slot, line);

    // j = i
    lget(chunk, i_slot, line);
    lset(chunk, j_slot, line);

    // Inner loop: for j in i..end
    let _ = chunk;
    let inner_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, end_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // key = keys[j]
    lget(chunk, keys_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);

    // if preserve: chunk_obj[key] = arr[key] ; else chunk_obj.push(arr[key])
    lget(chunk, preserve_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, chunk_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, arr_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_else(line);
    lget(chunk, chunk_slot, line);
    lget(chunk, arr_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);

    // j++
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, inner_state, line);
    let chunk = &mut chunks[current];

    // out.push(chunk_obj)
    lget(chunk, out_slot, line);
    lget(chunk, chunk_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];

    // i = end
    lget(chunk, end_slot, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, outer_state, line);
    let chunk = &mut chunks[current];

    chunk.emit_end(line);
    chunk.patch_block(done_block);
    lget(chunk, out_slot, line);
}

// ── array_combine ──────────────────────────────────────────────────

/// PHP `array_combine(keys, values)` — zip into Object (assoc array).
pub fn emit_array_combine(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let values_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);

    lset(chunk, values_slot, line);
    lset(chunk, keys_slot, line);

    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, out_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // out[keys[i]] = values[i]
    lget(chunk, out_slot, line);
    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lget(chunk, values_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

/// PHP `array_fill_keys(keys, value)` — build an associative array
/// whose keys come from `keys` and all map to the same `value`.
pub fn emit_array_fill_keys(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);

    lset(chunk, value_slot, line);
    lset(chunk, keys_slot, line);

    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, out_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

// ── array_flip ─────────────────────────────────────────────────────

/// PHP `array_flip(obj)` — swap keys and values.
/// Stack: `[obj]` → `[Object<value→key>]`.
pub fn emit_array_flip(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let arr_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let k_slot = alloc_local(chunk);

    lset(chunk, arr_slot, line);

    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, out_slot, line);

    // keys = Object.keys(arr)
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // k = keys[i]; out[arr[k]] = k
    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, k_slot, line);

    lget(chunk, out_slot, line);
    lget(chunk, arr_slot, line);
    lget(chunk, k_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lget(chunk, k_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], out_slot, line);
}

// ── array_diff / array_intersect (value-only, sequential arrays) ──

fn emit_array_diff_or_intersect(chunks: &mut [Chunk], current: usize, intersect: bool, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let a_slot = alloc_local(chunk);
    let seen_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);
    let blen_slot = alloc_local(chunk);
    let alen_slot = alloc_local(chunk);
    let v_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);
    let has_slot = alloc_local(chunk);

    lset(chunk, b_slot, line);
    lset(chunk, a_slot, line);

    // seen = Object.new()
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, seen_slot, line);

    // out = []
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, out_slot, line);

    // for i in 0..b.length: seen[String(b[i])] = true
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, blen_slot, line);

    let _ = chunk;
    let loop1_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, blen_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, b_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    push_str(chunk, "", line);
    let _ = chunk;
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    let chunk = &mut chunks[current];
    lset(chunk, key_slot, line);
    lget(chunk, seen_slot, line);
    lget(chunk, key_slot, line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_op(Op::ARRAY_SET, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop1_state, line);
    let chunk = &mut chunks[current];

    // for j in 0..a.length: if (seen[String(a[j])] == intersect): out.push(a[j])
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, j_slot, line);
    lget(chunk, a_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, alen_slot, line);

    let _ = chunk;
    let loop2_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j_slot, line);
    lget(chunk, alen_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // v = a[j]; key = "" + v; has = seen[key]
    lget(chunk, a_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, v_slot, line);
    push_str(chunk, "", line);
    lget(chunk, v_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    lset(chunk, key_slot, line);
    lget(chunk, seen_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, has_slot, line);

    // if intersect ? has : !has → push v
    lget(chunk, has_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    if !intersect {
        crate::emitter::ops::emit_dyn_not(chunk, line);
    }
    chunk.emit_if(line);
    lget(chunk, out_slot, line);
    lget(chunk, v_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop2_state, line);
    let chunk = &mut chunks[current];

    lget(chunk, out_slot, line);
}

pub fn emit_array_diff(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_array_diff_or_intersect(chunks, current, /*intersect=*/ false, line);
}
pub fn emit_array_intersect(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_array_diff_or_intersect(chunks, current, /*intersect=*/ true, line);
}

// ── array_count_values ─────────────────────────────────────────────

pub fn emit_array_count_values(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let arr_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let key_slot = alloc_local(chunk);
    let cur_slot = alloc_local(chunk);

    lset(chunk, arr_slot, line);

    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    lset(chunk, out_slot, line);

    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    lget(chunk, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    let _ = chunk;
    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // key = "" + arr[i]
    push_str(chunk, "", line);
    lget(chunk, arr_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    lset(chunk, key_slot, line);

    // cur = (out[key] || 0) + 1
    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, cur_slot, line);
    // if cur is null/undefined: cur = 0
    lget(chunk, cur_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, cur_slot, line);
    chunk.emit_end(line);

    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, cur_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], out_slot, line);
}

// ── array_column ───────────────────────────────────────────────────

pub fn emit_array_column(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let has_index = argc >= 3;
    // Allocate slots first.
    let (index_key_slot, col_slot, rows_slot, out_slot, i_slot, len_slot, row_slot, value_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        if has_index {
            lset(chunk, index_key_slot, line);
        }
        lset(chunk, col_slot, line);
        lset(chunk, rows_slot, line);
    }

    // out = has_index ? Object.new() : []
    if has_index {
        call_import(chunks, current, "ecma:map", "new", 0, line);
    } else {
        chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    }
    {
        let chunk = &mut chunks[current];
        lset(chunk, out_slot, line);
        // i = 0; len = rows.length
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, i_slot, line);
        lget(chunk, rows_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, len_slot, line);
    }

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];

        lget(chunk, rows_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, row_slot, line);

        lget(chunk, col_slot, line);
        chunk.emit_op(Op::NULL, line);
        crate::emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);

        lget(chunk, row_slot, line);
        lset(chunk, value_slot, line);
        chunk.emit_else(line);
        lget(chunk, row_slot, line);
        lget(chunk, col_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, value_slot, line);
        chunk.emit_end(line);

        if has_index {
            lget(chunk, out_slot, line);
            lget(chunk, row_slot, line);
            lget(chunk, index_key_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            lget(chunk, value_slot, line);
            chunk.emit_op(Op::ARRAY_SET, line);
        }
    }
    if !has_index {
        {
            let chunk = &mut chunks[current];
            lget(chunk, out_slot, line);
            lget(chunk, value_slot, line);
        }
        call_import(chunks, current, "ecma:array", "push", 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

// ── array_key_first / array_key_last ───────────────────────────────

fn emit_array_key_first_or_last(chunks: &mut [Chunk], current: usize, last: bool, line: u32) {
    let chunk = &mut chunks[current];
    let arr_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);

    lset(chunk, arr_slot, line);

    emit_php_key_list_from_slot(chunks, current, arr_slot, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_else(line);
    if last {
        lget(chunk, keys_slot, line);
        lget(chunk, len_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_SUB, line);
        chunk.emit_op(Op::ARRAY_GET, line);
    } else {
        lget(chunk, keys_slot, line);
        push_const(chunk, Value::F64(0.0), line);
        chunk.emit_op(Op::ARRAY_GET, line);
    }
    chunk.emit_end(line);
}

pub fn emit_array_key_first(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_array_key_first_or_last(chunks, current, /*last=*/ false, line);
}
pub fn emit_array_key_last(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_array_key_first_or_last(chunks, current, /*last=*/ true, line);
}

// ── array_diff_key / array_diff_assoc / array_intersect_key / array_replace ─────────

/// PHP `array_diff_key(a, b)` — entries in a whose keys do not exist in b.
pub fn emit_array_diff_key(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (b_slot, a_slot, out_slot, keys_slot, i_slot, len_slot, k_slot, av_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, b_slot, line);
        lset(chunk, a_slot, line);
    }
    call_import(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, i_slot, line);
        lget(chunk, keys_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, len_slot, line);
    }

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, k_slot, line);

        lget(chunk, a_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, av_slot, line);

        lget(chunk, b_slot, line);
        lget(chunk, k_slot, line);
    }
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    {
        let chunk = &mut chunks[current];
        crate::emitter::ops::emit_dyn_not(chunk, line);
        chunk.emit_if(line);
        lget(chunk, out_slot, line);
        lget(chunk, k_slot, line);
        lget(chunk, av_slot, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);

        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    lget(&mut chunks[current], out_slot, line);
}

/// PHP `array_diff_assoc(a, b)` — entries in a whose key→value pair
/// differs in b.
pub fn emit_array_diff_assoc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (b_slot, a_slot, out_slot, keys_slot, i_slot, len_slot, k_slot, av_slot, bv_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, b_slot, line);
        lset(chunk, a_slot, line);
    }
    call_import(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, i_slot, line);
        lget(chunk, keys_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, len_slot, line);
    }

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, k_slot, line);

        lget(chunk, a_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, av_slot, line);

        lget(chunk, b_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, bv_slot, line);

        lget(chunk, bv_slot, line);
        crate::emitter::convert::emit_to_string(chunk, line);
        lget(chunk, av_slot, line);
        crate::emitter::convert::emit_to_string(chunk, line);
        crate::emitter::ops::emit_dyn_eq(chunk, line);
        crate::emitter::ops::emit_dyn_not(chunk, line);
        chunk.emit_if(line);
        lget(chunk, out_slot, line);
        lget(chunk, k_slot, line);
        lget(chunk, av_slot, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);

        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    lget(&mut chunks[current], out_slot, line);
}

/// PHP `array_intersect_assoc(a, b)` — entries in a whose key→value pair
/// matches the corresponding pair in b.
pub fn emit_array_intersect_assoc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (b_slot, a_slot, out_slot, keys_slot, i_slot, len_slot, k_slot, av_slot, bv_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, b_slot, line);
        lset(chunk, a_slot, line);
    }
    call_import(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, i_slot, line);
        lget(chunk, keys_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, len_slot, line);
    }

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, k_slot, line);

        lget(chunk, a_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, av_slot, line);

        lget(chunk, b_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, bv_slot, line);

        lget(chunk, bv_slot, line);
        crate::emitter::convert::emit_to_string(chunk, line);
        lget(chunk, av_slot, line);
        crate::emitter::convert::emit_to_string(chunk, line);
        crate::emitter::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        lget(chunk, out_slot, line);
        lget(chunk, k_slot, line);
        lget(chunk, av_slot, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);

        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    lget(&mut chunks[current], out_slot, line);
}

/// PHP `array_intersect_key(a, b)` — entries from a whose keys exist in b.
pub fn emit_array_intersect_key(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (b_slot, a_slot, out_slot, keys_slot, i_slot, len_slot, k_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, b_slot, line);
        lset(chunk, a_slot, line);
    }
    call_import(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, i_slot, line);
        lget(chunk, keys_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, len_slot, line);
    }

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, k_slot, line);

        lget(chunk, b_slot, line);
        lget(chunk, k_slot, line);
    }
    call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
    {
        let chunk = &mut chunks[current];
        chunk.emit_if(line);
        lget(chunk, out_slot, line);
        lget(chunk, k_slot, line);
        lget(chunk, a_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);

        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    lget(&mut chunks[current], out_slot, line);
}

/// PHP `array_replace(a, b)` — a + b, b's keys override a's.
pub fn emit_array_replace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (b_slot, a_slot, out_slot, keys_slot, i_slot, len_slot, k_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, b_slot, line);
        lset(chunk, a_slot, line);
    }
    call_import(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    // Copy a then b. Two passes for simplicity.
    for src_slot in &[a_slot, b_slot] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, *src_slot, line);
        call_import(chunks, current, "ecma:object", "keys", 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
        chunks[current].emit_op(Op::DROP, line);

        {
            let chunk = &mut chunks[current];
            push_const(chunk, Value::F64(0.0), line);
            lset(chunk, i_slot, line);
            lget(chunk, keys_slot, line);
            chunk.emit_op(Op::ARRAY_LENGTH, line);
            lset(chunk, len_slot, line);
        }
        let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
        {
            let chunk = &mut chunks[current];
            lget(chunk, i_slot, line);
            lget(chunk, len_slot, line);
            crate::emitter::ops::emit_dyn_lt(chunk, line);
        }
        crate::emitter::loops::emit_loop_cond(chunks, current, line);
        {
            let chunk = &mut chunks[current];

            lget(chunk, keys_slot, line);
            lget(chunk, i_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            lset(chunk, k_slot, line);

            lget(chunk, out_slot, line);
            lget(chunk, k_slot, line);
            chunk.emit_op_u16(Op::LOCAL_GET, *src_slot, line);
            lget(chunk, k_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op(Op::ARRAY_SET, line);
            chunk.emit_op(Op::DROP, line);

            lget(chunk, i_slot, line);
            push_const(chunk, Value::F64(1.0), line);
            chunk.emit_op(Op::F64_ADD, line);
            lset(chunk, i_slot, line);
        }
        crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn emit_copy_object_entries(
    chunks: &mut [Chunk],
    current: usize,
    src_slot: u16,
    dst_slot: u16,
    line: u32,
) {
    let keys_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let i_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let len_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };
    let key_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };

    emit_php_key_list_from_slot(chunks, current, src_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, i_slot, line);
        lget(chunk, keys_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, len_slot, line);
    }

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];

        lget(chunk, keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, key_slot, line);

        lget(chunk, dst_slot, line);
        lget(chunk, key_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, src_slot, line);
        {
            let chunk = &mut chunks[current];
            lget(chunk, key_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op(Op::ARRAY_SET, line);
            chunk.emit_op(Op::DROP, line);

            lget(chunk, i_slot, line);
            push_const(chunk, Value::F64(1.0), line);
            chunk.emit_op(Op::F64_ADD, line);
            lset(chunk, i_slot, line);
        }
    }
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
}

fn emit_generator_yield_value_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    yielded_slot: u16,
    line: u32,
) {
    let payload_id_slot = {
        let chunk = &mut chunks[current];
        alloc_local(chunk)
    };

    let chunk = &mut chunks[current];
    lget(chunk, yielded_slot, line);
    emit_test_object(chunk, line);
    chunk.emit_if(line);

    lget(chunk, yielded_slot, line);
    let marker_key = chunk.add_constant(Value::String(Arc::from("__vybe_generator_yield")));
    chunk.emit_op_u16(Op::STRUCT_GET, marker_key, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    lget(chunk, yielded_slot, line);
    let payload_id_key = chunk.add_constant(Value::String(Arc::from("payload_id")));
    chunk.emit_op_u16(Op::STRUCT_GET, payload_id_key, line);
    lset(chunk, payload_id_slot, line);

    lget(chunk, payload_id_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);

    lget(chunk, yielded_slot, line);
    let value_key = chunk.add_constant(Value::String(Arc::from("value")));
    chunk.emit_op_u16(Op::STRUCT_GET, value_key, line);

    chunk.emit_else(line);

    let store_key = chunk.add_constant(Value::String(Arc::from("__vybe_generator_payloads")));
    chunk.emit_op_u16(Op::GLOBAL_GET, store_key, line);
    lget(chunk, payload_id_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);

    chunk.emit_end(line);
    chunk.emit_else(line);
    lget(chunk, yielded_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    lget(chunk, yielded_slot, line);
    chunk.emit_end(line);
}

fn emit_generator_yield_key_or_fallback_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    yielded_slot: u16,
    fallback_slot: Option<u16>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    lget(chunk, yielded_slot, line);
    emit_test_object(chunk, line);
    chunk.emit_if(line);

    lget(chunk, yielded_slot, line);
    let marker_key = chunk.add_constant(Value::String(Arc::from("__vybe_generator_yield")));
    chunk.emit_op_u16(Op::STRUCT_GET, marker_key, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    lget(chunk, yielded_slot, line);
    let key_key = chunk.add_constant(Value::String(Arc::from("key")));
    chunk.emit_op_u16(Op::STRUCT_GET, key_key, line);

    chunk.emit_else(line);
    if let Some(slot) = fallback_slot {
        lget(chunk, slot, line);
    } else {
        chunk.emit_op(Op::NULL, line);
    }
    chunk.emit_end(line);
    chunk.emit_else(line);
    if let Some(slot) = fallback_slot {
        lget(chunk, slot, line);
    } else {
        chunk.emit_op(Op::NULL, line);
    }
    chunk.emit_end(line);
}

pub fn emit_iterator_to_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (
        preserve_keys_slot,
        iter_slot,
        out_slot,
        value_slot,
        has_more_slot,
        index_slot,
        keys_slot,
        i_slot,
        len_slot,
        key_slot,
    ) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };

    if argc >= 2 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, preserve_keys_slot, line);
        chunks[current].emit_op(Op::DROP, line);
    } else {
        push_const(&mut chunks[current], Value::Bool(true), line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, preserve_keys_slot, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, iter_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    {
        let chunk = &mut chunks[current];
        lget(chunk, preserve_keys_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        let _ = chunk;
        call_import(chunks, current, "ecma:map", "new", 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_end(line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, iter_slot, line);
    call_import(chunks, current, "ecma:value", "isGenerator", 1, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, index_slot, line);
    }

    let gen_loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, iter_slot, line);
        crate::emitter::generators::emit_next(chunk, line);
        lset(chunk, has_more_slot, line);
        lset(chunk, value_slot, line);

        lget(chunk, has_more_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];

        lget(chunk, preserve_keys_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);

        lget(chunk, out_slot, line);
        let _ = chunk;
        emit_generator_yield_key_or_fallback_from_slot(
            chunks,
            current,
            value_slot,
            Some(index_slot),
            line,
        );
        emit_generator_yield_value_from_slot(chunks, current, value_slot, line);
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_op(Op::DROP, line);

        chunk.emit_else(line);
        lget(chunk, out_slot, line);
        let _ = chunk;
        emit_generator_yield_value_from_slot(chunks, current, value_slot, line);
        call_import(chunks, current, "ecma:array", "push", 2, line);
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);

        lget(chunk, index_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, index_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, gen_loop_state, line);
    // After the generator is exhausted, store the return value for getReturn().
    // The last emit_next returned (false, return_value); value_slot has it.
    {
        let chunk = &mut chunks[current];
        lget(chunk, iter_slot, line);
        lget(chunk, value_slot, line);
        let ret_k = chunk.add_constant(Value::String(std::sync::Arc::from("__php_gen_return")));
        chunk.emit_op_u16(Op::STRUCT_SET, ret_k, line);
        chunk.emit_op(Op::DROP, line);
    }
    lget(&mut chunks[current], out_slot, line);

    chunks[current].emit_else(line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, preserve_keys_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);

        let _ = chunk;
        emit_copy_object_entries(chunks, current, iter_slot, out_slot, line);
        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);

        chunk.emit_else(line);
    }
    emit_php_key_list_from_slot(chunks, current, iter_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, i_slot, line);
        lget(chunk, keys_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, len_slot, line);
    }

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, key_slot, line);
        lget(chunk, out_slot, line);
        lget(chunk, iter_slot, line);
        lget(chunk, key_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        let _ = chunk;
        call_import(chunks, current, "ecma:array", "push", 2, line);
        let chunk = &mut chunks[current];
        chunk.emit_op(Op::DROP, line);
        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    lget(&mut chunks[current], out_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// PHP `array_replace_recursive(a, b)` — recursive key replacement for
/// nested associative arrays. This adapter handles the common object/map
/// shape directly in emitted ops.
pub fn emit_array_replace_recursive(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (
        over_slot,
        base_slot,
        out_slot,
        keys_slot,
        i_slot,
        len_slot,
        key_slot,
        over_val_slot,
        cur_val_slot,
        merged_slot,
        cur_keys_slot,
        over_keys_slot,
        should_merge_slot,
    ) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        lset(chunk, over_slot, line);
        lset(chunk, base_slot, line);
    }

    call_import(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_copy_object_entries(chunks, current, base_slot, out_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, over_slot, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    {
        let chunk = &mut chunks[current];
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, i_slot, line);
        lget(chunk, keys_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        lset(chunk, len_slot, line);
    }

    let loop_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let chunk = &mut chunks[current];

        lget(chunk, keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, key_slot, line);

        lget(chunk, over_slot, line);
        lget(chunk, key_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, over_val_slot, line);

        lget(chunk, out_slot, line);
        lget(chunk, key_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, cur_val_slot, line);

        push_const(chunk, Value::Bool(false), line);
        lset(chunk, should_merge_slot, line);

        lget(chunk, cur_val_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        crate::emitter::ops::emit_dyn_not(chunk, line);
        chunk.emit_if(line);

        let _ = chunk;
        emit_php_key_list_from_slot(chunks, current, cur_val_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, cur_keys_slot, line);
        chunks[current].emit_op(Op::DROP, line);
        let chunk = &mut chunks[current];
        lget(chunk, cur_keys_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        chunk.emit_op(Op::I32_CONST_0, line);
        crate::emitter::ops::emit_dyn_gt(chunk, line);
        chunk.emit_if(line);

        lget(chunk, over_val_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        crate::emitter::ops::emit_dyn_not(chunk, line);
        chunk.emit_if(line);

        let _ = chunk;
        emit_php_key_list_from_slot(chunks, current, over_val_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, over_keys_slot, line);
        chunks[current].emit_op(Op::DROP, line);
        let chunk = &mut chunks[current];
        lget(chunk, over_keys_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
        chunk.emit_op(Op::I32_CONST_0, line);
        crate::emitter::ops::emit_dyn_gt(chunk, line);
        chunk.emit_if(line);

        push_const(chunk, Value::Bool(true), line);
        lset(chunk, should_merge_slot, line);

        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunk.emit_end(line);

        lget(chunk, should_merge_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);

        let _ = chunk;
        call_import(chunks, current, "ecma:map", "new", 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, merged_slot, line);
        chunks[current].emit_op(Op::DROP, line);
        emit_copy_object_entries(chunks, current, cur_val_slot, merged_slot, line);
        emit_copy_object_entries(chunks, current, over_val_slot, merged_slot, line);

        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        lget(chunk, key_slot, line);
        lget(chunk, merged_slot, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_op(Op::DROP, line);

        chunk.emit_else(line);
        lget(chunk, out_slot, line);
        lget(chunk, key_slot, line);
        lget(chunk, over_val_slot, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);

        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, loop_state, line);
    lget(&mut chunks[current], out_slot, line);
}

// ── asort / arsort / ksort / krsort / uasort / uksort ───────────────
//
// Selection-sort building a NEW sorted-keys Map (avoids in-place
// `ecma:array.set` mutation which doesn't round-trip through
// `ecma:object.keys`-returned arrays in some cases). For each round,
// scan the unused entries and pick the "best" by `mode` comparison,
// mark it used, and append to the sorted result. After sorting,
// delete every original key from `obj` and re-insert in sorted order.
//
// `mode` selects the comparison:
//   0 = asc-by-value, 1 = desc-by-value,
//   2 = asc-by-key,   3 = desc-by-key,
//   4 = user(value),  5 = user(key)
// `cmp_slot` holds the user callback for modes 4/5; ignored otherwise.
fn emit_assoc_sort_impl(
    chunks: &mut [Chunk],
    current: usize,
    mode: u8,
    cmp_slot: Option<u16>,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let obj_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let used_slot = alloc_local(chunk); // Map<index, true>
    let sorted_keys_slot = alloc_local(chunk); // Array of sorted keys
    let sorted_vals_slot = alloc_local(chunk); // Array of sorted values
    let outer_slot = alloc_local(chunk);
    let inner_slot = alloc_local(chunk);
    let best_slot = alloc_local(chunk);
    let is_list_slot = alloc_local(chunk);

    // obj = pop()
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op(Op::DROP, line);

    // is_list = isArray(obj) — a sequential (list) array reorders by
    // POSITION, not by Map insertion order; the delete-then-reinsert dance
    // below only reorders a Map, so lists must be written back positionally.
    emit_is_array(chunks, current, obj_slot, line);
    let chunk = &mut chunks[current];
    lset(chunk, is_list_slot, line);

    // keys = Object.keys(obj)
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunk.emit_op(Op::DROP, line);

    // n = keys.length
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);

    // used = ecma:map.new(); sorted_keys = []; sorted_vals = []
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, used_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, sorted_keys_slot, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, sorted_vals_slot, line);

    // outer loop: for outer in 0..n
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, outer_slot, line);

    let _ = chunk;
    let outer_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, outer_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // best = -1
    push_const(chunk, Value::F64(-1.0), line);
    lset(chunk, best_slot, line);

    // inner loop: for inner in 0..n
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, inner_slot, line);

    let _ = chunk;
    let inner_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, inner_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    // if used[inner]: skip
    lget(chunk, used_slot, line);
    lget(chunk, inner_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    crate::emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_if(line);

    // if best === -1: best = inner ; else compare
    lget(chunk, best_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    lget(chunk, inner_slot, line);
    lset(chunk, best_slot, line);
    chunk.emit_else(line);

    // Compare: should `inner` replace `best`?
    // For asort (mode=0): obj[keys[inner]] < obj[keys[best]] → replace
    // For arsort (1): obj[keys[inner]] > obj[keys[best]] → replace
    // For ksort (2): keys[inner] < keys[best] → replace
    // For krsort (3): keys[inner] > keys[best] → replace
    // For uasort (4): cmp(obj[keys[inner]], obj[keys[best]]) < 0 → replace
    // For uksort (5): cmp(keys[inner], keys[best]) < 0 → replace
    match mode {
        0 | 1 => {
            // numeric value comparison
            lget(chunk, obj_slot, line);
            lget(chunk, keys_slot, line);
            lget(chunk, inner_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            push_const(chunk, Value::F64(0.0), line);
            crate::emitter::ops::emit_dyn_add(chunk, line);
            lget(chunk, obj_slot, line);
            lget(chunk, keys_slot, line);
            lget(chunk, best_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            push_const(chunk, Value::F64(0.0), line);
            crate::emitter::ops::emit_dyn_add(chunk, line);
            if mode == 0 {
                crate::emitter::ops::emit_dyn_lt(chunk, line);
            } else {
                crate::emitter::ops::emit_dyn_gt(chunk, line);
            }
        }
        2 | 3 => {
            // string key comparison
            lget(chunk, keys_slot, line);
            lget(chunk, inner_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            lget(chunk, keys_slot, line);
            lget(chunk, best_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            { let idx = chunk.add_import("wasm:js-string", "compare"); chunk.emit_call(idx, 2, line); }
            push_const(chunk, Value::F64(0.0), line);
            if mode == 2 {
                crate::emitter::ops::emit_dyn_lt(chunk, line);
            } else {
                crate::emitter::ops::emit_dyn_gt(chunk, line);
            }
        }
        4 => {
            // user(value): cmp(obj[keys[inner]], obj[keys[best]]) < 0
            let cs = cmp_slot.expect("uasort needs cmp_slot");
            lget(chunk, cs, line);
            lget(chunk, obj_slot, line);
            lget(chunk, keys_slot, line);
            lget(chunk, inner_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            lget(chunk, obj_slot, line);
            lget(chunk, keys_slot, line);
            lget(chunk, best_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op_u8(Op::CALL_REF, 2, line);
            push_const(chunk, Value::F64(0.0), line);
            crate::emitter::ops::emit_dyn_lt(chunk, line);
        }
        5 => {
            // user(key): cmp(keys[inner], keys[best]) < 0
            let cs = cmp_slot.expect("uksort needs cmp_slot");
            lget(chunk, cs, line);
            lget(chunk, keys_slot, line);
            lget(chunk, inner_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            lget(chunk, keys_slot, line);
            lget(chunk, best_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op_u8(Op::CALL_REF, 2, line);
            push_const(chunk, Value::F64(0.0), line);
            crate::emitter::ops::emit_dyn_lt(chunk, line);
        }
        _ => {
            push_const(chunk, Value::Bool(false), line);
        }
    }
    chunk.emit_if(line);
    lget(chunk, inner_slot, line);
    lset(chunk, best_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // inner++
    lget(chunk, inner_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, inner_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, inner_state, line);
    let chunk = &mut chunks[current];

    // used[best] = true
    lget(chunk, used_slot, line);
    lget(chunk, best_slot, line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);

    // sorted_keys.push(keys[best])
    lget(chunk, sorted_keys_slot, line);
    lget(chunk, keys_slot, line);
    lget(chunk, best_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    // sorted_vals.push(obj[keys[best]])
    let chunk = &mut chunks[current];
    lget(chunk, sorted_vals_slot, line);
    lget(chunk, obj_slot, line);
    lget(chunk, keys_slot, line);
    lget(chunk, best_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    // outer++
    let chunk = &mut chunks[current];
    lget(chunk, outer_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, outer_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, outer_state, line);

    // Delete every original key from obj.
    let chunk = &mut chunks[current];
    let i_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    let del_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, obj_slot, line);
    lget(chunk, keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let _ = chunk;
    // PHP `array` is a Map in Vybe — must use `ecma:map.delete` (the
    // `ecma:object.delete` path only removes from `properties`,
    // bypassing the IndexMap backing).
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, del_state, line);
    let chunk = &mut chunks[current];

    // Re-insert in sorted order: obj[sorted_keys[i]] = sorted_vals[i].
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    let ins_state = crate::emitter::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, obj_slot, line);
    // key = is_list ? i : sorted_keys[i]
    lget(chunk, is_list_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, i_slot, line);
    chunk.emit_else(line);
    lget(chunk, sorted_keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_end(line);
    lget(chunk, sorted_vals_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    let _ = chunk;
    crate::emitter::loops::emit_loop_end(chunks, current, ins_state, line);

    // PHP sort family returns true.
    push_const(&mut chunks[current], Value::Bool(true), line);
}

pub fn emit_php_asort(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_assoc_sort_impl(chunks, current, 0, None, line);
}
pub fn emit_php_arsort(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_assoc_sort_impl(chunks, current, 1, None, line);
}
pub fn emit_php_ksort(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_assoc_sort_impl(chunks, current, 2, None, line);
}
pub fn emit_php_krsort(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_assoc_sort_impl(chunks, current, 3, None, line);
}

pub fn emit_php_uasort(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let cmp_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, s, line);
        chunk.emit_op(Op::DROP, line);
        s
    };
    emit_assoc_sort_impl(chunks, current, 4, Some(cmp_slot), line);
}
pub fn emit_php_uksort(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let cmp_slot = {
        let chunk = &mut chunks[current];
        let s = alloc_local(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, s, line);
        chunk.emit_op(Op::DROP, line);
        s
    };
    emit_assoc_sort_impl(chunks, current, 5, Some(cmp_slot), line);
}

/// PHP `array_merge($a, $b, ...)`. Numeric keys → reindexed (appended).
/// String keys → preserved (later wins). Uses ecma:object.entries to
/// distinguish key types. Output is always a Map to preserve string keys.
pub fn emit_php_array_merge(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Save all args into an array (they arrive in reverse on stack)
    let (args_arr_slot, out_slot, vals_slot, i_slot, n_slot, j_slot, m_slot) = {
        let c = &mut chunks[current];
        (alloc_local(c), alloc_local(c), alloc_local(c),
         alloc_local(c), alloc_local(c), alloc_local(c), alloc_local(c))
    };
    {
        let c = &mut chunks[current];
        // Collect args into array (reversed on stack, so reverse after)
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        lset(c, args_arr_slot, line);
    }
    for _ in 0..argc {
        {
            let c = &mut chunks[current];
            let tmp = alloc_local(c);
            lset(c, tmp, line);
            lget(c, args_arr_slot, line);
            lget(c, tmp, line);
        }
        call_import(chunks, current, "ecma:array", "push", 2, line);
        { let c = &mut chunks[current]; c.emit_op(Op::DROP, line); }
    }
    // Reverse so first arg is at index 0
    { let c = &mut chunks[current]; lget(c, args_arr_slot, line); }
    call_import(chunks, current, "ecma:array", "reverse", 1, line);
    { let c = &mut chunks[current]; c.emit_op(Op::DROP, line); }
    // Output: Map (to support both string and numeric keys).
    // Numeric keys get auto-incremented, string keys are set directly.
    let (idx_slot, entry_slot, key_slot) = {
        let c = &mut chunks[current];
        (alloc_local(c), alloc_local(c), alloc_local(c))
    };
    call_import(chunks, current, "ecma:map", "new", 0, line);
    {
        let c = &mut chunks[current];
        lset(c, out_slot, line);
        push_const(c, Value::F64(0.0), line);
        lset(c, idx_slot, line); // auto-increment index for numeric keys
        push_const(c, Value::F64(0.0), line);
        lset(c, i_slot, line);
        lget(c, args_arr_slot, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        lset(c, n_slot, line);
    }
    // Outer loop: each arg
    let lp1 = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, i_slot, line); lget(c, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    // entries = ecma:object.entries(args[i])
    {
        let c = &mut chunks[current];
        lget(c, args_arr_slot, line); lget(c, i_slot, line);
        c.emit_op(Op::ARRAY_GET, line);
    }
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    {
        let c = &mut chunks[current];
        lset(c, vals_slot, line); // reuse as entries
        push_const(c, Value::F64(0.0), line); lset(c, j_slot, line);
        lget(c, vals_slot, line); c.emit_op(Op::ARRAY_LENGTH, line);
        lset(c, m_slot, line);
    }
    // Inner loop: for each entry, check if key is numeric string
    let lp2 = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, j_slot, line); lget(c, m_slot, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, vals_slot, line); lget(c, j_slot, line);
        c.emit_op(Op::ARRAY_GET, line); lset(c, entry_slot, line);
        // key = entry[0]
        lget(c, entry_slot, line); push_const(c, Value::F64(0.0), line);
        c.emit_op(Op::ARRAY_GET, line); lset(c, key_slot, line);
        // Check if key is numeric (I32 type) — numeric keys reindex
        lget(c, key_slot, line);
    }
    call_import(chunks, current, "wasm:js-number", "test", 1, line);
    {
        let c = &mut chunks[current];
        c.emit_if(line);
        // Numeric key → append with auto-increment index
        lget(c, out_slot, line);
        lget(c, idx_slot, line);
        lget(c, entry_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_op(Op::ARRAY_SET, line); c.emit_op(Op::DROP, line);
        // idx++
        lget(c, idx_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line); lset(c, idx_slot, line);
        c.emit_else(line);
        // String key → set with original key (later wins)
        lget(c, out_slot, line);
        lget(c, key_slot, line);
        lget(c, entry_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_op(Op::ARRAY_SET, line); c.emit_op(Op::DROP, line);
        c.emit_end(line);
        lget(c, j_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line); lset(c, j_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, lp2, line);
    {
        let c = &mut chunks[current];
        lget(c, i_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line); lset(c, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, lp1, line);
    // If no string keys were used, convert Map → Array (values only)
    // Check: idx_slot == total count of entries in out → all numeric
    {
        let c = &mut chunks[current];
        lget(c, out_slot, line);
    }
    call_import(chunks, current, "ecma:map", "size", 1, line);
    {
        let c = &mut chunks[current];
        lget(c, idx_slot, line);
        crate::emitter::ops::emit_dyn_eq(c, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
        c.emit_if_value(line);
        // All numeric → return values as Array
        lget(c, out_slot, line);
    }
    call_import(chunks, current, "ecma:object", "values", 1, line);
    {
        let c = &mut chunks[current];
        c.emit_else(line);
        lget(c, out_slot, line);
        c.emit_end(line);
    }
}

/// PHP `array_unique($arr)` — remove duplicate values, preserve keys.
/// Uses ecma:object.entries + a seen-set to deduplicate.
pub fn emit_php_array_unique(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (arr_slot, entries_slot, out_slot, seen_slot, i_slot, n_slot, entry_slot, val_slot) = {
        let c = &mut chunks[current];
        (alloc_local(c), alloc_local(c), alloc_local(c), alloc_local(c),
         alloc_local(c), alloc_local(c), alloc_local(c), alloc_local(c))
    };
    { let c = &mut chunks[current]; lset(c, arr_slot, line); }
    // entries = ecma:object.entries(arr)
    { let c = &mut chunks[current]; lget(c, arr_slot, line); }
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    { let c = &mut chunks[current]; lset(c, entries_slot, line); }
    // out = new Map, seen = new Set (use Map as set — key=value, val=true)
    call_import(chunks, current, "ecma:map", "new", 0, line);
    { let c = &mut chunks[current]; lset(c, out_slot, line); }
    call_import(chunks, current, "ecma:map", "new", 0, line);
    {
        let c = &mut chunks[current];
        lset(c, seen_slot, line);
        push_const(c, Value::F64(0.0), line); lset(c, i_slot, line);
        lget(c, entries_slot, line); c.emit_op(Op::ARRAY_LENGTH, line);
        lset(c, n_slot, line);
    }
    let lp = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, i_slot, line); lget(c, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, entries_slot, line); lget(c, i_slot, line);
        c.emit_op(Op::ARRAY_GET, line); lset(c, entry_slot, line);
        // val = entry[1] converted to string for comparison
        lget(c, entry_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::ARRAY_GET, line); lset(c, val_slot, line);
        // Check if seen has this value
        lget(c, seen_slot, line); lget(c, val_slot, line);
    }
    call_import(chunks, current, "ecma:map", "has", 2, line);
    {
        let c = &mut chunks[current];
        crate::emitter::ops::emit_dyn_to_bool(c, line);
        c.emit_op(Op::I32_EQZ, line); // NOT seen
        c.emit_if(line);
        // Not seen → add to output and mark seen
        lget(c, out_slot, line);
        lget(c, entry_slot, line); push_const(c, Value::F64(0.0), line);
        c.emit_op(Op::ARRAY_GET, line); // key
        lget(c, val_slot, line); // value
        c.emit_op(Op::ARRAY_SET, line); c.emit_op(Op::DROP, line);
        lget(c, seen_slot, line); lget(c, val_slot, line);
        push_const(c, Value::Bool(true), line);
    }
    call_import(chunks, current, "ecma:map", "set", 3, line);
    {
        let c = &mut chunks[current];
        c.emit_op(Op::DROP, line);
        c.emit_end(line);
        lget(c, i_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line); lset(c, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, lp, line);
    { let c = &mut chunks[current]; lget(c, out_slot, line); }
}

/// PHP `$a + $b` array union — first-wins merge via ecma:object.entries.
pub fn emit_php_array_union(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let (a_slot, b_slot, out_slot, entries_slot, i_slot, n_slot, entry_slot, key_slot) = {
        let c = &mut chunks[current];
        (alloc_local(c), alloc_local(c), alloc_local(c), alloc_local(c),
         alloc_local(c), alloc_local(c), alloc_local(c), alloc_local(c))
    };
    {
        let c = &mut chunks[current];
        lset(c, b_slot, line);
        lset(c, a_slot, line);
    }
    // out = new Map from a's entries (copy)
    call_import(chunks, current, "ecma:map", "new", 0, line);
    { let c = &mut chunks[current]; lset(c, out_slot, line); }
    // Copy a's entries into out
    { let c = &mut chunks[current]; lget(c, a_slot, line); }
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    {
        let c = &mut chunks[current];
        lset(c, entries_slot, line);
        push_const(c, Value::F64(0.0), line);
        lset(c, i_slot, line);
        lget(c, entries_slot, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        lset(c, n_slot, line);
    }
    let lp1 = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, i_slot, line); lget(c, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, entries_slot, line); lget(c, i_slot, line);
        c.emit_op(Op::ARRAY_GET, line); lset(c, entry_slot, line);
        lget(c, out_slot, line);
        lget(c, entry_slot, line); push_const(c, Value::F64(0.0), line);
        c.emit_op(Op::ARRAY_GET, line);
        lget(c, entry_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_op(Op::ARRAY_SET, line); c.emit_op(Op::DROP, line);
        lget(c, i_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line); lset(c, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, lp1, line);
    // Now add b's entries only if key doesn't already exist
    { let c = &mut chunks[current]; lget(c, b_slot, line); }
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    {
        let c = &mut chunks[current];
        lset(c, entries_slot, line);
        push_const(c, Value::F64(0.0), line); lset(c, i_slot, line);
        lget(c, entries_slot, line); c.emit_op(Op::ARRAY_LENGTH, line);
        lset(c, n_slot, line);
    }
    let lp2 = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, i_slot, line); lget(c, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, entries_slot, line); lget(c, i_slot, line);
        c.emit_op(Op::ARRAY_GET, line); lset(c, entry_slot, line);
        // key = entry[0]
        lget(c, entry_slot, line); push_const(c, Value::F64(0.0), line);
        c.emit_op(Op::ARRAY_GET, line); lset(c, key_slot, line);
        // check if out already has key via ARRAY_GET + REF_IS_NULL
        lget(c, out_slot, line); lget(c, key_slot, line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_if(line);
        // key doesn't exist → set it
        lget(c, out_slot, line); lget(c, key_slot, line);
        lget(c, entry_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_op(Op::ARRAY_SET, line); c.emit_op(Op::DROP, line);
        c.emit_end(line);
        lget(c, i_slot, line); push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line); lset(c, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, lp2, line);
    { let c = &mut chunks[current]; lget(c, out_slot, line); }
}

/// PHP `array_reverse($arr, $preserve_keys?)`.
/// Uses `ecma:object.entries` → reverse → rebuild.
pub fn emit_php_array_reverse(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (preserve_slot, arr_slot, entries_slot, out_slot, i_slot, entry_slot) = {
        let c = &mut chunks[current];
        (alloc_local(c), alloc_local(c), alloc_local(c),
         alloc_local(c), alloc_local(c), alloc_local(c))
    };
    {
        let c = &mut chunks[current];
        if argc >= 2 { lset(c, preserve_slot, line); }
        else { push_const(c, Value::Bool(false), line); lset(c, preserve_slot, line); }
        lset(c, arr_slot, line);
        lget(c, arr_slot, line);
    }
    // entries = ecma:object.entries(arr)
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    {
        let c = &mut chunks[current];
        lset(c, entries_slot, line);
        // reverse entries in place
        lget(c, entries_slot, line);
    }
    call_import(chunks, current, "ecma:array", "reverse", 1, line);
    { let c = &mut chunks[current]; c.emit_op(Op::DROP, line); }
    // Decide output: preserve_keys → Map, else → Array
    {
        let c = &mut chunks[current];
        lget(c, preserve_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
        c.emit_if_value(line);
    }
    call_import(chunks, current, "ecma:map", "new", 0, line);
    {
        let c = &mut chunks[current];
        c.emit_else(line);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        c.emit_end(line);
        lset(c, out_slot, line);
        push_const(c, Value::F64(0.0), line);
        lset(c, i_slot, line);
    }
    let lp = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, i_slot, line);
        lget(c, entries_slot, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, entries_slot, line);
        lget(c, i_slot, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, entry_slot, line);
        lget(c, preserve_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        // Map set: out[entry[0]] = entry[1]
        lget(c, out_slot, line);
        lget(c, entry_slot, line);
        push_const(c, Value::F64(0.0), line);
        c.emit_op(Op::ARRAY_GET, line);
        lget(c, entry_slot, line);
        push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_op(Op::ARRAY_SET, line);
        c.emit_op(Op::DROP, line);
        c.emit_else(line);
        // Array push: out.push(entry[1])
        lget(c, out_slot, line);
        lget(c, entry_slot, line);
        push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::ARRAY_GET, line);
    }
    call_import(chunks, current, "ecma:array", "push", 2, line);
    {
        let c = &mut chunks[current];
        c.emit_op(Op::DROP, line);
        c.emit_end(line);
        lget(c, i_slot, line);
        push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line);
        lset(c, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, lp, line);
    { let c = &mut chunks[current]; lget(c, out_slot, line); }
}

/// PHP `array_slice($arr, $offset, $length?, $preserve_keys?)`.
/// Uses `ecma:object.entries` to handle both Array and Map inputs.
/// Returns a Map (preserving keys) when preserve_keys=true or input is Map,
/// otherwise returns an Array (reindexed).
pub fn emit_php_array_slice(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (preserve_slot, len_slot, offset_slot, arr_slot, entries_slot,
         out_slot, i_slot, n_slot, entry_slot) = {
        let c = &mut chunks[current];
        (alloc_local(c), alloc_local(c), alloc_local(c), alloc_local(c),
         alloc_local(c), alloc_local(c), alloc_local(c), alloc_local(c),
         alloc_local(c))
    };
    // Pop args
    {
        let c = &mut chunks[current];
        if argc >= 4 { lset(c, preserve_slot, line); }
        else { push_const(c, Value::Bool(false), line); lset(c, preserve_slot, line); }
        if argc >= 3 {
            lset(c, len_slot, line);
        } else {
            push_const(c, Value::I32(i32::MAX), line);
            lset(c, len_slot, line);
        }
        lset(c, offset_slot, line);
        lset(c, arr_slot, line);
    }
    // entries = ecma:object.entries(arr) — works for both Array and Map
    {
        let c = &mut chunks[current];
        lget(c, arr_slot, line);
    }
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    {
        let c = &mut chunks[current];
        lset(c, entries_slot, line);
        // n = entries.length
        lget(c, entries_slot, line);
        c.emit_op(Op::ARRAY_LENGTH, line);
        lset(c, n_slot, line);
        // Normalize negative offset
        lget(c, offset_slot, line);
        push_const(c, Value::F64(0.0), line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        lget(c, n_slot, line);
        lget(c, offset_slot, line);
        c.emit_op(Op::F64_ADD, line);
        push_const(c, Value::F64(0.0), line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        push_const(c, Value::F64(0.0), line);
        lset(c, offset_slot, line);
        c.emit_else(line);
        lget(c, n_slot, line);
        lget(c, offset_slot, line);
        c.emit_op(Op::F64_ADD, line);
        lset(c, offset_slot, line);
        c.emit_end(line);
        c.emit_end(line);
        // Decide output type: preserve_keys → Map, else → Array
        lget(c, preserve_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
        c.emit_if_value(line);
    }
    call_import(chunks, current, "ecma:map", "new", 0, line);
    {
        let c = &mut chunks[current];
        c.emit_else(line);
        c.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
        c.emit_end(line);
        lset(c, out_slot, line);
        // i = offset
        lget(c, offset_slot, line);
        lset(c, i_slot, line);
    }
    // Loop: copy entries[offset..offset+length]
    let lp = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        // cond: i < n && i < offset + length
        lget(c, i_slot, line);
        lget(c, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        lget(c, i_slot, line);
        lget(c, offset_slot, line);
        lget(c, len_slot, line);
        c.emit_op(Op::F64_ADD, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
        c.emit_op(Op::I32_AND, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        // entry = entries[i]
        lget(c, entries_slot, line);
        lget(c, i_slot, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, entry_slot, line);
        // if preserve_keys: out[entry[0]] = entry[1]
        // else: out.push(entry[1])
        lget(c, preserve_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        lget(c, out_slot, line);
        lget(c, entry_slot, line);
        push_const(c, Value::F64(0.0), line);
        c.emit_op(Op::ARRAY_GET, line);
        lget(c, entry_slot, line);
        push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_op(Op::ARRAY_SET, line);
        c.emit_op(Op::DROP, line);
        c.emit_else(line);
        lget(c, out_slot, line);
        lget(c, entry_slot, line);
        push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::ARRAY_GET, line);
    }
    call_import(chunks, current, "ecma:array", "push", 2, line);
    {
        let c = &mut chunks[current];
        c.emit_op(Op::DROP, line);
        c.emit_end(line);
        // i++
        lget(c, i_slot, line);
        push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line);
        lset(c, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, lp, line);
    { let c = &mut chunks[current]; lget(c, out_slot, line); }
}

// ── implode (PHP bool-to-string coercion) ─────────────────────────
/// PHP `implode($glue, $arr)` — like JS join but bools → "1"/"".
/// Stack: [arr, glue] → [string].
pub fn emit_php_implode(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let glue_slot = alloc_local(chunk);
    let arr_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let v_slot = alloc_local(chunk);
    let str_slot = alloc_local(chunk);
    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");

    // Args come as (arr, glue) after walker reorder
    lset(chunk, glue_slot, line);
    lset(chunk, arr_slot, line);

    // Build a new array with stringified elements
    chunk.emit_op_u8(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, out_slot, line);

    // Get values from array (works for both Array and Map)
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "values", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, arr_slot, line);
    lget(chunk, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let lp = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, i_slot, line);
        lget(c, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, arr_slot, line);
        lget(c, i_slot, line);
        c.emit_op(Op::ARRAY_GET, line);
        lset(c, v_slot, line);

        // null → ""
        lget(c, v_slot, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_if(line);
        push_str(c, "", line);
        lset(c, str_slot, line);
        c.emit_else(line);
        // boolean check
        lget(c, v_slot, line);
        c.emit_call(test_bool, 1, line);
        c.emit_if(line);
        lget(c, v_slot, line);
        c.emit_call(cast_bool, 1, line);
        c.emit_if(line);
        push_str(c, "1", line);
        lset(c, str_slot, line);
        c.emit_else(line);
        push_str(c, "", line);
        lset(c, str_slot, line);
        c.emit_end(line);
        c.emit_else(line);
        lget(c, v_slot, line);
        lset(c, str_slot, line);
        c.emit_end(line);
        c.emit_end(line);

        lget(c, out_slot, line);
        lget(c, str_slot, line);
    }
    call_import(chunks, current, "ecma:array", "push", 2, line);
    {
        let c = &mut chunks[current];
        c.emit_op(Op::DROP, line);
        // i++
        lget(c, i_slot, line);
        push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line);
        lset(c, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, lp, line);
    // join the stringified array
    {
        let c = &mut chunks[current];
        lget(c, out_slot, line);
        lget(c, glue_slot, line);
    }
    call_import(chunks, current, "ecma:array", "join", 2, line);
}

// ── in_array (loose by default) ───────────────────────────────────
/// PHP `in_array($needle, $haystack, $strict=false)`.
/// After arg reorder: stack [haystack, needle, strict?] → [bool].
pub fn emit_php_in_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let strict_slot = alloc_local(chunk);
    let needle_slot = alloc_local(chunk);
    let arr_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let found_slot = alloc_local(chunk);

    if argc >= 3 {
        lset(chunk, strict_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, strict_slot, line);
    }
    lset(chunk, needle_slot, line);
    lset(chunk, arr_slot, line);

    push_const(chunk, Value::Bool(false), line);
    lset(chunk, found_slot, line);

    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, n_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);

    let _ = chunk;
    let lp = crate::emitter::loops::emit_loop_start(chunks, current, line);
    {
        let c = &mut chunks[current];
        lget(c, i_slot, line);
        lget(c, n_slot, line);
        crate::emitter::ops::emit_dyn_lt(c, line);
    }
    crate::emitter::loops::emit_loop_cond(chunks, current, line);
    {
        let c = &mut chunks[current];
        // val = arr[keys[i]]
        lget(c, arr_slot, line);
        lget(c, keys_slot, line);
        lget(c, i_slot, line);
        c.emit_op(Op::ARRAY_GET, line);
        c.emit_op(Op::ARRAY_GET, line);

        // Compare with needle
        let val_tmp = alloc_local(c);
        lset(c, val_tmp, line);

        lget(c, strict_slot, line);
        crate::emitter::ops::emit_dyn_to_bool(c, line);
        c.emit_if(line);
        // strict: val === needle
        lget(c, val_tmp, line);
        lget(c, needle_slot, line);
        crate::emitter::ops::emit_js_strict_eq(c, line);
        c.emit_else(line);
        // loose: dyn_eq first, then numeric coercion
        lget(c, val_tmp, line);
        lget(c, needle_slot, line);
        crate::emitter::ops::emit_dyn_eq(c, line);
        // dyn_eq returns i32. If 1, found. If 0, try numeric coercion.
        c.emit_if(line);
        push_const(c, Value::I32(1), line);
        c.emit_else(line);
        // parseFloat(toString(val)) == parseFloat(toString(needle))
        lget(c, val_tmp, line);
        crate::emitter::strings::emit_to_string(c, line);
    }
    call_import(chunks, current, "ecma:number", "parseFloat", 1, line);
    {
        let c = &mut chunks[current];
        lget(c, needle_slot, line);
        crate::emitter::strings::emit_to_string(c, line);
    }
    call_import(chunks, current, "ecma:number", "parseFloat", 1, line);
    {
        let c = &mut chunks[current];
        c.emit_op(Op::F64_EQ, line); // i32: 0 if NaN or different
        c.emit_end(line); // end dyn_eq if/else
        c.emit_end(line); // end strict if/else

        // result is i32 (0 or 1) on stack
        c.emit_if(line);
        push_const(c, Value::Bool(true), line);
        lset(c, found_slot, line);
        c.emit_end(line);

        // i++
        lget(c, i_slot, line);
        push_const(c, Value::F64(1.0), line);
        c.emit_op(Op::F64_ADD, line);
        lset(c, i_slot, line);
    }
    crate::emitter::loops::emit_loop_end(chunks, current, lp, line);
    let c = &mut chunks[current];
    lget(c, found_slot, line);
}

// ── (object)$arr — array to object ───────────────────────────────
/// Stack: [arr] → [object].
/// Uses `ecma:object.entries` → `ecma:object.fromEntries`.
pub fn emit_php_array_to_object(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // entries(arr) → [[k,v],...] → fromEntries → object
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    call_import(chunks, current, "ecma:object", "fromEntries", 1, line);
}

// ── (array)$obj — object to array ────────────────────────────────
/// Stack: [obj] → [map].
/// Uses `ecma:object.entries` → `ecma:map.fromEntries`.
pub fn emit_php_obj_to_array(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // entries(obj) → [[k,v],...] → map.fromEntries → Map
    call_import(chunks, current, "ecma:object", "entries", 1, line);
    call_import(chunks, current, "ecma:map", "fromEntries", 1, line);
}

// ── var_export ────────────────────────────────────────────────────
/// PHP `var_export($val [, $return])`.
/// Stack: [val, return?] → [string|null].
pub fn emit_php_var_export(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let return_slot = alloc_local(chunk);
    let val_slot = alloc_local(chunk);

    if argc >= 2 {
        lset(chunk, return_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, return_slot, line);
    }
    lset(chunk, val_slot, line);

    lget(chunk, val_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:json", "stringify", 1, line);
    let chunk = &mut chunks[current];

    lget(chunk, return_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // return mode — string stays on stack
    chunk.emit_else(line);
    // echo mode — use wasi:cli/stdout write (no newline)
    let write_idx = chunk.add_import("wasi:cli/stdout", "write-via-stream");
    let out_slot = alloc_local(chunk);
    lset(chunk, out_slot, line);
    let rd_slot = alloc_local(chunk);
    let wr_slot = alloc_local(chunk);
    crate::emitter::io::emit_write_stdout_with_imports(
        chunk, write_idx, rd_slot, wr_slot, line,
        |c| c.emit_op_u16(Op::LOCAL_GET, out_slot, line),
    );
    chunk.emit_op(Op::NULL, line);
    chunk.emit_end(line);
}

// ── print_r ──────────────────────────────────────────────────────
/// PHP `print_r($val [, $return])`.
/// Stack: [val, return?] → [string|true|null].
pub fn emit_php_print_r(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let return_slot = alloc_local(chunk);
    let val_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);

    if argc >= 2 {
        lset(chunk, return_slot, line);
    } else {
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, return_slot, line);
    }
    lset(chunk, val_slot, line);

    // Check if val is array-like (object test)
    lget(chunk, val_slot, line);
    emit_test_object(chunk, line);
    chunk.emit_if(line);
    push_str(chunk, "Array\n(\n)", line);
    lset(chunk, result_slot, line);
    chunk.emit_else(line);
    lget(chunk, val_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:json", "stringify", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);
    chunk.emit_end(line);

    lget(chunk, return_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    // return mode
    lget(chunk, result_slot, line);
    chunk.emit_else(line);
    // echo mode — use wasi:cli/stdout write (no newline)
    let write_idx = chunk.add_import("wasi:cli/stdout", "write-via-stream");
    lget(chunk, result_slot, line);
    let out_slot = alloc_local(chunk);
    lset(chunk, out_slot, line);
    let rd_slot = alloc_local(chunk);
    let wr_slot = alloc_local(chunk);
    crate::emitter::io::emit_write_stdout_with_imports(
        chunk, write_idx, rd_slot, wr_slot, line,
        |c| c.emit_op_u16(Op::LOCAL_GET, out_slot, line),
    );
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_end(line);
}
