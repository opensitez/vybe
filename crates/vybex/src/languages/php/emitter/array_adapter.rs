//! PHP array helpers — Rust inline opcode emitters.
//!
//! Mirrors the inline-emit shape in `datetime_adapter.rs`: each
//! `emit_*(chunks, current, argc, line)` writes WASM opcodes directly
//! into `chunks[current]`. Composes only WASM ops + `ecma:array.*` /
//! `ecma:object.*` host imports — no PHP-specific host fns; no JS
//! polyfills. PHP `array` ≡ JS `Map` (assoc) or `Array` (sequential)
//! per the cross-language type model.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
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
fn call_import(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[0].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}

fn emit_is_array(chunks: &mut [Chunk], current: usize, arr_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    chunks[current].emit_op(Op::DYN_TO_BOOL, line);
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

fn emit_object_from_keys(chunks: &mut [Chunk], current: usize, source_slot: u16, keys_slot: u16, line: u32) {
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

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

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
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, entries_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "fromEntries", 1, line);
}

pub fn emit_php_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let mode_slot = if argc >= 2 { Some(alloc_local(chunk)) } else { None };
    let value_slot = alloc_local(chunk);
    let base_len_slot = alloc_local(chunk);
    let extra_len_slot = alloc_local(chunk);

    if let Some(slot) = mode_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, value_slot, line);

    emit_is_array(chunks, current, value_slot, line);
    let chunk = &mut chunks[current];
    let not_array = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, base_len_slot, line);

    lget(chunk, value_slot, line);
    let assoc_key = chunk.add_constant(Value::String(Arc::from("vybe$assoc_keys_csv")));
    chunk.emit_op_u16(Op::STRUCT_GET, assoc_key, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let no_keys = chunk.emit_jump(Op::BR_IF_TRUE, line);
    push_str(chunk, "\x1F", line);
    let _ = chunk;
    call_import(chunks, current, "ecma:string", "split", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, extra_len_slot, line);
    lget(chunk, base_len_slot, line);
    lget(chunk, extra_len_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    let done = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(no_keys);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, base_len_slot, line);
    let done_no_keys = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(not_array);
    lget(chunk, value_slot, line);
    crate::emitter::collections::emit_len(chunks, current, line);

    chunks[current].patch_jump(done);
    chunks[current].patch_jump(done_no_keys);
}

pub fn emit_php_json_encode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let depth_slot = if argc >= 3 { Some(alloc_local(chunk)) } else { None };
    let flags_slot = if argc >= 2 { Some(alloc_local(chunk)) } else { None };
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

    emit_is_array(chunks, current, value_slot, line);
    let chunk = &mut chunks[current];
    let not_array = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, value_slot, line);
    let assoc_key = chunk.add_constant(Value::String(Arc::from("vybe$assoc_keys_csv")));
    chunk.emit_op_u16(Op::STRUCT_GET, assoc_key, line);
    chunk.emit_op(Op::DUP, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let no_assoc = chunk.emit_jump(Op::BR_IF_TRUE, line);
    push_str(chunk, "\x1F", line);
    let _ = chunk;
    call_import(chunks, current, "ecma:string", "split", 2, line);
    let chunk = &mut chunks[current];
    lset(chunk, assoc_keys_slot, line);
    lget(chunk, assoc_keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op(Op::I32_CONST_0, line);
    chunk.emit_op(Op::DYN_GT, line);
    let has_assoc = chunk.emit_jump(Op::BR_IF_TRUE, line);

    lget(chunk, value_slot, line);
    lset(chunk, render_slot, line);
    let after_array = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(no_assoc);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, value_slot, line);
    lset(chunk, render_slot, line);
    let after_no_assoc = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(has_assoc);
    emit_object_from_keys(chunks, current, value_slot, assoc_keys_slot, line);
    let chunk = &mut chunks[current];
    lset(chunk, render_slot, line);
    let after_assoc = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(not_array);
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, object_keys_slot, line);
    lget(chunk, object_keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op(Op::I32_CONST_0, line);
    chunk.emit_op(Op::DYN_GT, line);
    let has_object_keys = chunk.emit_jump(Op::BR_IF_TRUE, line);

    lget(chunk, value_slot, line);
    lset(chunk, render_slot, line);
    let after_object = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(has_object_keys);
    emit_object_from_keys(chunks, current, value_slot, object_keys_slot, line);
    let chunk = &mut chunks[current];
    lset(chunk, render_slot, line);

    chunks[current].patch_jump(after_array);
    chunks[current].patch_jump(after_no_assoc);
    chunks[current].patch_jump(after_assoc);
    chunks[current].patch_jump(after_object);

    emit_json_stringify_slots(chunks, current, render_slot, flags_slot, depth_slot, argc, line);
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
) where F: FnMut(&mut [Chunk], usize) {
    // Branch on `typeof fn === "function"`.
    let chunk = &mut chunks[current];
    lget(chunk, fn_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "function", line);
    chunk.emit_op(Op::DYN_EQ, line);
    let not_func = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // Function: call directly.
    let chunk = &mut chunks[current];
    lget(chunk, fn_slot, line);
    push_args(chunks, current);
    let chunk = &mut chunks[current];
    chunk.emit_op_u8(Op::CALL_REF, argc, line);
    let done = chunk.emit_jump(Op::BR, line);

    // Object: call $obj->__invoke(args). PHP method ABI passes `$this`
    // as arg0, so push fn (the receiver) twice — once as the function
    // ref (resolved via STRUCT_GET on __invoke), once as `$this`.
    chunk.patch_jump(not_func);
    lget(chunk, fn_slot, line);
    let invoke_key = chunk.add_constant(Value::String(Arc::from("__invoke")));
    chunk.emit_op_u16(Op::STRUCT_GET, invoke_key, line);
    lget(chunk, fn_slot, line);
    push_args(chunks, current);
    let chunk = &mut chunks[current];
    chunk.emit_op_u8(Op::CALL_REF, argc + 1, line);

    chunks[current].patch_jump(done);
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
    let not_array = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    let out_ready = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(not_array);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    chunk.patch_jump(out_ready);
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

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

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
    let map_branch = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, out_slot, line);
    lget(chunk, mapped_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    let after_store = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(map_branch);
    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, mapped_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.patch_jump(after_store);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

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

    lget(chunk, is_array_slot, line);
    let not_array = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    let out_ready = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(not_array);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    chunk.patch_jump(out_ready);
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

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

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
    let has_callback = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    let after_predicate = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(has_callback);

    lget(chunk, flag_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let not_use_key = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, fn_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    let after_callback = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(not_use_key);

    lget(chunk, flag_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let not_use_both = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, fn_slot, line);
    lget(chunk, value_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    let after_both = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(not_use_both);

    lget(chunk, fn_slot, line);
    lget(chunk, value_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 1, line);
    chunk.patch_jump(after_both);
    chunk.patch_jump(after_callback);
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    chunk.patch_jump(after_predicate);

    let skip_store = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, is_array_slot, line);
    let map_branch = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, out_slot, line);
    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    let after_store = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(map_branch);
    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.patch_jump(after_store);
    chunk.patch_jump(skip_store);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, out_slot, line);
}

pub fn emit_array_walk_recursive(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let userdata_slot = alloc_local(chunk);
    let fn_slot = alloc_local(chunk);
    let arr_slot = alloc_local(chunk);
    let work_slot = alloc_local(chunk);
    let cur_slot = alloc_local(chunk);
    let ty_slot = alloc_local(chunk);
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

    let chunk = &mut chunks[current];
    let loop_top = chunk.current_offset();
    lget(chunk, work_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_GT, line);
    let done = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, work_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "pop", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, cur_slot, line);

    lget(chunk, cur_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let non_null = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(non_null);

    lget(chunk, cur_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    lset(chunk, ty_slot, line);

    lget(chunk, ty_slot, line);
    push_str(chunk, "array", line);
    chunk.emit_op(Op::DYN_EQ, line);
    let not_array = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, cur_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, i_slot, line);

    let array_loop = chunk.current_offset();
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_GT, line);
    let array_done = chunk.emit_jump(Op::BR_IF_FALSE, line);
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
    chunk.emit_loop(array_loop, line);
    chunk.patch_jump(array_done);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(not_array);

    lget(chunk, ty_slot, line);
    push_str(chunk, "object", line);
    chunk.emit_op(Op::DYN_EQ, line);
    let not_object = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, cur_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);
    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, i_slot, line);

    let object_loop = chunk.current_offset();
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_GT, line);
    let object_done = chunk.emit_jump(Op::BR_IF_FALSE, line);
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
    chunk.emit_loop(object_loop, line);
    chunk.patch_jump(object_done);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(not_object);

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
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(done);
    chunk.emit_op(Op::TRUE, line);
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
    chunk.emit_op(Op::DYN_LT, line);
    let nonneg = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, size_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    let after_abs = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(nonneg);
    lget(chunk, size_slot, line);
    chunk.patch_jump(after_abs);
    lset(chunk, target_slot, line);

    // if target <= len: return arr.slice() (just a clone)
    lget(chunk, target_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_GT, line);
    let needs_pad = chunk.emit_jump(Op::BR_IF_TRUE, line);
    // No pad: clone via ecma:array.slice(arr)
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "slice", 1, line);
    let chunk = &mut chunks[current];
    let done_short = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(needs_pad);

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
    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, diff_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

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
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    // result = size < 0 ? pad.concat(arr) : arr.slice().concat(pad)
    lget(chunk, size_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let pad_right = chunk.emit_jump(Op::BR_IF_FALSE, line);
    // Pad-left: pad.concat(arr)
    lget(chunk, pad_slot, line);
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "concat", 2, line);
    let chunk = &mut chunks[current];
    let done_left_pad = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(pad_right);
    // Pad-right: arr.concat(pad)
    lget(chunk, arr_slot, line);
    lget(chunk, pad_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "concat", 2, line);
    let chunk = &mut chunks[current];
    chunk.patch_jump(done_left_pad);
    chunk.patch_jump(done_short);
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
        chunk.emit_op(Op::DYN_TO_BOOL, line);
        lset(chunk, preserve_slot, line);
    } else {
        chunk.emit_op(Op::FALSE, line);
        lset(chunk, preserve_slot, line);
    }
    lset(chunk, size_slot, line);
    lset(chunk, arr_slot, line);

    // out = []
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    lset(chunk, out_slot, line);

    // if size < 1: return out
    lget(chunk, size_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::DYN_LT, line);
    let valid = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, out_slot, line);
    let done_invalid = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(valid);

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
    let outer_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let outer_exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // end = min(i + size, n)
    lget(chunk, i_slot, line);
    lget(chunk, size_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, end_slot, line);
    lget(chunk, end_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_GT, line);
    let in_bounds = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, n_slot, line);
    lset(chunk, end_slot, line);
    chunk.patch_jump(in_bounds);

    // chunk_obj = preserve ? ecma:map.new() : []
    lget(chunk, preserve_slot, line);
    let scalar_chunk = chunk.emit_jump(Op::BR_IF_FALSE, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let chunk = &mut chunks[current];
    let after_chunk_init = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(scalar_chunk);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 0, line);
    chunk.patch_jump(after_chunk_init);
    lset(chunk, chunk_slot, line);

    // j = i
    lget(chunk, i_slot, line);
    lset(chunk, j_slot, line);

    // Inner loop: for j in i..end
    let inner_top = chunk.current_offset();
    lget(chunk, j_slot, line);
    lget(chunk, end_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let inner_exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // key = keys[j]
    lget(chunk, keys_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, key_slot, line);

    // if preserve: chunk_obj[key] = arr[key] ; else chunk_obj.push(arr[key])
    lget(chunk, preserve_slot, line);
    let scalar_push = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, chunk_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, arr_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    let after_push = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(scalar_push);
    lget(chunk, chunk_slot, line);
    lget(chunk, arr_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    let chunk = &mut chunks[current];
    chunk.patch_jump(after_push);

    // j++
    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    chunk.emit_loop(inner_top, line);
    chunk.patch_jump(inner_exit);

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
    chunk.emit_loop(outer_top, line);
    chunk.patch_jump(outer_exit);

    lget(chunk, out_slot, line);
    chunk.patch_jump(done_invalid);
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

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

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
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

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

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

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

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, out_slot, line);
}

// ── array_diff / array_intersect (value-only, sequential arrays) ──

fn emit_array_diff_or_intersect(
    chunks: &mut [Chunk],
    current: usize,
    intersect: bool,
    line: u32,
) {
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

    let loop1_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, blen_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit1 = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, b_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    push_str(chunk, "", line);
    let _ = chunk;
    chunks[current].emit_op(Op::DYN_ADD, line);
    let chunk = &mut chunks[current];
    lset(chunk, key_slot, line);
    lget(chunk, seen_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::TRUE, line);
    chunk.emit_op(Op::ARRAY_SET, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop1_top, line);
    chunk.patch_jump(exit1);

    // for j in 0..a.length: if (seen[String(a[j])] == intersect): out.push(a[j])
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, j_slot, line);
    lget(chunk, a_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, alen_slot, line);

    let loop2_top = chunk.current_offset();
    lget(chunk, j_slot, line);
    lget(chunk, alen_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit2 = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // v = a[j]; key = "" + v; has = seen[key]
    lget(chunk, a_slot, line);
    lget(chunk, j_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, v_slot, line);
    push_str(chunk, "", line);
    lget(chunk, v_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, key_slot, line);
    lget(chunk, seen_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, has_slot, line);

    // if intersect ? has : !has → push v
    lget(chunk, has_slot, line);
    let skip = if intersect {
        chunk.emit_jump(Op::BR_IF_FALSE, line)
    } else {
        chunk.emit_jump(Op::BR_IF_TRUE, line)
    };
    lget(chunk, out_slot, line);
    lget(chunk, v_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "push", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.patch_jump(skip);

    lget(chunk, j_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, j_slot, line);
    chunk.emit_loop(loop2_top, line);
    chunk.patch_jump(exit2);

    lget(chunk, out_slot, line);
}

pub fn emit_array_diff(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_array_diff_or_intersect(chunks, current, /*intersect=*/false, line);
}
pub fn emit_array_intersect(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_array_diff_or_intersect(chunks, current, /*intersect=*/true, line);
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

    let loop_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // key = "" + arr[i]
    push_str(chunk, "", line);
    lget(chunk, arr_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::DYN_ADD, line);
    lset(chunk, key_slot, line);

    // cur = (out[key] || 0) + 1
    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, cur_slot, line);
    // if cur is null/undefined: cur = 0
    lget(chunk, cur_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    let not_null = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, cur_slot, line);
    chunk.patch_jump(not_null);

    lget(chunk, out_slot, line);
    lget(chunk, key_slot, line);
    lget(chunk, cur_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::ARRAY_SET, line);

    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    lget(chunk, out_slot, line);
}

// ── array_column ───────────────────────────────────────────────────

pub fn emit_array_column(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let has_index = argc >= 3;
    // Allocate slots first.
    let (index_key_slot, col_slot, rows_slot, out_slot, i_slot, len_slot, row_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
            alloc_local(chunk),
        )
    };
    {
        let chunk = &mut chunks[current];
        if has_index { lset(chunk, index_key_slot, line); }
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

    let loop_top = chunks[current].current_offset();
    let exit = {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        chunk.emit_op(Op::DYN_LT, line);
        let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

        lget(chunk, rows_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, row_slot, line);

        if has_index {
            lget(chunk, out_slot, line);
            lget(chunk, row_slot, line);
            lget(chunk, index_key_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            lget(chunk, row_slot, line);
            lget(chunk, col_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op(Op::ARRAY_SET, line);
        }
        exit
    };
    if !has_index {
        {
            let chunk = &mut chunks[current];
            lget(chunk, out_slot, line);
            lget(chunk, row_slot, line);
            lget(chunk, col_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
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
        chunk.emit_loop(loop_top, line);
        chunk.patch_jump(exit);
        lget(chunk, out_slot, line);
    }
}

// ── array_key_first / array_key_last ───────────────────────────────

fn emit_array_key_first_or_last(
    chunks: &mut [Chunk],
    current: usize,
    last: bool,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let arr_slot = alloc_local(chunk);
    let keys_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);

    lset(chunk, arr_slot, line);

    // Object.keys(arr) — works for both arrays and Maps
    lget(chunk, arr_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, keys_slot, line);

    lget(chunk, keys_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len_slot, line);

    // if len === 0: return null
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let nonempty = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::NULL, line);
    let done_empty = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(nonempty);

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
    chunk.patch_jump(done_empty);
}

pub fn emit_array_key_first(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_array_key_first_or_last(chunks, current, /*last=*/false, line);
}
pub fn emit_array_key_last(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_array_key_first_or_last(chunks, current, /*last=*/true, line);
}

// ── array_diff_key / array_diff_assoc / array_intersect_key / array_replace ─────────

/// PHP `array_diff_key(a, b)` — entries in a whose keys do not exist in b.
pub fn emit_array_diff_key(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (b_slot, a_slot, out_slot, keys_slot, i_slot, len_slot, k_slot, av_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
            alloc_local(chunk), alloc_local(chunk),
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

    let loop_top = chunks[current].current_offset();
    let exit = {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        chunk.emit_op(Op::DYN_LT, line);
        let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

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
        call_import(chunks, current, "ecma:object", "hasOwn", 2, line);
        let exists = chunks[current].emit_jump(Op::BR_IF_TRUE, line);

        let chunk = &mut chunks[current];
        lget(chunk, out_slot, line);
        lget(chunk, k_slot, line);
        lget(chunk, av_slot, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.patch_jump(exists);

        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
        chunk.emit_loop(loop_top, line);
        chunk.patch_jump(exit);

        lget(chunk, out_slot, line);
        exit
    };
    let _ = exit;
}

/// PHP `array_diff_assoc(a, b)` — entries in a whose key→value pair
/// differs in b.
pub fn emit_array_diff_assoc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (b_slot, a_slot, out_slot, keys_slot, i_slot, len_slot, k_slot, av_slot, bv_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
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

    let loop_top = chunks[current].current_offset();
    let exit = {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        chunk.emit_op(Op::DYN_LT, line);
        let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

        // k = keys[i]
        lget(chunk, keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, k_slot, line);

        // av = a[k]
        lget(chunk, a_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, av_slot, line);

        // bv = b[k]
        lget(chunk, b_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, bv_slot, line);

        // if String(bv) !== String(av): out[k] = av
        push_str(chunk, "", line);
        lget(chunk, bv_slot, line);
        chunk.emit_op(Op::DYN_ADD, line);
        push_str(chunk, "", line);
        lget(chunk, av_slot, line);
        chunk.emit_op(Op::DYN_ADD, line);
        chunk.emit_op(Op::DYN_EQ, line);
        let same = chunk.emit_jump(Op::BR_IF_TRUE, line);
        // differ → keep
        lget(chunk, out_slot, line);
        lget(chunk, k_slot, line);
        lget(chunk, av_slot, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.patch_jump(same);

        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
        chunk.emit_loop(loop_top, line);
        chunk.patch_jump(exit);

        lget(chunk, out_slot, line);
        exit
    };
    let _ = exit;
}

/// PHP `array_intersect_key(a, b)` — entries from a whose keys exist in b.
pub fn emit_array_intersect_key(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (b_slot, a_slot, out_slot, keys_slot, i_slot, len_slot, k_slot, bv_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
            alloc_local(chunk), alloc_local(chunk),
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

    let loop_top = chunks[current].current_offset();
    {
        let chunk = &mut chunks[current];
        lget(chunk, i_slot, line);
        lget(chunk, len_slot, line);
        chunk.emit_op(Op::DYN_LT, line);
        let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

        lget(chunk, keys_slot, line);
        lget(chunk, i_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, k_slot, line);

        // bv = b[k]; if !is_null(bv): out[k] = a[k]
        lget(chunk, b_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        lset(chunk, bv_slot, line);
        lget(chunk, bv_slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        let no_key = chunk.emit_jump(Op::BR_IF_TRUE, line);
        lget(chunk, out_slot, line);
        lget(chunk, k_slot, line);
        lget(chunk, a_slot, line);
        lget(chunk, k_slot, line);
        chunk.emit_op(Op::ARRAY_GET, line);
        chunk.emit_op(Op::ARRAY_SET, line);
        chunk.patch_jump(no_key);

        lget(chunk, i_slot, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_ADD, line);
        lset(chunk, i_slot, line);
        chunk.emit_loop(loop_top, line);
        chunk.patch_jump(exit);

        lget(chunk, out_slot, line);
    }
}

/// PHP `array_replace(a, b)` — a + b, b's keys override a's.
pub fn emit_array_replace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let (b_slot, a_slot, out_slot, keys_slot, i_slot, len_slot, k_slot) = {
        let chunk = &mut chunks[current];
        (
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
            alloc_local(chunk), alloc_local(chunk), alloc_local(chunk),
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
        let loop_top = chunks[current].current_offset();
        {
            let chunk = &mut chunks[current];
            lget(chunk, i_slot, line);
            lget(chunk, len_slot, line);
            chunk.emit_op(Op::DYN_LT, line);
            let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

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

            lget(chunk, i_slot, line);
            push_const(chunk, Value::F64(1.0), line);
            chunk.emit_op(Op::F64_ADD, line);
            lset(chunk, i_slot, line);
            chunk.emit_loop(loop_top, line);
            chunk.patch_jump(exit);
        }
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
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
    let used_slot = alloc_local(chunk);          // Map<index, true>
    let sorted_keys_slot = alloc_local(chunk);   // Array of sorted keys
    let sorted_vals_slot = alloc_local(chunk);   // Array of sorted values
    let outer_slot = alloc_local(chunk);
    let inner_slot = alloc_local(chunk);
    let best_slot = alloc_local(chunk);

    // obj = pop()
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    chunk.emit_op(Op::DROP, line);

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

    let outer_top = chunk.current_offset();
    lget(chunk, outer_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let outer_exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // best = -1
    push_const(chunk, Value::F64(-1.0), line);
    lset(chunk, best_slot, line);

    // inner loop: for inner in 0..n
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, inner_slot, line);

    let inner_top = chunk.current_offset();
    lget(chunk, inner_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let inner_exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // if used[inner]: skip
    lget(chunk, used_slot, line);
    lget(chunk, inner_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::DYN_TO_BOOL, line);
    let unused = chunk.emit_jump(Op::BR_IF_FALSE, line);
    let skip_used = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(unused);

    // if best === -1: best = inner ; else compare
    lget(chunk, best_slot, line);
    push_const(chunk, Value::F64(-1.0), line);
    chunk.emit_op(Op::DYN_EQ, line);
    let have_best = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, inner_slot, line);
    lset(chunk, best_slot, line);
    let compared_done = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(have_best);

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
            chunk.emit_op(Op::DYN_ADD, line);
            lget(chunk, obj_slot, line);
            lget(chunk, keys_slot, line);
            lget(chunk, best_slot, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            chunk.emit_op(Op::ARRAY_GET, line);
            push_const(chunk, Value::F64(0.0), line);
            chunk.emit_op(Op::DYN_ADD, line);
            if mode == 0 {
                chunk.emit_op(Op::DYN_LT, line);
            } else {
                chunk.emit_op(Op::DYN_GT, line);
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
            chunk.emit_op(Op::STR_COMPARE, line);
            push_const(chunk, Value::F64(0.0), line);
            if mode == 2 {
                chunk.emit_op(Op::DYN_LT, line);
            } else {
                chunk.emit_op(Op::DYN_GT, line);
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
            chunk.emit_op(Op::DYN_LT, line);
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
            chunk.emit_op(Op::DYN_LT, line);
        }
        _ => { chunk.emit_op(Op::FALSE, line); }
    }
    let no_replace = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, inner_slot, line);
    lset(chunk, best_slot, line);
    chunk.patch_jump(no_replace);
    chunk.patch_jump(compared_done);
    chunk.patch_jump(skip_used);

    // inner++
    lget(chunk, inner_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, inner_slot, line);
    chunk.emit_loop(inner_top, line);
    chunk.patch_jump(inner_exit);

    // used[best] = true
    lget(chunk, used_slot, line);
    lget(chunk, best_slot, line);
    chunk.emit_op(Op::TRUE, line);
    chunk.emit_op(Op::ARRAY_SET, line);

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
    chunk.emit_loop(outer_top, line);
    chunk.patch_jump(outer_exit);

    // Delete every original key from obj.
    let i_slot = alloc_local(chunk);
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    let del_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let del_exit = chunk.emit_jump(Op::BR_IF_FALSE, line);
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
    chunk.emit_loop(del_top, line);
    chunk.patch_jump(del_exit);

    // Re-insert in sorted order: obj[sorted_keys[i]] = sorted_vals[i].
    push_const(chunk, Value::F64(0.0), line);
    lset(chunk, i_slot, line);
    let ins_top = chunk.current_offset();
    lget(chunk, i_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let ins_exit = chunk.emit_jump(Op::BR_IF_FALSE, line);
    lget(chunk, obj_slot, line);
    lget(chunk, sorted_keys_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lget(chunk, sorted_vals_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    lget(chunk, i_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, i_slot, line);
    chunk.emit_loop(ins_top, line);
    chunk.patch_jump(ins_exit);

    // PHP sort family returns true.
    chunk.emit_op(Op::TRUE, line);
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
