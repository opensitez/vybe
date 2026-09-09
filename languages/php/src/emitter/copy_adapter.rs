//! PHP copy/identity adapters.
//!
//! PHP arrays are value-copied on assignment, while objects and callables stay
//! reference-like. The syntax-level walker still marks only the RHS sites that
//! can alias; this adapter provides the PHP leaf behind that marker without a
//! PHP copy-on-assign and strict-equality adapter.

use std::sync::Arc;

use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),
        _ => unreachable!("push_const: unsupported constant"),
    }
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
}

fn call_import(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

fn call_ref(chunk: &mut Chunk, argc: u8, line: u32) {
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, argc, line);
}

fn ref_func(chunk: &mut Chunk, func_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::REF_FUNC, func_idx as u16, line);
    chunk.emit(0, line);
}

fn bump_i32(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, slot, line);
}

fn helper_loop_start(chunk: &mut Chunk, line: u32) -> vybe_compiler::primitives::loops::LoopState {
    let block_patch = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    vybe_compiler::primitives::loops::LoopState {
        block_patch,
        loop_patch,
        body_block_patch: None,
    }
}

fn helper_loop_end(
    chunk: &mut Chunk,
    state: vybe_compiler::primitives::loops::LoopState,
    line: u32,
) {
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(state.loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(state.block_patch);
}

fn emit_string_eq_lit(chunk: &mut Chunk, slot: u16, value: &str, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, value, line);
    call_import(chunk, "wasm:js-string", "equals", 2, line);
}

/// Stack: [] -> [i32 bool]. True for packed arrays and PHP associative arrays;
/// false for class objects/callables/scalars.
fn emit_php_arrayish_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    let type_slot = alloc_local(chunk);

    lget(chunk, slot, line);
    call_import(chunk, "ecma:array", "isArray", 1, line);
    chunk.emit_if_i32(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);

    lget(chunk, slot, line);
    vybe_compiler::primitives::instructions::recipes::is_object(chunk, line);
    chunk.emit_if_i32(line);

    lget(chunk, slot, line);
    push_str(chunk, "__type", line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, type_slot, line);

    lget(chunk, type_slot, line);
    call_import(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_if_i32(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);

    lget(chunk, type_slot, line);
    call_import(chunk, "wasm:js-string", "test", 1, line);
    chunk.emit_if_i32(line);
    emit_string_eq_lit(chunk, type_slot, "Array", line);
    emit_string_eq_lit(chunk, type_slot, "Map", line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
}

fn build_copy_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut c = create_function_chunk("__php_copy_on_assign_impl", 1);
    let value_slot = 0;
    let out_slot = alloc_local(&mut c);
    let entries_slot = alloc_local(&mut c);
    let len_slot = alloc_local(&mut c);
    let i_slot = alloc_local(&mut c);
    let entry_slot = alloc_local(&mut c);
    let key_slot = alloc_local(&mut c);
    let item_slot = alloc_local(&mut c);
    let copied_slot = alloc_local(&mut c);
    let packed_slot = alloc_local(&mut c);

    emit_php_arrayish_slot(&mut c, value_slot, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if_i32(line);
    lget(&mut c, value_slot, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    lget(&mut c, value_slot, line);
    call_import(&mut c, "ecma:array", "isArray", 1, line);
    lset(&mut c, packed_slot, line);
    lget(&mut c, packed_slot, line);
    c.emit_if_i32(line);
    c.emit_array_new_fixed(0, 0, line);
    c.emit_else(line);
    call_import(&mut c, "ecma:map", "new", 0, line);
    c.emit_end(line);
    lset(&mut c, out_slot, line);

    lget(&mut c, value_slot, line);
    call_import(&mut c, "ecma:object", "entries", 1, line);
    lset(&mut c, entries_slot, line);
    lget(&mut c, entries_slot, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut c, len_slot, line);
    c.emit_i32_const(0, line);
    lset(&mut c, i_slot, line);

    let loop_state = helper_loop_start(&mut c, line);
    lget(&mut c, i_slot, line);
    lget(&mut c, len_slot, line);
    c.emit_op(Op::I32_LT_S, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);

    lget(&mut c, entries_slot, line);
    lget(&mut c, i_slot, line);
    c.emit_op(Op::ARRAY_GET, line);
    lset(&mut c, entry_slot, line);
    lget(&mut c, entry_slot, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::ARRAY_GET, line);
    lset(&mut c, key_slot, line);
    lget(&mut c, entry_slot, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::ARRAY_GET, line);
    lset(&mut c, item_slot, line);

    emit_php_arrayish_slot(&mut c, item_slot, line);
    c.emit_if_i32(line);
    ref_func(&mut c, helper_idx, line);
    lget(&mut c, item_slot, line);
    call_ref(&mut c, 1, line);
    c.emit_else(line);
    lget(&mut c, item_slot, line);
    c.emit_end(line);
    lset(&mut c, copied_slot, line);

    lget(&mut c, packed_slot, line);
    c.emit_if_i32(line);
    lget(&mut c, out_slot, line);
    lget(&mut c, copied_slot, line);
    call_import(&mut c, "ecma:array", "push", 2, line);
    c.emit_op(Op::DROP, line);
    c.emit_else(line);
    lget(&mut c, out_slot, line);
    lget(&mut c, key_slot, line);
    lget(&mut c, copied_slot, line);
    call_import(&mut c, "ecma:array", "set", 3, line);
    c.emit_op(Op::DROP, line);
    c.emit_end(line);

    bump_i32(&mut c, i_slot, line);
    helper_loop_end(&mut c, loop_state, line);

    lget(&mut c, out_slot, line);
    c.emit_op(Op::RETURN, line);

    chunks.push(c);
    helper_idx
}

fn build_strict_eq_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut c = create_function_chunk("__php_strict_eq_impl", 2);
    let a_slot = 0;
    let b_slot = 1;
    let a_is_slot = alloc_local(&mut c);
    let b_is_slot = alloc_local(&mut c);
    let entries_a_slot = alloc_local(&mut c);
    let entries_b_slot = alloc_local(&mut c);
    let len_slot = alloc_local(&mut c);
    let i_slot = alloc_local(&mut c);
    let pair_a_slot = alloc_local(&mut c);
    let pair_b_slot = alloc_local(&mut c);
    let result_slot = alloc_local(&mut c);

    emit_php_arrayish_slot(&mut c, a_slot, line);
    lset(&mut c, a_is_slot, line);
    emit_php_arrayish_slot(&mut c, b_slot, line);
    lset(&mut c, b_is_slot, line);

    lget(&mut c, a_is_slot, line);
    lget(&mut c, b_is_slot, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_if_i32(line);

    lget(&mut c, a_is_slot, line);
    lget(&mut c, b_is_slot, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if_i32(line);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    lget(&mut c, a_slot, line);
    call_import(&mut c, "ecma:object", "entries", 1, line);
    lset(&mut c, entries_a_slot, line);
    lget(&mut c, b_slot, line);
    call_import(&mut c, "ecma:object", "entries", 1, line);
    lset(&mut c, entries_b_slot, line);

    lget(&mut c, entries_a_slot, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut c, len_slot, line);
    lget(&mut c, len_slot, line);
    lget(&mut c, entries_b_slot, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    c.emit_op(Op::I32_NE, line);
    c.emit_if_i32(line);
    c.emit_bool_const(false, line);
    c.emit_op(Op::RETURN, line);
    c.emit_end(line);

    c.emit_bool_const(true, line);
    lset(&mut c, result_slot, line);
    c.emit_i32_const(0, line);
    lset(&mut c, i_slot, line);

    let loop_state = helper_loop_start(&mut c, line);
    lget(&mut c, i_slot, line);
    lget(&mut c, len_slot, line);
    c.emit_op(Op::I32_LT_S, line);
    lget(&mut c, result_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
    c.emit_op(Op::I32_AND, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);

    lget(&mut c, entries_a_slot, line);
    lget(&mut c, i_slot, line);
    c.emit_op(Op::ARRAY_GET, line);
    lset(&mut c, pair_a_slot, line);
    lget(&mut c, entries_b_slot, line);
    lget(&mut c, i_slot, line);
    c.emit_op(Op::ARRAY_GET, line);
    lset(&mut c, pair_b_slot, line);

    lget(&mut c, pair_a_slot, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::ARRAY_GET, line);
    lget(&mut c, pair_b_slot, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::ops::emit_js_strict_eq(&mut c, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_if_i32(line);
    c.emit_bool_const(false, line);
    lset(&mut c, result_slot, line);
    c.emit_end(line);

    lget(&mut c, result_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
    c.emit_if_i32(line);
    ref_func(&mut c, helper_idx, line);
    lget(&mut c, pair_a_slot, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::ARRAY_GET, line);
    lget(&mut c, pair_b_slot, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::ARRAY_GET, line);
    call_ref(&mut c, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
    lset(&mut c, result_slot, line);
    c.emit_end(line);

    bump_i32(&mut c, i_slot, line);
    helper_loop_end(&mut c, loop_state, line);

    lget(&mut c, result_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut c, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut c, line);
    c.emit_op(Op::RETURN, line);

    c.emit_else(line);

    lget(&mut c, a_slot, line);
    lget(&mut c, b_slot, line);
    vybe_compiler::primitives::ops::emit_js_strict_eq(&mut c, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut c, line);
    c.emit_op(Op::RETURN, line);

    c.emit_end(line);

    chunks.push(c);
    helper_idx
}

pub fn emit_php_copy_on_assign(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let helper_idx = build_copy_helper(chunks, line);
    let value_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], value_slot, line);
    ref_func(&mut chunks[current], helper_idx, line);
    lget(&mut chunks[current], value_slot, line);
    call_ref(&mut chunks[current], 1, line);
}

pub fn emit_php_strict_eq(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let helper_idx = build_strict_eq_helper(chunks, line);
    let b_slot = alloc_local(&mut chunks[current]);
    let a_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], b_slot, line);
    lset(&mut chunks[current], a_slot, line);
    ref_func(&mut chunks[current], helper_idx, line);
    lget(&mut chunks[current], a_slot, line);
    lget(&mut chunks[current], b_slot, line);
    call_ref(&mut chunks[current], 2, line);
}
