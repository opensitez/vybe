//! Fortran math helpers — Rust inline opcode emitters.
//!
//! Implements `max(a, b, c, ...)` / `min(a, b, c, ...)` as variadic
//! intrinsics. Composes pure WASM `f64.max` / `f64.min` opcodes —
//! no host calls.

use crate::emitter::instructions::host;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
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
    let idx = chunks[0].add_import(module.to_string(), name.to_string());
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}

fn emit_numeric_zero(chunk: &mut Chunk, line: u32) {
    push_const(chunk, Value::F64(0.0), line);
}

fn emit_numeric_one(chunk: &mut Chunk, line: u32) {
    push_const(chunk, Value::F64(1.0), line);
}

fn emit_numeric_coerce_from_top(chunk: &mut Chunk, line: u32) {
    emit_numeric_zero(chunk, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
}

fn emit_array_get_into_slot(
    chunk: &mut Chunk,
    array_slot: u16,
    index_slot: u16,
    out_slot: u16,
    line: u32,
) {
    lget(chunk, array_slot, line);
    lget(chunk, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, out_slot, line);
}

fn emit_increment_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    emit_numeric_one(chunk, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    lset(chunk, slot, line);
}

pub fn emit_fortran_matmul(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc != 2 {
        for _ in 0..argc {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_op(Op::NULL, line);
        return;
    }

    let right_slot = alloc_local(chunk);
    let left_slot = alloc_local(chunk);
    let result_slot = alloc_local(chunk);
    let left_len_slot = alloc_local(chunk);
    let right_len_slot = alloc_local(chunk);
    let right_first_slot = alloc_local(chunk);
    let right_is_matrix_slot = alloc_local(chunk);
    let col_count_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let row_slot = alloc_local(chunk);
    let row_len_slot = alloc_local(chunk);
    let row_result_slot = alloc_local(chunk);
    let j_slot = alloc_local(chunk);
    let k_slot = alloc_local(chunk);
    let acc_slot = alloc_local(chunk);
    let left_value_slot = alloc_local(chunk);
    let right_value_slot = alloc_local(chunk);
    let right_row_slot = alloc_local(chunk);

    lset(chunk, right_slot, line);
    lset(chunk, left_slot, line);

    lget(chunk, left_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, left_len_slot, line);

    lget(chunk, right_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, right_len_slot, line);

    lget(chunk, left_len_slot, line);
    let _ = chunk;
    call_import(chunks, current, "vybe:js-array", "newWithLength", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, result_slot, line);

    emit_numeric_zero(chunk, line);
    lset(chunk, right_is_matrix_slot, line);
    emit_numeric_zero(chunk, line);
    lset(chunk, col_count_slot, line);

    lget(chunk, right_len_slot, line);
    emit_numeric_zero(chunk, line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    emit_numeric_zero(chunk, line);
    lset(chunk, k_slot, line);
    emit_array_get_into_slot(chunk, right_slot, k_slot, right_first_slot, line);
    lget(chunk, right_first_slot, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    lset(chunk, right_is_matrix_slot, line);

    lget(chunk, right_is_matrix_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    lget(chunk, right_first_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, col_count_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    emit_numeric_zero(chunk, line);
    lset(chunk, i_slot, line);

    // BLOCK+LOOP: br_if 1 exits outer_block (depth 0=LOOP, depth 1=BLOCK)
    let outer_block = chunk.emit_block(line);
    let (outer_loop, _) = chunk.emit_loop_s(line);
    lget(chunk, i_slot, line);
    lget(chunk, left_len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    emit_array_get_into_slot(chunk, left_slot, i_slot, row_slot, line);
    lget(chunk, row_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, row_len_slot, line);

    lget(chunk, right_is_matrix_slot, line);
    crate::emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    lget(chunk, col_count_slot, line);
    let _ = chunk;
    call_import(chunks, current, "vybe:js-array", "newWithLength", 1, line);
    let chunk = &mut chunks[current];
    lset(chunk, row_result_slot, line);
    emit_numeric_zero(chunk, line);
    lset(chunk, j_slot, line);

    // BLOCK+LOOP: br_if 1 exits col_block (depth 0=LOOP, depth 1=BLOCK)
    let col_block = chunk.emit_block(line);
    let (col_loop, _) = chunk.emit_loop_s(line);
    lget(chunk, j_slot, line);
    lget(chunk, col_count_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    emit_numeric_zero(chunk, line);
    lset(chunk, acc_slot, line);
    emit_numeric_zero(chunk, line);
    lset(chunk, k_slot, line);

    // BLOCK+LOOP: br_if 1 exits dot_block (depth 0=LOOP, depth 1=BLOCK)
    let dot_block = chunk.emit_block(line);
    let (dot_loop, _) = chunk.emit_loop_s(line);
    lget(chunk, k_slot, line);
    lget(chunk, row_len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    emit_array_get_into_slot(chunk, row_slot, k_slot, left_value_slot, line);
    emit_array_get_into_slot(chunk, right_slot, k_slot, right_row_slot, line);
    emit_array_get_into_slot(chunk, right_row_slot, j_slot, right_value_slot, line);

    lget(chunk, acc_slot, line);
    lget(chunk, left_value_slot, line);
    emit_numeric_coerce_from_top(chunk, line);
    lget(chunk, right_value_slot, line);
    emit_numeric_coerce_from_top(chunk, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, acc_slot, line);

    emit_increment_slot(chunk, k_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(dot_loop);
    chunk.emit_end(line);
    chunk.patch_block(dot_block);

    lget(chunk, row_result_slot, line);
    lget(chunk, j_slot, line);
    lget(chunk, acc_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);

    emit_increment_slot(chunk, j_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(col_loop);
    chunk.emit_end(line);
    chunk.patch_block(col_block);

    lget(chunk, result_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, row_result_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_else(line);

    emit_numeric_zero(chunk, line);
    lset(chunk, acc_slot, line);
    emit_numeric_zero(chunk, line);
    lset(chunk, k_slot, line);

    let (vec_dot_loop, _) = chunk.emit_loop_s(line);
    lget(chunk, k_slot, line);
    lget(chunk, row_len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    emit_array_get_into_slot(chunk, row_slot, k_slot, left_value_slot, line);
    emit_array_get_into_slot(chunk, right_slot, k_slot, right_value_slot, line);

    lget(chunk, acc_slot, line);
    lget(chunk, left_value_slot, line);
    emit_numeric_coerce_from_top(chunk, line);
    lget(chunk, right_value_slot, line);
    emit_numeric_coerce_from_top(chunk, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, acc_slot, line);

    emit_increment_slot(chunk, k_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(vec_dot_loop);

    lget(chunk, result_slot, line);
    lget(chunk, i_slot, line);
    lget(chunk, acc_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    emit_increment_slot(chunk, i_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(outer_loop);
    chunk.emit_end(line);
    chunk.patch_block(outer_block);

    lget(chunk, result_slot, line);
}

/// Fortran `max(a, b, c, ...)` — variadic.
/// Stack on entry: `[arg0, arg1, ..., argN-1]` (argc args).
/// Stack on exit: `[largest]`.
///
/// Composes: chained `f64.max` (one fewer than argc).
pub fn emit_fortran_max(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_op(Op::NULL, line);
        return;
    }
    for _ in 1..argc {
        chunk.emit_op(Op::F64_MAX, line);
    }
}

/// Fortran `min(a, b, c, ...)` — variadic.
/// Stack on entry: `[arg0, arg1, ..., argN-1]` (argc args).
/// Stack on exit: `[smallest]`.
pub fn emit_fortran_min(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_op(Op::NULL, line);
        return;
    }
    for _ in 1..argc {
        chunk.emit_op(Op::F64_MIN, line);
    }
}
