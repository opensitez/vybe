//! Shared ranked-array transforms.
//!
//! These are language-neutral array operations such as Fortran `PACK` and
//! `UNPACK`: they reshape/compact/scatter array values. They are intentionally
//! separate from `packing.rs`, which is byte/binary struct packing.

use crate::primitives::{collections, loops, ops};
use vybe_ast::ArrayTraversalOrder;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn i32_const(chunk: &mut Chunk, line: u32, value: i32) {
    chunk.emit_i32_const(value, line);
}

fn emit_array_len(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:array", "length");
    chunks[current].emit_call(idx, 1, line);
}

fn emit_array_flat(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:array", "flat");
    chunks[current].emit_call(idx, 2, line);
}

fn emit_is_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(idx, 1, line);
}

fn emit_flatten_slot(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    result_slot: u16,
    order: ArrayTraversalOrder,
    line: u32,
) {
    // Fortran currently materializes column-major arrays as nested arrays whose
    // ECMA flat order is already logical element order. Keep the order in the
    // AST contract now so future lowerings can distinguish row/column storage
    // without changing frontends again.
    let _ = order;

    lget(&mut chunks[current], value_slot, line);
    emit_is_array(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    lget(&mut chunks[current], value_slot, line);
    i32_const(&mut chunks[current], line, 1024);
    emit_array_flat(chunks, current, line);

    chunks[current].emit_else(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_array_new_fixed(0, 1, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], result_slot, line);
}

fn emit_flat_result_shaped_like_mask(
    chunks: &mut [Chunk],
    current: usize,
    mask_slot: u16,
    flat_slot: u16,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(6);
    let shaped_slot = base;
    let i_slot = base + 1;
    let pos_slot = base + 2;
    let row_slot = base + 3;
    let row_len_slot = base + 4;
    let next_pos_slot = base + 5;

    lget(&mut chunks[current], mask_slot, line);
    emit_is_array(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    lget(&mut chunks[current], mask_slot, line);
    i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    emit_is_array(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], shaped_slot, line);
    i32_const(&mut chunks[current], line, 0);
    lset(&mut chunks[current], i_slot, line);
    i32_const(&mut chunks[current], line, 0);
    lset(&mut chunks[current], pos_slot, line);

    let state = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], mask_slot, line);
    emit_array_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], mask_slot, line);
    lget(&mut chunks[current], i_slot, line);
    collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], row_slot, line);

    lget(&mut chunks[current], row_slot, line);
    emit_array_len(chunks, current, line);
    lset(&mut chunks[current], row_len_slot, line);

    lget(&mut chunks[current], pos_slot, line);
    lget(&mut chunks[current], row_len_slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    lset(&mut chunks[current], next_pos_slot, line);

    lget(&mut chunks[current], shaped_slot, line);
    lget(&mut chunks[current], flat_slot, line);
    lget(&mut chunks[current], pos_slot, line);
    lget(&mut chunks[current], next_pos_slot, line);
    collections::emit_slice(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], next_pos_slot, line);
    lset(&mut chunks[current], pos_slot, line);
    lget(&mut chunks[current], i_slot, line);
    i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    loops::emit_loop_end(chunks, current, state, line);

    lget(&mut chunks[current], shaped_slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], flat_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    lget(&mut chunks[current], flat_slot, line);
    chunks[current].emit_end(line);
}

/// `PACK(array, mask[, vector])`.
///
/// Stack before: `[array, mask]` or `[array, mask, vector]`.
/// Stack after: `[result]`.
pub fn emit_pack_mask(
    chunks: &mut [Chunk],
    current: usize,
    has_vector: bool,
    order: ArrayTraversalOrder,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(12);
    let source_slot = base;
    let mask_slot = base + 1;
    let vector_slot = base + 2;
    let source_flat_slot = base + 3;
    let mask_flat_slot = base + 4;
    let vector_flat_slot = base + 5;
    let result_slot = base + 6;
    let i_slot = base + 7;
    let selected_count_slot = base + 8;
    let vector_len_slot = base + 9;
    let prefix_slot = base + 10;
    let pad_slot = base + 11;

    if has_vector {
        lset(&mut chunks[current], vector_slot, line);
    }
    lset(&mut chunks[current], mask_slot, line);
    lset(&mut chunks[current], source_slot, line);

    emit_flatten_slot(chunks, current, source_slot, source_flat_slot, order, line);
    emit_flatten_slot(chunks, current, mask_slot, mask_flat_slot, order, line);
    if has_vector {
        emit_flatten_slot(chunks, current, vector_slot, vector_flat_slot, order, line);
    }

    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], result_slot, line);
    i32_const(&mut chunks[current], line, 0);
    lset(&mut chunks[current], i_slot, line);

    let state = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], source_flat_slot, line);
    emit_array_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], mask_flat_slot, line);
    lget(&mut chunks[current], i_slot, line);
    collections::emit_get(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], result_slot, line);
    lget(&mut chunks[current], source_flat_slot, line);
    lget(&mut chunks[current], i_slot, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], i_slot, line);
    i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    loops::emit_loop_end(chunks, current, state, line);

    if has_vector {
        lget(&mut chunks[current], result_slot, line);
        emit_array_len(chunks, current, line);
        lset(&mut chunks[current], selected_count_slot, line);

        lget(&mut chunks[current], vector_flat_slot, line);
        emit_array_len(chunks, current, line);
        lset(&mut chunks[current], vector_len_slot, line);

        lget(&mut chunks[current], result_slot, line);
        i32_const(&mut chunks[current], line, 0);
        lget(&mut chunks[current], vector_len_slot, line);
        collections::emit_slice(chunks, current, line);
        lset(&mut chunks[current], prefix_slot, line);

        lget(&mut chunks[current], vector_flat_slot, line);
        lget(&mut chunks[current], selected_count_slot, line);
        lget(&mut chunks[current], vector_len_slot, line);
        collections::emit_slice(chunks, current, line);
        lset(&mut chunks[current], pad_slot, line);

        lget(&mut chunks[current], prefix_slot, line);
        lget(&mut chunks[current], pad_slot, line);
        collections::emit_concat(chunks, current, line);
    } else {
        lget(&mut chunks[current], result_slot, line);
    }
}

/// `UNPACK(vector, mask, field)`.
///
/// Stack before: `[vector, mask, field]`.
/// Stack after: `[result]`.
pub fn emit_unpack_mask(
    chunks: &mut [Chunk],
    current: usize,
    order: ArrayTraversalOrder,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(10);
    let vector_slot = base;
    let mask_slot = base + 1;
    let field_slot = base + 2;
    let vector_flat_slot = base + 3;
    let mask_flat_slot = base + 4;
    let field_flat_slot = base + 5;
    let field_is_array_slot = base + 6;
    let result_slot = base + 7;
    let i_slot = base + 8;
    let vector_index_slot = base + 9;

    lset(&mut chunks[current], field_slot, line);
    lset(&mut chunks[current], mask_slot, line);
    lset(&mut chunks[current], vector_slot, line);

    emit_flatten_slot(chunks, current, vector_slot, vector_flat_slot, order, line);
    emit_flatten_slot(chunks, current, mask_slot, mask_flat_slot, order, line);

    lget(&mut chunks[current], field_slot, line);
    emit_is_array(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    lset(&mut chunks[current], field_is_array_slot, line);

    lget(&mut chunks[current], field_is_array_slot, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], field_slot, line);
    i32_const(&mut chunks[current], line, 1024);
    emit_array_flat(chunks, current, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], field_slot, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], field_flat_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], result_slot, line);
    i32_const(&mut chunks[current], line, 0);
    lset(&mut chunks[current], i_slot, line);
    i32_const(&mut chunks[current], line, 0);
    lset(&mut chunks[current], vector_index_slot, line);

    let state = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], mask_flat_slot, line);
    emit_array_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], mask_flat_slot, line);
    lget(&mut chunks[current], i_slot, line);
    collections::emit_get(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], result_slot, line);
    lget(&mut chunks[current], vector_flat_slot, line);
    lget(&mut chunks[current], vector_index_slot, line);
    collections::emit_get(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], vector_index_slot, line);
    i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], vector_index_slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], result_slot, line);
    lget(&mut chunks[current], field_is_array_slot, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], field_flat_slot, line);
    lget(&mut chunks[current], i_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], field_slot, line);
    chunks[current].emit_end(line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], i_slot, line);
    i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    loops::emit_loop_end(chunks, current, state, line);

    emit_flat_result_shaped_like_mask(chunks, current, mask_slot, result_slot, line);
}
