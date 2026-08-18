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

/// Every count here — an extent, a subscript, a linear position — is an INDEX,
/// and the arithmetic on it is exact integer arithmetic. A shape element
/// arrives as whatever the frontend built it from (an `I32` for a literal, an
/// `F64` once anything has computed it), so it is coerced the way ECMA ToInt32
/// coerces: `i32.or` against zero.
fn emit_to_i32(chunk: &mut Chunk, line: u32) {
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
}

/// `ecma:array.length` answers with an `F64` — it is a JS property, and JS
/// numbers are doubles. A length that is about to be counted with, rather than
/// compared, has to be coerced first or the `i32` opcodes read a float.
fn emit_array_len_i32(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_array_len(chunks, current, line);
    emit_to_i32(&mut chunks[current], line);
}

/// `shape[index]`, as an `i32`.
fn emit_shape_extent(chunks: &mut [Chunk], current: usize, shape: u16, index: u16, line: u32) {
    lget(&mut chunks[current], shape, line);
    lget(&mut chunks[current], index, line);
    collections::emit_get(chunks, current, line);
    emit_to_i32(&mut chunks[current], line);
}

/// `RESHAPE(source, shape[, pad])` — the source's elements, rebuilt into an
/// array of the given shape.
///
/// Stack before: `[source, shape]` or `[source, shape, pad]`.
/// Stack after: `[result]`.
///
/// The result is a nest — `shape[0]` entries, each `shape[1]` entries, and so
/// on — because that is how a ranked array is materialized here. What `order`
/// decides is which SUBSCRIPT runs fastest as the linear source is consumed:
/// `ColumnMajor` fills the first subscript fastest (Fortran's element order),
/// `RowMajor` the last (C's). Fortran's `ORDER=` argument is a permutation
/// vector, and the identity and the full reversal are these two.
///
/// `pad` is CYCLED, not repeated once: a shape asking for more elements than
/// the source has takes `pad(1), pad(2), …, pad(1), …` until it is full. A
/// shape asking for fewer truncates. Both are what the standard says, and both
/// fall out of filling exactly `product(shape)` positions.
///
/// The source is flattened by [`emit_flatten_slot`], which walks a nest in
/// storage order rather than element order. For a rank-1 source — what a
/// reshape is nearly always given — the two are the same sequence and the
/// distinction cannot be observed. A RANKED source under a non-default `order`
/// is where they would diverge, and that is the one case this does not yet
/// separate.
pub fn emit_reshape(
    chunks: &mut [Chunk],
    current: usize,
    has_pad: bool,
    order: ArrayTraversalOrder,
    line: u32,
) {
    let base = chunks[current].alloc_scratch(22);
    let source_slot = base;
    let shape_slot = base + 1;
    let pad_slot = base + 2;
    let flat_slot = base + 3;
    let rank_slot = base + 4;
    let total_slot = base + 5;
    let i_slot = base + 6;
    let weights_slot = base + 7;
    let source_len_slot = base + 8;
    let pad_len_slot = base + 9;
    let ordered_slot = base + 10;
    let k_slot = base + 11;
    let linear_slot = base + 12;
    let extent_slot = base + 13;
    let weight_slot = base + 14;
    let stride_slot = base + 15;
    let subscript_slot = base + 16;
    let nest_slot = base + 17;
    let grouped_slot = base + 18;
    let pos_slot = base + 19;
    let level_slot = base + 20;
    let dim_slot = base + 21;

    if has_pad {
        lset(&mut chunks[current], pad_slot, line);
    }
    lset(&mut chunks[current], shape_slot, line);
    lset(&mut chunks[current], source_slot, line);

    emit_flatten_slot(chunks, current, source_slot, flat_slot, order, line);

    lget(&mut chunks[current], shape_slot, line);
    emit_array_len_i32(chunks, current, line);
    lset(&mut chunks[current], rank_slot, line);

    // `weights[j]` is the product of the extents BEFORE dimension `j`, which is
    // the distance between consecutive values of subscript `j` in column-major
    // order. Accumulating it dimension by dimension leaves `total` holding the
    // product of them all — the number of positions to fill.
    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], weights_slot, line);
    i32_const(&mut chunks[current], line, 1);
    lset(&mut chunks[current], total_slot, line);
    i32_const(&mut chunks[current], line, 0);
    lset(&mut chunks[current], i_slot, line);

    let state = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], rank_slot, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], weights_slot, line);
    lget(&mut chunks[current], total_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], total_slot, line);
    emit_shape_extent(chunks, current, shape_slot, i_slot, line);
    chunks[current].emit_op(Op::I32_MUL, line);
    lset(&mut chunks[current], total_slot, line);

    lget(&mut chunks[current], i_slot, line);
    i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    loops::emit_loop_end(chunks, current, state, line);

    // Pad out to `total`, cycling. The loop is bounded by `total` rather than
    // by "until long enough", so an empty `pad` cannot spin: the guard below
    // skips it entirely, and nothing inside can fail to make progress.
    lget(&mut chunks[current], flat_slot, line);
    emit_array_len_i32(chunks, current, line);
    lset(&mut chunks[current], source_len_slot, line);
    if has_pad {
        lget(&mut chunks[current], pad_slot, line);
        emit_array_len_i32(chunks, current, line);
        lset(&mut chunks[current], pad_len_slot, line);

        lget(&mut chunks[current], pad_len_slot, line);
        i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if(line);

        lget(&mut chunks[current], source_len_slot, line);
        lset(&mut chunks[current], i_slot, line);
        let state = loops::emit_loop_start(chunks, current, line);
        lget(&mut chunks[current], i_slot, line);
        lget(&mut chunks[current], total_slot, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        loops::emit_loop_cond(chunks, current, line);

        lget(&mut chunks[current], flat_slot, line);
        lget(&mut chunks[current], pad_slot, line);
        lget(&mut chunks[current], i_slot, line);
        lget(&mut chunks[current], source_len_slot, line);
        chunks[current].emit_op(Op::I32_SUB, line);
        lget(&mut chunks[current], pad_len_slot, line);
        chunks[current].emit_op(Op::I32_REM_S, line);
        collections::emit_get(chunks, current, line);
        collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);

        lget(&mut chunks[current], i_slot, line);
        i32_const(&mut chunks[current], line, 1);
        chunks[current].emit_op(Op::I32_ADD, line);
        lset(&mut chunks[current], i_slot, line);
        loops::emit_loop_end(chunks, current, state, line);
        chunks[current].emit_end(line);
    }
    lget(&mut chunks[current], flat_slot, line);
    i32_const(&mut chunks[current], line, 0);
    lget(&mut chunks[current], total_slot, line);
    collections::emit_slice(chunks, current, line);
    lset(&mut chunks[current], flat_slot, line);

    // `ordered[k]` is the element belonging at the k-th position of the NEST,
    // which the grouping below reads in order. For `RowMajor` that is the
    // source order already. For `ColumnMajor` position `k` decomposes into
    // subscripts, and those subscripts name a different place in the source.
    match order {
        ArrayTraversalOrder::RowMajor => {
            lget(&mut chunks[current], flat_slot, line);
            lset(&mut chunks[current], ordered_slot, line);
        }
        ArrayTraversalOrder::ColumnMajor => {
            collections::emit_array_new(chunks, current, 0, line);
            lset(&mut chunks[current], ordered_slot, line);
            i32_const(&mut chunks[current], line, 0);
            lset(&mut chunks[current], k_slot, line);

            let outer = loops::emit_loop_start(chunks, current, line);
            lget(&mut chunks[current], k_slot, line);
            lget(&mut chunks[current], total_slot, line);
            chunks[current].emit_op(Op::I32_LT_S, line);
            loops::emit_loop_cond(chunks, current, line);

            i32_const(&mut chunks[current], line, 0);
            lset(&mut chunks[current], linear_slot, line);
            i32_const(&mut chunks[current], line, 0);
            lset(&mut chunks[current], i_slot, line);

            let inner = loops::emit_loop_start(chunks, current, line);
            lget(&mut chunks[current], i_slot, line);
            lget(&mut chunks[current], rank_slot, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            loops::emit_loop_cond(chunks, current, line);

            emit_shape_extent(chunks, current, shape_slot, i_slot, line);
            lset(&mut chunks[current], extent_slot, line);
            lget(&mut chunks[current], weights_slot, line);
            lget(&mut chunks[current], i_slot, line);
            collections::emit_get(chunks, current, line);
            emit_to_i32(&mut chunks[current], line);
            lset(&mut chunks[current], weight_slot, line);

            // `total / (weight * extent)` is how far apart consecutive values of
            // subscript `i` sit in the NEST's own row-major numbering, so
            // dividing `k` by it and taking the remainder recovers that
            // subscript without a second decomposition pass.
            lget(&mut chunks[current], total_slot, line);
            lget(&mut chunks[current], weight_slot, line);
            lget(&mut chunks[current], extent_slot, line);
            chunks[current].emit_op(Op::I32_MUL, line);
            chunks[current].emit_op(Op::I32_DIV_S, line);
            lset(&mut chunks[current], stride_slot, line);

            lget(&mut chunks[current], k_slot, line);
            lget(&mut chunks[current], stride_slot, line);
            chunks[current].emit_op(Op::I32_DIV_S, line);
            lget(&mut chunks[current], extent_slot, line);
            chunks[current].emit_op(Op::I32_REM_S, line);
            lset(&mut chunks[current], subscript_slot, line);

            lget(&mut chunks[current], linear_slot, line);
            lget(&mut chunks[current], subscript_slot, line);
            lget(&mut chunks[current], weight_slot, line);
            chunks[current].emit_op(Op::I32_MUL, line);
            chunks[current].emit_op(Op::I32_ADD, line);
            lset(&mut chunks[current], linear_slot, line);

            lget(&mut chunks[current], i_slot, line);
            i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            lset(&mut chunks[current], i_slot, line);
            loops::emit_loop_end(chunks, current, inner, line);

            lget(&mut chunks[current], ordered_slot, line);
            lget(&mut chunks[current], flat_slot, line);
            lget(&mut chunks[current], linear_slot, line);
            collections::emit_get(chunks, current, line);
            collections::emit_push(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);

            lget(&mut chunks[current], k_slot, line);
            i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(Op::I32_ADD, line);
            lset(&mut chunks[current], k_slot, line);
            loops::emit_loop_end(chunks, current, outer, line);
        }
    }

    // Group the flat run into the nest, innermost dimension first: chopping it
    // into runs of `shape[rank-1]` builds the last axis, chopping THAT into
    // runs of `shape[rank-2]` builds the one above it, and so on up to the
    // first — which is left flat, because it is the result itself. A rank-1
    // shape runs this zero times and the flat run IS the answer.
    lget(&mut chunks[current], ordered_slot, line);
    lset(&mut chunks[current], nest_slot, line);
    i32_const(&mut chunks[current], line, 0);
    lset(&mut chunks[current], level_slot, line);

    let levels = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], level_slot, line);
    lget(&mut chunks[current], rank_slot, line);
    i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], rank_slot, line);
    i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    lget(&mut chunks[current], level_slot, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(&mut chunks[current], dim_slot, line);
    emit_shape_extent(chunks, current, shape_slot, dim_slot, line);
    lset(&mut chunks[current], extent_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], grouped_slot, line);
    i32_const(&mut chunks[current], line, 0);
    lset(&mut chunks[current], pos_slot, line);

    let chunking = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], pos_slot, line);
    lget(&mut chunks[current], nest_slot, line);
    emit_array_len(chunks, current, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], grouped_slot, line);
    lget(&mut chunks[current], nest_slot, line);
    lget(&mut chunks[current], pos_slot, line);
    lget(&mut chunks[current], pos_slot, line);
    lget(&mut chunks[current], extent_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    collections::emit_slice(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], pos_slot, line);
    lget(&mut chunks[current], extent_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], pos_slot, line);
    loops::emit_loop_end(chunks, current, chunking, line);

    lget(&mut chunks[current], grouped_slot, line);
    lset(&mut chunks[current], nest_slot, line);
    lget(&mut chunks[current], level_slot, line);
    i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], level_slot, line);
    loops::emit_loop_end(chunks, current, levels, line);

    lget(&mut chunks[current], nest_slot, line);
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
