use vybe_compiler::primitives::collections::{emit_get, emit_len};
use vybe_runtime::{Chunk, Op};
fn alloc_local(c: &mut Chunk) -> u16 {
    c.alloc_scratch(1)
}
fn lget(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn lset(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
pub fn emit_sequence_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let left_slot = alloc_local(&mut chunks[current]);
    let right_slot = alloc_local(&mut chunks[current]);
    let left_len_slot = alloc_local(&mut chunks[current]);
    let right_len_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let idx_slot = alloc_local(&mut chunks[current]);
    let left_elem_slot = alloc_local(&mut chunks[current]);
    let right_elem_slot = alloc_local(&mut chunks[current]);
    let result_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], right_slot, line);
    lset(&mut chunks[current], left_slot, line);

    lget(&mut chunks[current], left_slot, line);
    emit_len(chunks, current, line);
    lset(&mut chunks[current], left_len_slot, line);

    lget(&mut chunks[current], right_slot, line);
    emit_len(chunks, current, line);
    lset(&mut chunks[current], right_len_slot, line);

    lget(&mut chunks[current], left_len_slot, line);
    lget(&mut chunks[current], right_len_slot, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], left_len_slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], right_len_slot, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], len_slot, line);

    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], idx_slot, line);

    // The private counter and emit_len bounds are numeric. Keep element
    // comparisons dynamic, but omit generic coercion and string-add paths
    // from loop bookkeeping. F64 addition preserves the existing counter type.
    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], left_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    emit_get(chunks, current, line);
    lset(&mut chunks[current], left_elem_slot, line);

    lget(&mut chunks[current], right_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    emit_get(chunks, current, line);
    lset(&mut chunks[current], right_elem_slot, line);

    lget(&mut chunks[current], left_elem_slot, line);
    lget(&mut chunks[current], right_elem_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(-1, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], left_elem_slot, line);
    lget(&mut chunks[current], right_elem_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_br(3, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], idx_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    lget(&mut chunks[current], result_slot, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], left_len_slot, line);
    lget(&mut chunks[current], right_len_slot, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(-1, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], left_len_slot, line);
    lget(&mut chunks[current], right_len_slot, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], result_slot, line);
}
