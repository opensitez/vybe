use vybe_compiler::primitives::collections::{emit_array_new, emit_get, emit_len, emit_push};
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
pub fn emit_compact_adjacent(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_local(&mut chunks[current]);
    let out_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let idx_slot = alloc_local(&mut chunks[current]);
    let value_slot = alloc_local(&mut chunks[current]);
    let prev_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], arr_slot, line);
    emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out_slot, line);

    lget(&mut chunks[current], arr_slot, line);
    emit_len(chunks, current, line);
    lset(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], idx_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], len_slot, line);
    // Private numeric index and length; user-value comparisons stay dynamic.
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], arr_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    emit_get(chunks, current, line);
    lset(&mut chunks[current], value_slot, line);

    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], value_slot, line);
    emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], arr_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    emit_get(chunks, current, line);
    lset(&mut chunks[current], prev_slot, line);

    lget(&mut chunks[current], value_slot, line);
    lget(&mut chunks[current], prev_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    // `emit_dyn_not` coerces on the way in and yields i32, which `emit_if`
    // consumes directly — the ToBoolean ladders on either side of it were both
    // no-ops. See `loops::emit_loop_cond`.
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], value_slot, line);
    emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], idx_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    lget(&mut chunks[current], out_slot, line);
}
