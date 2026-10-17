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
pub fn emit_index_func(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_local(&mut chunks[current]);
    let pred_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let idx_slot = alloc_local(&mut chunks[current]);
    let result_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], pred_slot, line);
    lset(&mut chunks[current], arr_slot, line);

    lget(&mut chunks[current], arr_slot, line);
    emit_len(chunks, current, line);
    lset(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(-1, line);
    lset(&mut chunks[current], result_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], result_slot, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    let __abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    lget(&mut chunks[current], pred_slot, line);
    let __recv = vybe_compiler::primitives::callable::emit_callback_receiver(
        &mut chunks[current],
        __abi,
        line,
    );
    lget(&mut chunks[current], arr_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    emit_get(chunks, current, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(
        &mut chunks[current],
        1 + __recv,
        line,
    );
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], idx_slot, line);
    lset(&mut chunks[current], result_slot, line);
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

    lget(&mut chunks[current], result_slot, line);
}
