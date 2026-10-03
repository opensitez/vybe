//! Original selection sort retained only as a fixed benchmark reference.
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::{Chunk, Op};
fn alloc_local(c: &mut Chunk) -> u16 {
    c.alloc_scratch(1)
}
fn lget(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn lset(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}
pub fn emit_sort_by_key_in_place(chunks: &mut [Chunk], current: usize, line: u32) {
    // ⛔ RESOLVED BEFORE THE BORROW. `key_fn` is the caller's block, so it takes
    // a receiver wherever the region declares one, and the ask needs the whole
    // vector while the body below holds a single chunk.
    let abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    let chunk = &mut chunks[current];
    let key_fn = alloc_local(chunk);
    let arr = alloc_local(chunk);
    let len = alloc_local(chunk);
    let i = alloc_local(chunk);
    let j = alloc_local(chunk);
    let best = alloc_local(chunk);
    let tmp = alloc_local(chunk);

    lset(chunk, key_fn, line);
    lset(chunk, arr, line);

    lget(chunk, arr, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, len, line);

    core_wasm::i32_const(chunk, line, 0);
    lset(chunk, i, line);

    let _ = chunk;
    let outer = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, i, line);
    lget(chunk, len, line);
    chunk.emit_op(Op::I32_LT_S, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, i, line);
    lset(chunk, best, line);
    lget(chunk, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, j, line);

    let _ = chunk;
    let inner = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    let chunk = &mut chunks[current];
    lget(chunk, j, line);
    lget(chunk, len, line);
    chunk.emit_op(Op::I32_LT_S, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
    let chunk = &mut chunks[current];

    lget(chunk, key_fn, line);
    let recv = vybe_compiler::primitives::callable::emit_callback_receiver(chunk, abi, line);
    lget(chunk, arr, line);
    lget(chunk, j, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1 + recv, line);
    lget(chunk, key_fn, line);
    let recv = vybe_compiler::primitives::callable::emit_callback_receiver(chunk, abi, line);
    lget(chunk, arr, line);
    lget(chunk, best, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1 + recv, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, j, line);
    lset(chunk, best, line);
    chunk.emit_end(line);

    lget(chunk, j, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, j, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, inner, line);
    let chunk = &mut chunks[current];

    lget(chunk, best, line);
    lget(chunk, i, line);
    chunk.emit_op(Op::I32_NE, line);
    chunk.emit_if(line);
    lget(chunk, arr, line);
    lget(chunk, i, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    lset(chunk, tmp, line);

    lget(chunk, arr, line);
    lget(chunk, i, line);
    lget(chunk, arr, line);
    lget(chunk, best, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op(Op::ARRAY_SET, line);

    lget(chunk, arr, line);
    lget(chunk, best, line);
    lget(chunk, tmp, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_end(line);

    lget(chunk, i, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, i, line);
    let _ = chunk;
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, outer, line);
    lget(&mut chunks[current], arr, line);
}
