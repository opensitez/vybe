//! Captured reversal emitter before outlining and private cursor specialization.
use vybe_compiler::primitives::collections::{emit_array_new, emit_get, emit_len, emit_push};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::{Chunk, Op};
pub fn emit_reversed(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(3); // seq(base) + result(base+1) + i(base+2)
    let seq = base;
    let result = base + 1;
    let i = base + 2;

    // seq = <input on stack>
    chunks[current].emit_op_u16(Op::LOCAL_SET, seq, line);
    // result = []
    emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    // i = len(seq) - 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, seq, line);
    emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);
    // if !(i >= 0) → break
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);
    // result.push(seq[i])
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, seq, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    emit_get(chunks, current, line);
    emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    // i -= 1
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}
