// Captured i32 short-circuit emitter before typed IF repair.
use vybe_runtime::{Chunk, Op};
pub fn emit_and_i32_start(chunk: &mut Chunk, line: u32) -> usize {
    let block = chunk.emit_block_params(line, 1, 1);
    chunk.emit_dup(line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    block
}

// ── Short-circuit logical OR ────────────────────────────────────────────
//
// Usage:
//   compile_left(chunk);
//   let jump = emit_or_start(chunk);
//   compile_right(chunk);
//   emit_short_circuit_end(chunk, jump);

/// After left operand: if truthy, short-circuit (keep left as result).
/// Stack before: [left]  Stack after: [] (right will be compiled next)
/// ⛔ THE LEFT OPERAND IS ALREADY ON THE STACK, SO THIS BLOCK TAKES IT AS A
/// PARAMETER. `emit_block` declares `(0,0)`: our VM has ONE operand stack and
/// happily reaches under a block boundary, but a WASM block ISOLATES its stack
/// and the `emit_dup` below then has nothing to read — `expected externref but
/// nothing on stack`. Both exits leave exactly one value (the `br_if` carries
/// `left` out; the fall-through drops it and the right operand replaces it), so
/// the signature is `(1, 1)`.
pub fn emit_or_start(chunk: &mut Chunk, line: u32) -> usize {
    let block = chunk.emit_block_params(line, 1, 1);
    chunk.emit_dup(line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line); // discard left, right becomes result
    block
}

/// Static-profile `||` where the left operand is already a raw i32 condition.
pub fn emit_or_i32_start(chunk: &mut Chunk, line: u32) -> usize {
    let block = chunk.emit_block_params(line, 1, 1);
    chunk.emit_dup(line);
    chunk.emit_br_if(0, line);
    chunk.emit_op(Op::DROP, line);
    block
}

/// End a short-circuit AND or OR.
/// Stack: [result_value]
pub fn emit_short_circuit_end(chunk: &mut Chunk, block: usize) {
    chunk.emit_end(0);
    chunk.patch_block(block);
}
