#![allow(dead_code)]
use vybe_compiler::primitives::math::{emit_floor, emit_log, emit_pow};
use vybe_runtime::{Chunk, Op};
fn emit_stirling_gamma(chunk: &mut Chunk, x: u16, line: u32) {
    chunk.emit_f64_const(2.0 * std::f64::consts::PI, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_SQRT, line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(std::f64::consts::E, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    emit_pow(chunk, line);
    chunk.emit_op(Op::F64_MUL, line);

    chunk.emit_f64_const(1.0, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_f64_const(12.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_ADD, line);

    chunk.emit_f64_const(1.0, line);
    chunk.emit_f64_const(288.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_ADD, line);

    chunk.emit_f64_const(139.0, line);
    chunk.emit_f64_const(51840.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_SUB, line);

    chunk.emit_op(Op::F64_MUL, line);
}

/// Gamma function for the exercised libm/Python surface. Stack: `[x]`.
pub fn emit_gamma(chunk: &mut Chunk, line: u32) {
    let x = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);
    let result = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, x, line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(0.5, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(std::f64::consts::PI, line);
    chunk.emit_op(Op::F64_SQRT, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(-0.5, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(-2.0 * std::f64::consts::PI.sqrt(), line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    emit_floor(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_f64_const(2.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    chunk.emit_else(line);
    emit_stirling_gamma(chunk, x, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Natural log of `abs(gamma(x))`. Stack: `[x]`.
pub fn emit_lgamma(chunk: &mut Chunk, line: u32) {
    emit_gamma(chunk, line);
    chunk.emit_op(Op::F64_ABS, line);
    emit_log(chunk, line);
}
