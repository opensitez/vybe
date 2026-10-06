#![allow(dead_code)]
use vybe_compiler::primitives::math::emit_exp;
use vybe_runtime::{Chunk, Op};
pub fn emit_erf(chunk: &mut Chunk, line: u32) {
    let x = chunk.alloc_scratch(1);
    let t = chunk.alloc_scratch(1);
    let poly = chunk.alloc_scratch(1);
    let y = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_SET, x, line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(f64::NEG_INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(-1.0, line);
    chunk.emit_else(line);

    chunk.emit_f64_const(1.0, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_ABS, line);
    chunk.emit_f64_const(0.3275911, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op_u16(Op::LOCAL_SET, t, line);

    chunk.emit_f64_const(1.061405429, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_f64_const(-1.453152027, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_f64_const(1.421413741, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_f64_const(-0.284496736, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_f64_const(0.254829592, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, poly, line);

    chunk.emit_f64_const(1.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, poly, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_NEG, line);
    emit_exp(chunk, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, y, line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, y, line);
    chunk.emit_op(Op::F64_NEG, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, y, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Complementary error function. Stack: `[x]` -> `[1 - erf(x)]`.
pub fn emit_erfc(chunk: &mut Chunk, line: u32) {
    emit_erf(chunk, line);
    let erf = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, erf, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, erf, line);
    chunk.emit_op(Op::F64_SUB, line);
}
