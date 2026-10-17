use vybe_runtime::{Chunk, opcode::Op};
pub fn emit_before(chunk: &mut Chunk, line: u32) {
    let base = chunk.alloc_scratch(2);
    chunk.emit_op_u16(Op::LOCAL_SET, base + 1, line); // sign
    chunk.emit_op_u16(Op::LOCAL_SET, base, line); // x
    // x == +Inf
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    // sign >= 0
    chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_op(Op::I32_AND, line);
    // x == -Inf
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_f64_const(f64::NEG_INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    // sign <= 0
    chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

