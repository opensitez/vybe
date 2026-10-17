use vybe_runtime::{Chunk,Op};
pub fn emit_is_pow2(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

