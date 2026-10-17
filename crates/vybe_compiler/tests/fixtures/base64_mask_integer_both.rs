use vybe_runtime::{Chunk, Op};
pub fn emit_code_is_allowed(chunk: &mut Chunk, code_slot: u16, url: bool, line: u32) {
    // ASCII letters: fold the case bit without discarding high bits.
    chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_i32_const(97, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(25, line);
    chunk.emit_op(Op::I32_LE_U, line);
    // Bits0..31 name ASCII32..63. Digits, padding and alphabet punctuation
    // fit one word. Wasm shifts mask the count; the block check below prevents
    // any other32-unit block from aliasing these bits.
    chunk.emit_i32_const(if url { 0x23ff2000 } else { 0x23ff8800 }, line);
    chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunk.emit_i32_const(5, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_OR, line);
    if url {
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        chunk.emit_i32_const(95, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_op(Op::I32_OR, line);
    }
}
