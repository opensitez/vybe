use vybe_runtime::{Chunk, Op};
pub fn emit_code_is_allowed(chunk: &mut Chunk, code_slot: u16, url: bool, line: u32) {
    // wasm:js-string.charCodeAt produces a UTF16 unit as strict I32.
    // ASCII letter folding preserves high bits, so nonASCII stays rejected.
    chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_i32_const(97, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(25, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunk.emit_i32_const(48, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_i32_const(9, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_op(Op::I32_OR, line);
    for allowed in if url { [45, 95, 61] } else { [43, 47, 61] } {
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        chunk.emit_i32_const(allowed, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_op(Op::I32_OR, line);
    }
}
