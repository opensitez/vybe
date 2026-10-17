use vybe_runtime::{Chunk, Op};
pub fn emit_byte_length(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_with_limit(chunks, current, line, 715_827_882);
}
pub fn emit_fallback(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_with_limit(chunks, current, line, 0);
}
pub fn emit_wrapping_fallback(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_with_seed(chunks, current, line, 0, i32::MAX);
}
pub fn emit_unsigned_wrapping_fallback(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_with_seed(chunks, current, line, 0, -1);
}
fn emit_with_limit(chunks: &mut [Chunk], current: usize, line: u32, limit: i32) {
    emit_with_seed(chunks, current, line, limit, 0);
}
fn emit_with_seed(chunks: &mut [Chunk], current: usize, line: u32, limit: i32, initial: i32) {
    let c = &mut chunks[current];
    let b = c.alloc_scratch(5);
    let (s, i, n, bytes, cp) = (b, b + 1, b + 2, b + 3, b + 4);
    c.emit_op_u16(Op::LOCAL_SET, s, line);
    c.emit_i32_const(0, line);
    c.emit_op_u16(Op::LOCAL_TEE, i, line);
    c.emit_op_u16(Op::LOCAL_SET, bytes, line);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    let length = c.add_import("wasm:js-string", "length");
    c.emit_call(length, 1, line);
    c.emit_op_u16(Op::LOCAL_TEE, n, line);
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_op(Op::STRING_MEASURE_WTF8, line);
    // A negative-one measurement requests the old wrapping counter.
    c.emit_op_u16(Op::LOCAL_TEE, cp, line);
    if limit == 0 {
        c.emit_op(Op::DROP, line);
        c.emit_i32_const(0, line);
    } else {
        c.emit_i32_const(-1, line);
        c.emit_op(Op::NE, line);
    }
    c.emit_if_value(line);
    c.emit_op_u16(Op::LOCAL_GET, cp, line);
    c.emit_else(line);
    // Carry the wrapping counter in the standard multi-value loop parameter.
    c.emit_i32_const(initial, line);
    c.emit_block_params(line, 1, 1);
    c.emit_loop_params(line, 1, 1);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_op_u16(Op::LOCAL_GET, n, line);
    c.emit_op(Op::I32_GE_U, line);
    c.emit_br_if(1, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    // Three bits count the UTF-8 width thresholds after the leading-zero
    // shift. OR 1 makes the zero code point use shift 31 rather than 32.
    c.emit_i32_const(0x0110_8000, line);
    c.emit_op_u16(Op::LOCAL_GET, s, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    let point = c.add_import("wasm:js-string", "codePointAt");
    c.emit_call(point, 2, line);
    c.emit_op_u16(Op::LOCAL_TEE, cp, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_OR, line);
    c.emit_op(Op::I32_CLZ, line);
    c.emit_op(Op::I32_SHR_U, line);
    c.emit_op(Op::I32_POPCNT, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_GET, i, line);
    c.emit_i32_const(1, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_GET, cp, line);
    c.emit_i32_const(65536, line);
    c.emit_op(Op::I32_GE_U, line);
    c.emit_op(Op::I32_ADD, line);
    c.emit_op_u16(Op::LOCAL_SET, i, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_end(line);
    c.emit_else(line);
    c.emit_i32_const(0, line);
    c.emit_end(line);
}
