use vybe_runtime::{Chunk, Op};
pub fn emit_code_is_allowed(chunk: &mut Chunk, code_slot: u16, url: bool, line: u32) {
    // Eight UTF16 cells encode128 ASCII membership bits. The producer's
    // UTF16 unit is unsigned; the guard rejects every nonASCII unit.
    chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunk.emit_i32_const(127, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_if_i32(line);
    const STANDARD: &str = "\0\0\u{8800}\u{23ff}\u{fffe}\u{07ff}\u{fffe}\u{07ff}";
    const URL: &str = "\0\0\u{2000}\u{23ff}\u{fffe}\u{87ff}\u{fffe}\u{07ff}";
    chunk.emit_string_const(if url { URL } else { STANDARD }, line);
    chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    let read = chunk.add_import("wasm:js-string", "charCodeAt");
    chunk.emit_call(read, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunk.emit_i32_const(15, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_SHR_U, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);
}
