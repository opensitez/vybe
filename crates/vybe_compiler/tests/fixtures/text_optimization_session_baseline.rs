// Frozen session-id encoder before paired-byte lookup.
use vybe_runtime::{Chunk, Op};
pub fn emit_new_id(chunks: &mut [Chunk], current: usize, line: u32) {
    let bytes = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);

    chunks[current].emit_i32_const(16, line);
    let get_bytes = chunks[current].add_import("wasi:random/random", "get-random-bytes");
    chunks[current].emit_call(get_bytes, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bytes, line);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let block = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    // Two hex digits per byte, zero-padded — a 31-char id would be a subtly
    // different credential each time a byte happened to be < 16.
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    emit_hex_nibble(chunks, current, bytes, i, true, line);
    emit_hex_nibble(chunks, current, bytes, i, false, line);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// One hex digit of `bytes[i]` — the high nibble when `high`, else the low one.
/// Stack: [] → [string].
fn emit_hex_nibble(
    chunks: &mut [Chunk],
    current: usize,
    bytes: u16,
    i: u16,
    high: bool,
    line: u32,
) {
    let nibble = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    if high {
        chunks[current].emit_i32_const(4, line);
        chunks[current].emit_op(Op::I32_SHR_U, line);
    }
    chunks[current].emit_i32_const(15, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nibble, line);

    // `substring(s, n, n+1)` — the js-string builtins have no `charAt`.
    chunks[current].emit_string_const("0123456789abcdef", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nibble, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nibble, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    chunks[current].emit_call(substring, 3, line);
}
