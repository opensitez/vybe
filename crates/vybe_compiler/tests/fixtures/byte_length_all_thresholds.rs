use vybe_runtime::{Chunk, Op};
fn get(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn set(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
use vybe_compiler::primitives::strings::emit_length;
pub fn emit_byte_length(chunks: &mut [Chunk], current: usize, line: u32) {
    // Length, indices and codePointAt results are already i32. Keep the
    // counter typed instead of emitting dynamic type dispatch per character.
    let base = chunks[current].alloc_scratch(5);
    let (s, i, n, bytes, cp) = (base, base + 1, base + 2, base + 3, base + 4);

    set(&mut chunks[current], s, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(0, line);
    set(&mut chunks[current], bytes, line);
    get(&mut chunks[current], s, line);
    emit_length(&mut chunks[current], line);
    set(&mut chunks[current], n, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    get(&mut chunks[current], i, line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::I32_GE_U, line);
    vybe_compiler::primitives::loops::emit_loop_break_if_from_i32(chunks, current, line);

    // cp = s.codePointAt(i) — the WHOLE code point, not a surrogate half.
    get(&mut chunks[current], s, line);
    get(&mut chunks[current], i, line);
    {
        let idx = chunks[current].add_import("wasm:js-string", "codePointAt");
        chunks[current].emit_call(idx, 2, line);
    }
    set(&mut chunks[current], cp, line);

    // Width is one plus the three crossed UTF-8 thresholds.
    chunks[current].emit_i32_const(1, line);
    for bound in [128, 2048, 65536] {
        get(&mut chunks[current], cp, line);
        chunks[current].emit_i32_const(bound, line);
        chunks[current].emit_op(Op::I32_GE_U, line);
        chunks[current].emit_op(Op::I32_ADD, line);
    }
    get(&mut chunks[current], bytes, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], bytes, line);

    // Each predicate is already 0 or 1, so advance by one plus astral.
    chunks[current].emit_i32_const(1, line);
    get(&mut chunks[current], cp, line);
    chunks[current].emit_i32_const(65535, line);
    chunks[current].emit_op(Op::I32_GT_U, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    get(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], i, line);

    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    get(&mut chunks[current], bytes, line);
}
