use vybe_compiler::primitives::{
    collections::{emit_get, emit_len},
    strings::emit_length,
};
use vybe_runtime::{Chunk, Op};
fn set(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn get(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn lset(c: &mut Chunk, slot: u16, line: u32) {
    set(c, slot, line);
}
fn lget(c: &mut Chunk, slot: u16, line: u32) {
    get(c, slot, line);
}
fn alloc_local(c: &mut Chunk) -> u16 {
    c.alloc_scratch(1)
}
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

    // bytes += UTF-8 width of cp
    for (bound, width) in [(128, 1), (2048, 2), (65536, 3)] {
        get(&mut chunks[current], cp, line);
        chunks[current].emit_i32_const(bound, line);
        chunks[current].emit_op(Op::I32_LT_U, line);
        chunks[current].emit_if_i32(line);
        chunks[current].emit_i32_const(width, line);
        chunks[current].emit_else(line);
    }
    chunks[current].emit_i32_const(4, line);
    for _ in 0..3 {
        chunks[current].emit_end(line);
    }
    get(&mut chunks[current], bytes, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], bytes, line);

    // i += cp > 0xFFFF ? 2 : 1 — an astral code point spans two UTF-16 units.
    get(&mut chunks[current], cp, line);
    chunks[current].emit_i32_const(65535, line);
    chunks[current].emit_op(Op::I32_GT_U, line);
    chunks[current].emit_if_i32(line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], i, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    set(&mut chunks[current], i, line);

    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
    get(&mut chunks[current], bytes, line);
}
pub fn emit_is_sorted(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let idx_slot = alloc_local(&mut chunks[current]);
    let result_slot = alloc_local(&mut chunks[current]);
    let prev_slot = alloc_local(&mut chunks[current]);
    let curr_slot = alloc_local(&mut chunks[current]);

    chunks[current].emit_op_u16(Op::LOCAL_TEE, arr_slot, line);
    emit_len(chunks, current, line);
    lset(&mut chunks[current], len_slot, line);
    chunks[current].emit_bool_const(true, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_i32_const(1, line);
    lset(&mut chunks[current], idx_slot, line);

    // The private counter and emit_len bounds are numeric. Keep element
    // comparisons dynamic, but omit generic coercion and string-add paths
    // from loop bookkeeping. F64 addition preserves the existing counter type.
    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    lget(&mut chunks[current], idx_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    lget(&mut chunks[current], arr_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    emit_get(chunks, current, line);
    lset(&mut chunks[current], curr_slot, line);

    lget(&mut chunks[current], arr_slot, line);
    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    emit_get(chunks, current, line);
    lset(&mut chunks[current], prev_slot, line);

    lget(&mut chunks[current], curr_slot, line);
    lget(&mut chunks[current], prev_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], idx_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    lget(&mut chunks[current], result_slot, line);
}
