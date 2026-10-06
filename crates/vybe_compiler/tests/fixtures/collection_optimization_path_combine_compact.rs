//! Artifact-only compact Path.Combine: private GC buffer and loop kernels.
use vybe_compiler::primitives::{ops, strings};
use vybe_runtime::{Chunk, Op};
#[path = "collection_optimization_path_combine_before.rs"]
mod baseline;
pub const LONG_LOCAL_BUDGET: u16 = 15;
fn get(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_GET, s, l);
}
fn set(c: &mut Chunk, s: u16, l: u32) {
    c.emit_op_u16(Op::LOCAL_SET, s, l);
}
fn int(c: &mut Chunk, n: i32, l: u32) {
    c.emit_i32_const(n, l);
}
fn init(c: &mut Chunk, slot: u16, n: i32, l: u32) {
    int(c, n, l);
    set(c, slot, l);
}
fn step(c: &mut Chunk, slot: u16, n: i32, l: u32) {
    get(c, slot, l);
    int(c, n, l);
    c.emit_op(Op::I32_ADD, l);
    set(c, slot, l);
}
fn read(c: &mut Chunk, buf: u16, idx: u16, dst: u16, l: u32) {
    get(c, buf, l);
    get(c, idx, l);
    c.emit_op(Op::ARRAY_GET, l);
    set(c, dst, l);
}
fn write(c: &mut Chunk, buf: u16, idx: u16, src: u16, l: u32) {
    get(c, buf, l);
    get(c, idx, l);
    get(c, src, l);
    c.emit_op(Op::ARRAY_SET, l);
}
fn positive(c: &mut Chunk, seg: u16, l: u32) {
    get(c, seg, l);
    strings::emit_length(c, l);
    c.emit_f64_const(0.0, l);
    c.emit_op(Op::F64_GT, l);
}
fn trailing(c: &mut Chunk, seg: u16, length: Option<u16>, l: u32) {
    get(c, seg, l);
    c.emit_string_const("/", l);
    strings::emit_last_index_of(c, l);
    c.emit_f64_const(1.0, l);
    c.emit_op(Op::F64_ADD, l);
    if let Some(length) = length {
        get(c, length, l);
    } else {
        get(c, seg, l);
        strings::emit_length(c, l);
    }
    c.emit_op(Op::F64_EQ, l);
}
fn finish_loop(c: &mut Chunk, done: usize, again: usize, l: u32) {
    c.emit_br(0, l);
    c.emit_end(l);
    c.patch_loop(again);
    c.emit_end(l);
    c.patch_block(done);
}
pub fn emit_combine(c: &mut Chunk, argc: u8, l: u32) {
    if argc < 16 {
        baseline::emit_combine(c, argc, l);
        return;
    }
    let before = c.local_count;
    let end = before
        .checked_add(LONG_LOCAL_BUDGET)
        .expect("path local slot overflow");
    let base = c.alloc_scratch(12);
    let buf = base;
    let idx = base + 1;
    let seg = base + 2;
    let total = base + 3;
    let full = base + 4;
    let tail = base + 5;
    let first = base + 6;
    let safe = base + 7;
    let acc = base + 8;
    let count = base + 9;
    let out = base + 10;
    let right = base + 11;
    // Existing GC value-array type zero; no escape or callback route. Keep all
    // evaluated original arguments intact until the fast path is selected.
    c.emit_array_new_fixed(0, argc as u16, l);
    set(c, buf, l);
    init(c, idx, 0, l);
    init(c, tail, 0, l);
    init(c, first, 0, l);
    init(c, safe, 1, l);
    c.emit_f64_const(0.0, l);
    set(c, total, l);
    let done = c.emit_block(l);
    let (again, _) = c.emit_loop_s(l);
    get(c, idx, l);
    int(c, argc as i32, l);
    c.emit_op(Op::I32_GE_S, l);
    c.emit_br_if(1, l);
    read(c, buf, idx, seg, l);
    positive(c, seg, l);
    c.emit_if(l); // Exact strict String gate first.
    get(c, seg, l);
    let length = c.add_import("ecma:string", "length");
    c.emit_call(length, 1, l);
    set(c, full, l);
    init(c, safe, 0, l);
    get(c, full, l);
    c.emit_f64_const(i32::MAX as f64, l);
    c.emit_op(Op::F64_GT, l);
    c.emit_br_if(2, l);
    init(c, safe, 1, l);
    get(c, seg, l);
    baseline::emit_is_rooted(c, l);
    ops::emit_dyn_to_bool(c, l);
    get(c, total, l);
    c.emit_f64_const(0.0, l);
    c.emit_op(Op::F64_EQ, l);
    c.emit_op(Op::I32_OR, l);
    c.emit_if_value(l);
    get(c, idx, l);
    set(c, first, l);
    get(c, full, l);
    c.emit_else(l);
    get(c, total, l);
    get(c, tail, l);
    c.emit_if_value(l);
    c.emit_f64_const(0.0, l);
    c.emit_else(l);
    c.emit_f64_const(1.0, l);
    c.emit_end(l);
    c.emit_op(Op::F64_ADD, l);
    get(c, full, l);
    c.emit_op(Op::F64_ADD, l);
    c.emit_end(l);
    set(c, total, l);
    init(c, safe, 0, l);
    get(c, total, l);
    c.emit_f64_const(i32::MAX as f64, l);
    c.emit_op(Op::F64_GT, l);
    c.emit_br_if(2, l);
    init(c, safe, 1, l);
    trailing(c, seg, Some(full), l);
    set(c, tail, l);
    c.emit_end(l);
    step(c, idx, 1, l);
    finish_loop(c, done, again, l);
    get(c, safe, l);
    c.emit_if_value(l);
    // Validation complete: reuse the private buffer slots as String pieces.
    get(c, first, l);
    set(c, idx, l);
    get(c, first, l);
    set(c, out, l);
    init(c, total, 0, l);
    init(c, tail, 0, l);
    let done = c.emit_block(l);
    let (again, _) = c.emit_loop_s(l);
    get(c, idx, l);
    int(c, argc as i32, l);
    c.emit_op(Op::I32_GE_S, l);
    c.emit_br_if(1, l);
    read(c, buf, idx, seg, l);
    positive(c, seg, l);
    c.emit_if(l);
    get(c, total, l);
    c.emit_if(l);
    get(c, tail, l);
    c.emit_if_value(l);
    c.emit_string_const("", l);
    c.emit_else(l);
    c.emit_string_const("/", l);
    c.emit_end(l);
    get(c, seg, l);
    strings::emit_concat(c, 2, l);
    set(c, seg, l);
    c.emit_end(l);
    trailing(c, seg, None, l);
    set(c, tail, l);
    init(c, total, 1, l);
    // Keep ONLY positive pieces. The output cursor cannot overtake the read
    // cursor, and skipping empty strings preserves the single-input Arc.
    write(c, buf, out, seg, l);
    step(c, out, 1, l);
    c.emit_end(l);
    step(c, idx, 1, l);
    finish_loop(c, done, again, l);
    // Merge adjacent pieces in-place, one balanced level per outer iteration.
    get(c, out, l);
    get(c, first, l);
    c.emit_op(Op::I32_SUB, l);
    set(c, count, l);
    let levels = c.emit_block(l);
    let (level_again, _) = c.emit_loop_s(l);
    get(c, count, l);
    int(c, 1, l);
    c.emit_op(Op::I32_LE_S, l);
    c.emit_br_if(1, l);
    get(c, first, l);
    set(c, idx, l);
    get(c, first, l);
    set(c, out, l);
    let pairs = c.emit_block(l);
    let (pair_again, _) = c.emit_loop_s(l);
    get(c, idx, l);
    get(c, first, l);
    get(c, count, l);
    c.emit_op(Op::I32_ADD, l);
    c.emit_op(Op::I32_GE_S, l);
    c.emit_br_if(1, l);
    read(c, buf, idx, acc, l);
    get(c, idx, l);
    int(c, 1, l);
    c.emit_op(Op::I32_ADD, l);
    set(c, right, l);
    get(c, right, l);
    get(c, first, l);
    get(c, count, l);
    c.emit_op(Op::I32_ADD, l);
    c.emit_op(Op::I32_LT_S, l);
    c.emit_if(l);
    get(c, acc, l);
    get(c, buf, l);
    get(c, right, l);
    c.emit_op(Op::ARRAY_GET, l);
    strings::emit_concat(c, 2, l);
    set(c, acc, l);
    c.emit_end(l);
    write(c, buf, out, acc, l);
    step(c, out, 1, l);
    step(c, idx, 2, l);
    finish_loop(c, pairs, pair_again, l);
    get(c, out, l);
    get(c, first, l);
    c.emit_op(Op::I32_SUB, l);
    set(c, count, l);
    finish_loop(c, levels, level_again, l);
    get(c, count, l);
    c.emit_if_value(l);
    get(c, buf, l);
    get(c, first, l);
    c.emit_op(Op::ARRAY_GET, l);
    c.emit_else(l);
    // Same imported empty constant as the frozen accumulator's initial value.
    c.emit_string_const("", l);
    c.emit_end(l);
    c.emit_else(l);
    // Rare signed-length-wrap fallback: original string-operation order over
    // the untouched evaluated argument buffer, now a compact runtime loop.
    init(c, idx, 0, l);
    c.emit_string_const("", l);
    set(c, acc, l);
    let done = c.emit_block(l);
    let (again, _) = c.emit_loop_s(l);
    get(c, idx, l);
    int(c, argc as i32, l);
    c.emit_op(Op::I32_GE_S, l);
    c.emit_br_if(1, l);
    read(c, buf, idx, seg, l);
    positive(c, seg, l);
    c.emit_if(l);
    get(c, seg, l);
    baseline::emit_is_rooted(c, l);
    ops::emit_dyn_to_bool(c, l);
    get(c, acc, l);
    strings::emit_length(c, l);
    c.emit_f64_const(0.0, l);
    c.emit_op(Op::F64_EQ, l);
    c.emit_op(Op::I32_OR, l);
    c.emit_if_value(l);
    get(c, seg, l);
    c.emit_else(l);
    get(c, acc, l);
    trailing(c, acc, None, l);
    c.emit_if_value(l);
    c.emit_string_const("", l);
    c.emit_else(l);
    c.emit_string_const("/", l);
    c.emit_end(l);
    get(c, seg, l);
    strings::emit_concat(c, 3, l);
    c.emit_end(l);
    set(c, acc, l);
    c.emit_end(l);
    step(c, idx, 1, l);
    finish_loop(c, done, again, l);
    get(c, acc, l);
    c.emit_end(l);
    assert_eq!(c.local_count, end, "complete compact body budget");
}
