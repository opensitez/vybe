//! PHP `ctype_*` predicates — Rust inline opcode emitters.
//!
//! Each `emit_ctype_*` writes a char-iteration loop that returns
//! `false` on empty input, otherwise returns `true` iff every char
//! lies in at least one of the predicate's UTF-16 code-unit ranges.
//! Composes only WASM string ops (`STR_LENGTH`, `STR_CODE_POINT_AT`)
//! + numeric comparisons. No host fns; no JS polyfills.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

/// One `lo..=hi` UTF-16 code-unit range.
#[derive(Copy, Clone)]
struct Range {
    lo: u32,
    hi: u32,
}

const fn r(lo: u32, hi: u32) -> Range { Range { lo, hi } }

/// Stack on entry: `[s]` ; Stack on exit: `[bool]`.
fn emit_check(chunks: &mut [Chunk], current: usize, ranges: &[Range], line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = chunk.local_count;
    let i_slot = s_slot + 1;
    let len_slot = s_slot + 2;
    let code_slot = s_slot + 3;
    chunk.local_count = s_slot + 4;

    // s = "" + s  (string coerce)
    let v_slot = chunk.local_count;
    chunk.local_count = v_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, v_slot, line);
    chunk.emit_op(Op::DROP, line);
    let empty = chunk.add_constant(Value::String(Arc::from("")));
    chunk.emit_op_u16(Op::CONST, empty, line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunk.emit_op(Op::DROP, line);

    // len = s.length
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunk.emit_op(Op::DROP, line);

    // if len === 0: push false and BR to end.
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    let zero = chunk.add_constant(Value::F64(0.0));
    chunk.emit_op_u16(Op::CONST, zero, line);
    chunk.emit_op(Op::DYN_EQ, line);
    let nonempty = chunk.emit_jump(Op::BR_IF_FALSE, line);
    chunk.emit_op(Op::FALSE, line);
    let done_empty = chunk.emit_jump(Op::BR, line);
    chunk.patch_jump(nonempty);

    // i = 0
    let zero2 = chunk.add_constant(Value::F64(0.0));
    chunk.emit_op_u16(Op::CONST, zero2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_op(Op::DROP, line);

    // while i < len: per-char range check
    let loop_top = chunk.current_offset();
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    let exit = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // code = s.codePointAt(i)
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::STR_CODE_POINT_AT, line);
    chunk.emit_op_u16(Op::LOCAL_SET, code_slot, line);
    chunk.emit_op(Op::DROP, line);

    // For each range: if lo<=code<=hi, jump to "matched". After all
    // ranges fail, push false and BR to end.
    let mut accept_jumps: Vec<usize> = Vec::new();
    for rng in ranges {
        // code >= lo  ≡  !(code < lo)
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        let lo_const = chunk.add_constant(Value::F64(rng.lo as f64));
        chunk.emit_op_u16(Op::CONST, lo_const, line);
        chunk.emit_op(Op::DYN_LT, line);
        chunk.emit_op(Op::DYN_NOT, line);
        let skip_lo = chunk.emit_jump(Op::BR_IF_FALSE, line);
        // code <= hi  ≡  !(code > hi)
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        let hi_const = chunk.add_constant(Value::F64(rng.hi as f64));
        chunk.emit_op_u16(Op::CONST, hi_const, line);
        chunk.emit_op(Op::DYN_GT, line);
        chunk.emit_op(Op::DYN_NOT, line);
        let skip_hi = chunk.emit_jump(Op::BR_IF_FALSE, line);
        accept_jumps.push(chunk.emit_jump(Op::BR, line));
        chunk.patch_jump(skip_hi);
        chunk.patch_jump(skip_lo);
    }
    // No range matched → false.
    chunk.emit_op(Op::FALSE, line);
    let done_failed = chunk.emit_jump(Op::BR, line);
    for j in accept_jumps { chunk.patch_jump(j); }

    // i++
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    let one = chunk.add_constant(Value::F64(1.0));
    chunk.emit_op_u16(Op::CONST, one, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(exit);

    // Loop ran to completion → all chars matched → true.
    chunk.emit_op(Op::TRUE, line);

    // Land here from empty-string and from "no range matched" paths.
    chunk.patch_jump(done_empty);
    chunk.patch_jump(done_failed);
}

pub fn emit_ctype_alpha(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(chunks, current, &[r(65, 90), r(97, 122)], line);
}
pub fn emit_ctype_digit(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(chunks, current, &[r(48, 57)], line);
}
pub fn emit_ctype_alnum(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(chunks, current, &[r(48, 57), r(65, 90), r(97, 122)], line);
}
pub fn emit_ctype_space(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(chunks, current, &[r(9, 13), r(32, 32)], line);
}
pub fn emit_ctype_upper(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(chunks, current, &[r(65, 90)], line);
}
pub fn emit_ctype_lower(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(chunks, current, &[r(97, 122)], line);
}
pub fn emit_ctype_xdigit(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(chunks, current, &[r(48, 57), r(65, 70), r(97, 102)], line);
}
pub fn emit_ctype_punct(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(
        chunks, current,
        &[r(33, 47), r(58, 64), r(91, 96), r(123, 126)],
        line,
    );
}
pub fn emit_ctype_print(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(chunks, current, &[r(32, 126)], line);
}
pub fn emit_ctype_cntrl(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_check(chunks, current, &[r(0, 31), r(127, 127)], line);
}
