//! PHP `ctype_*` predicates — Rust inline opcode emitters.
//!
//! Each `emit_ctype_*` writes a char-iteration loop that returns
//! `false` on empty input, otherwise returns `true` iff every char
//! lies in at least one of the predicate's UTF-16 code-unit ranges.
//! Composes only WASM string ops (`STR_LENGTH`, `STR_CODE_POINT_AT`)
//! + numeric comparisons. No host fns; no JS polyfills.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

/// One `lo..=hi` UTF-16 code-unit range.
#[derive(Copy, Clone)]
struct Range {
    lo: u32,
    hi: u32,
}

const fn r(lo: u32, hi: u32) -> Range {
    Range { lo, hi }
}

/// Stack on entry: `[s]` ; Stack on exit: `[bool]`.
fn emit_check(chunks: &mut [Chunk], current: usize, ranges: &[Range], line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = chunk.alloc_scratch(1);
    let i_slot = chunk.alloc_scratch(1);
    let len_slot = chunk.alloc_scratch(1);
    let code_slot = chunk.alloc_scratch(1);
    let matched_slot = chunk.alloc_scratch(1);
    let result_slot = chunk.alloc_scratch(1);

    // s = "" + s  (string coerce)
    let v_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, v_slot, line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);

    // len = s.length
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    let outer = chunk.emit_block(line);

    // if len === 0: result = false; break.
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_f64_const(0.0, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_br(1, line);
    chunk.emit_end(line);

    // i = 0
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunk.emit_bool_const(true, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    // while i < len: per-char range check
    let (loop_patch, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);

    // code = s.codePointAt(i)
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "codePointAt");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, code_slot, line);

    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::LOCAL_SET, matched_slot, line);

    // For each range: if lo<=code<=hi, remember that this char matched.
    for rng in ranges {
        // code >= lo  ≡  !(code < lo)
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        chunk.emit_f64_const(rng.lo as f64, line);
        crate::emitter::ops::emit_dyn_lt(chunk, line);
        crate::emitter::ops::emit_dyn_not(chunk, line);
        chunk.emit_if(line);
        // code <= hi  ≡  !(code > hi)
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        chunk.emit_f64_const(rng.hi as f64, line);
        crate::emitter::ops::emit_dyn_gt(chunk, line);
        crate::emitter::ops::emit_dyn_not(chunk, line);
        chunk.emit_if(line);
        chunk.emit_bool_const(true, line);
        chunk.emit_op_u16(Op::LOCAL_SET, matched_slot, line);
        chunk.emit_end(line);
        chunk.emit_end(line);
    }
    // No range matched → result = false; exit loop.
    // Inside this `if`, label depths are 0=if, 1=loop, 2=outer block, so the
    // loop-exiting break is `br 2` (to `outer`). `br 1` would target the loop
    // itself and restart it → infinite loop on the first non-matching char.
    chunk.emit_op_u16(Op::LOCAL_GET, matched_slot, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_br(2, line);
    chunk.emit_end(line);

    // i++
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);

    chunk.emit_end(line);
    chunk.patch_block(outer);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
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
        chunks,
        current,
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
