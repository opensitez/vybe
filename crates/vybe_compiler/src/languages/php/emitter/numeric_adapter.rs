//! PHP `$x++` / `$x--` numeric semantics — Rust inline opcode emitter.
//!
//! Mirrors the inline-emit shape used by `datetime_adapter.rs`: each
//! `emit_*(chunks, current, argc, line)` writes WASM opcodes directly
//! into `chunks[current]`. PHP's `++` / `--` are polymorphic — number
//! adds/subtracts 1, string-numeric coerces via `ecma:number.parseFloat`
//! before adding, and non-numeric strings with a trailing digit run bump
//! that suffix in place (`"2026-03-25"++ → "2026-03-26"`). Pure-alpha
//! carry (`"a"++ → "b"`) remains a follow-up.
//!
//! No new host fns; composes only `ecma:number.parseFloat` /
//! `ecma:number.parseInt` plus string opcodes.

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use std::sync::Arc;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count = slot + 1;
    slot
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn coerce_to_str(chunk: &mut Chunk, slot: u16, line: u32) {
    push_str(chunk, "", line);
    lget(chunk, slot, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
}

fn emit_numeric_fallback(
    chunk: &mut Chunk,
    source_slot: u16,
    parse_float: u16,
    plus: bool,
    line: u32,
) {
    lget(chunk, source_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, parse_float, line);
    chunk.emit(1, line);
    push_const(chunk, Value::F64(1.0), line);
    if plus { crate::emitter::ops::emit_dyn_add(chunk, line) } else { chunk.emit_op(Op::F64_SUB, line) };
}

fn emit_pad_to_width_from_slots(chunk: &mut Chunk, out_slot: u16, width_slot: u16, line: u32) {
    let loop_top = chunk.current_offset();
    lget(chunk, out_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lget(chunk, width_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let done = chunk.emit_jump(Op::BR_IF_FALSE, line);
    push_str(chunk, "0", line);
    lget(chunk, out_slot, line);
    chunk.emit_op(Op::STR_CONCAT, line);
    lset(chunk, out_slot, line);
    chunk.emit_loop(loop_top, line);
    chunk.patch_jump(done);
}

/// `__php_inc(v)` — PHP `$v++` arithmetic.
/// Stack on entry: `[v]` ; Stack on exit: `[v + 1]`.
pub fn emit_php_inc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_unary_arith(chunks, current, /*plus=*/true, line);
}

/// `__php_dec(v)` — PHP `$v--` arithmetic.
/// Stack on entry: `[v]` ; Stack on exit: `[v - 1]`.
pub fn emit_php_dec(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_unary_arith(chunks, current, /*plus=*/false, line);
}

fn emit_unary_arith(chunks: &mut [Chunk], current: usize, plus: bool, line: u32) {
    let v_slot = {
        let chunk = &mut chunks[current];
        let s = chunk.local_count;
        chunk.local_count = s + 1;
        chunk.emit_op_u16(Op::LOCAL_SET, s, line);
        chunk.emit_op(Op::DROP, line);
        s
    };

    let parse_float = chunks[0].add_import("ecma:number", "parseFloat");
    let parse_int = chunks[0].add_import("ecma:number", "parseInt");

    // typeof(v) === "string"?  if not, BR over the string-coerce arm.
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    let str_const = chunk.add_constant(Value::String(Arc::from("string")));
    chunk.emit_op_u16(Op::CONST, str_const, line);
    crate::emitter::ops::emit_dyn_eq(chunk, line);
    let not_string = chunk.emit_jump(Op::BR_IF_FALSE, line);

    // String case:
    // - if the string ends with a digit run and has a non-digit prefix,
    //   increment that suffix in place (`2026-03-25` -> `2026-03-26`)
    // - otherwise fall back to the existing parseFloat +/- 1 behavior.
    let s_slot = alloc_local(chunk);
    let len_slot = alloc_local(chunk);
    let i_slot = alloc_local(chunk);
    let suffix_start_slot = alloc_local(chunk);
    let code_slot = alloc_local(chunk);
    let width_slot = alloc_local(chunk);
    let prefix_slot = alloc_local(chunk);
    let suffix_slot = alloc_local(chunk);
    let out_slot = alloc_local(chunk);

    coerce_to_str(chunk, v_slot, line);
    lset(chunk, s_slot, line);

    lget(chunk, s_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lset(chunk, len_slot, line);

    // Recognise the common PHP shape of a non-digit prefix followed by
    // a one- or two-digit suffix. This covers the date-like strings
    // that currently hang inside loops (`2026-03-25`, `foo9`, `foo09`)
    // without relying on an inline scan loop.
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let use_numeric_short = chunk.emit_jump(Op::BR_IF_TRUE, line);

    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_CODE_AT, line);
    lset(chunk, code_slot, line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(47.0), line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    let use_numeric_last_lo = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(58.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let use_numeric_last_hi = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_CODE_AT, line);
    lset(chunk, code_slot, line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(47.0), line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    let one_digit_suffix = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(58.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let one_digit_suffix_hi = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(3.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let use_numeric_all_digits = chunk.emit_jump(Op::BR_IF_TRUE, line);

    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(3.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    chunk.emit_op(Op::STR_CHAR_CODE_AT, line);
    lset(chunk, code_slot, line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(47.0), line);
    crate::emitter::ops::emit_dyn_gt(chunk, line);
    let two_digit_suffix = chunk.emit_jump(Op::BR_IF_FALSE, line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(58.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    let two_digit_suffix_hi = chunk.emit_jump(Op::BR_IF_FALSE, line);

    let use_numeric_long_suffix = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(one_digit_suffix);
    chunk.patch_jump(one_digit_suffix_hi);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, suffix_start_slot, line);
    let suffix_ready = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(two_digit_suffix);
    chunk.patch_jump(two_digit_suffix_hi);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, suffix_start_slot, line);
    let two_digit_ready = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(use_numeric_short);
    chunk.patch_jump(use_numeric_last_lo);
    chunk.patch_jump(use_numeric_last_hi);
    chunk.patch_jump(use_numeric_all_digits);
    chunk.patch_jump(use_numeric_long_suffix);
    emit_numeric_fallback(chunk, s_slot, parse_float, plus, line);
    let string_done_numeric = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(suffix_ready);
    chunk.patch_jump(two_digit_ready);

    lget(chunk, len_slot, line);
    lget(chunk, suffix_start_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, width_slot, line);

    lget(chunk, s_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, suffix_start_slot, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    lset(chunk, prefix_slot, line);

    lget(chunk, s_slot, line);
    lget(chunk, suffix_start_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::STR_SUBSTRING, line);
    lset(chunk, suffix_slot, line);

    lget(chunk, suffix_slot, line);
    push_const(chunk, Value::F64(10.0), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, parse_int, line);
    chunk.emit(2, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(if plus { Op::F64_ADD } else { Op::F64_SUB }, line);

    push_str(chunk, "", line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    emit_pad_to_width_from_slots(chunk, out_slot, width_slot, line);

    lget(chunk, prefix_slot, line);
    lget(chunk, out_slot, line);
    chunk.emit_op(Op::STR_CONCAT, line);
    let string_done_suffix = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(not_string);

    // Numeric case: v ± 1
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    if plus { crate::emitter::ops::emit_dyn_add(chunk, line) } else { chunk.emit_op(Op::F64_SUB, line) };

    chunk.patch_jump(string_done_numeric);
    chunk.patch_jump(string_done_suffix);
}
