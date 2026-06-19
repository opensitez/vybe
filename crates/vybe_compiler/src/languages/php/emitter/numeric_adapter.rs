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

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

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
    if plus {
        crate::emitter::ops::emit_dyn_add(chunk, line)
    } else {
        chunk.emit_op(Op::F64_SUB, line)
    };
}

fn emit_pad_to_width_from_slots(chunk: &mut Chunk, out_slot: u16, width_slot: u16, line: u32) {
    // block { loop { ... } } — the surrounding block makes `br_if 1` (exit
    // once out.length >= width) a valid WASM label. This helper is called
    // from inside an `if/else`, so a bare loop's `br 1` would otherwise
    // branch to the enclosing `if` and skip the trailing concat.
    let pad_block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    lget(chunk, out_slot, line);
    chunk.emit_op(Op::STR_LENGTH, line);
    lget(chunk, width_slot, line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);
    push_str(chunk, "0", line);
    lget(chunk, out_slot, line);
    chunk.emit_op(Op::STR_CONCAT, line);
    lset(chunk, out_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line); // end pad block
    chunk.patch_block(pad_block);
}

/// `__php_inc(v)` — PHP `$v++` arithmetic.
/// Stack on entry: `[v]` ; Stack on exit: `[v + 1]`.
pub fn emit_php_inc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_unary_arith(chunks, current, /*plus=*/ true, line);
}

/// `__php_dec(v)` — PHP `$v--` arithmetic.
/// Stack on entry: `[v]` ; Stack on exit: `[v - 1]`.
pub fn emit_php_dec(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_unary_arith(chunks, current, /*plus=*/ false, line);
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
    chunk.emit_if(line);

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

    push_const(chunk, Value::F64(-1.0), line);
    lset(chunk, suffix_start_slot, line);

    // Recognise a non-digit prefix followed by a one- or two-digit suffix.
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
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
    chunk.emit_if(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(58.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
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
    chunk.emit_if(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(58.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(3.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
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
    chunk.emit_if(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(58.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, suffix_start_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, suffix_start_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, suffix_start_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, suffix_start_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    lget(chunk, suffix_start_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    crate::emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    emit_numeric_fallback(chunk, s_slot, parse_float, plus, line);
    chunk.emit_else(line);

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
    chunk.emit_end(line);
    chunk.emit_else(line);

    // Numeric case: v ± 1
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    if plus {
        crate::emitter::ops::emit_dyn_add(chunk, line)
    } else {
        chunk.emit_op(Op::F64_SUB, line)
    };
    chunk.emit_end(line);
}

/// PHP `rand([$min, $max])` / `mt_rand([$min, $max])` / `random_int($min, $max)`.
///
/// `wasi:random` returns a raw u64; PHP returns an integer uniformly in
/// `[$min, $max]`. Scale via `$min + floor((r / 2^64) * ($max - $min + 1))`.
/// Because `r / 2^64` lies in `[0, 1)`, the floored product is always in
/// `[0, range - 1]`, so the result never escapes `[$min, $max]` regardless
/// of f64 rounding. With no args, PHP `rand()` spans `[0, getrandmax()]`
/// (2147483647). Composes only `wasi:random` + arithmetic opcodes — no new
/// host fns.
pub fn emit_rand(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let rand_idx = chunks[0].add_import(
        "wasi:random/insecure".to_string(),
        "get-insecure-random-u64".to_string(),
    );
    let chunk = &mut chunks[current];
    let min_slot = alloc_local(chunk);
    let max_slot = alloc_local(chunk);
    let range_slot = alloc_local(chunk);

    if argc >= 2 {
        // Stack: [min, max] — max on top.
        lset(chunk, max_slot, line);
        lset(chunk, min_slot, line);
    } else {
        // No-arg form: drop any stray arg, default to [0, getrandmax()].
        for _ in 0..argc {
            chunk.emit_op(Op::DROP, line);
        }
        push_const(chunk, Value::F64(0.0), line);
        lset(chunk, min_slot, line);
        push_const(chunk, Value::F64(2147483647.0), line);
        lset(chunk, max_slot, line);
    }

    // range = (max - min) + 1
    lget(chunk, max_slot, line);
    lget(chunk, min_slot, line);
    chunk.emit_op(Op::F64_SUB, line);
    push_const(chunk, Value::F64(1.0), line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
    lset(chunk, range_slot, line);

    // result = min + floor( (abs(r) / 2^64) * range )
    lget(chunk, min_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, rand_idx, line);
    chunk.emit(0, line);
    chunk.emit_op(Op::F64_ABS, line);
    push_const(chunk, Value::F64(18446744073709551616.0), line); // 2^64
    chunk.emit_op(Op::F64_DIV, line);
    lget(chunk, range_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    crate::emitter::ops::emit_dyn_add(chunk, line);
}
