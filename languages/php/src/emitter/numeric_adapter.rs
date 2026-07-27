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
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_op(Op::NULL, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

/// `__php_float_bytes($v, $bytes)` → a Uint8Array holding the little-endian
/// IEEE-754 encoding of `$v` (`$bytes` = 4 → float32, 8 → float64). Backs PHP
/// `pack('f'|'d', …)`; PHP source can't reach DataView, so this adapter drives
/// ecma:arraybuffer/dataview/uint8array plus a `setFloat32`/`setFloat64` method
/// invoke. The PACK_PRELUDE reads the returned bytes with `chr($u[$j])`.
/// Stack in: `[v, bytes]`; out: `[Uint8Array]`.
pub fn emit_pack_float_bytes(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let v_slot = alloc_local(chunk);
    let buf_slot = alloc_local(chunk);
    let dv_slot = alloc_local(chunk);

    lset(chunk, b_slot, line); // bytes (top)
    lset(chunk, v_slot, line); // value

    // buf = new ArrayBuffer(bytes)
    lget(chunk, b_slot, line);
    let ab_new = chunk.add_import("ecma:arraybuffer", "new");
    chunk.emit_call(ab_new, 1, line);
    lset(chunk, buf_slot, line);

    // dv = new DataView(buf)
    lget(chunk, buf_slot, line);
    let dv_new = chunk.add_import("ecma:dataview", "new");
    chunk.emit_call(dv_new, 1, line);
    lset(chunk, dv_slot, line);

    // invokeMethod(dv, bytes==4 ? "setFloat32" : "setFloat64", 0, v, true)
    lget(chunk, dv_slot, line);
    lget(chunk, b_slot, line);
    push_const(chunk, Value::F64(4.0), line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_str(chunk, "setFloat32", line);
    chunk.emit_else(line);
    push_str(chunk, "setFloat64", line);
    chunk.emit_end(line);
    push_const(chunk, Value::F64(0.0), line); // byteOffset
    lget(chunk, v_slot, line); // value
    push_const(chunk, Value::Bool(true), line); // littleEndian (machine order)
    let invoke = chunk.add_import("ecma:value", "invokeMethod");
    chunk.emit_call(invoke, 5, line);
    chunk.emit_op(Op::DROP, line); // setFloat* returns undefined

    // return new Uint8Array(buf)
    lget(chunk, buf_slot, line);
    let u8_new = chunk.add_import("ecma:uint8array", "new");
    chunk.emit_call(u8_new, 1, line);
}

pub fn emit_php_int_max(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    push_const(&mut chunks[current], Value::bigint_i64(i64::MAX), line);
}

pub fn emit_php_int_min(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    push_const(&mut chunks[current], Value::bigint_i64(i64::MIN), line);
}

pub fn emit_php_is_int(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let v_slot = alloc_local(chunk);
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let number_is_integer = chunk.add_import("ecma:number", "isInteger");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");

    lset(chunk, v_slot, line);

    lget(chunk, v_slot, line);
    chunk.emit_call(test_bigint, 1, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, v_slot, line);
    chunk.emit_call(number_is_integer, 1, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_op(Op::F64_ABS, line);
    push_const(chunk, Value::F64(9_223_372_036_854_774_784.0), line);
    chunk.emit_op(Op::F64_LE, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);

    chunk.emit_end(line);
}

pub fn emit_php_is_float(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // PHP is_float: true for any Value::F64, false for everything else.
    // wasm:js-number.test returns true for F64 values.
    // BigInt is NOT float. I32 is NOT float.
    let chunk = &mut chunks[current];
    let v_slot = alloc_local(chunk);
    let test_number = chunk.add_import("wasm:js-number", "test");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");

    lset(chunk, v_slot, line);

    // BigInt → false
    lget(chunk, v_slot, line);
    chunk.emit_call(test_bigint, 1, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    // Number (F64) → true
    lget(chunk, v_slot, line);
    chunk.emit_call(test_number, 1, line);
    chunk.emit_end(line);
}

pub fn emit_php_abs(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let v_slot = alloc_local(chunk);
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");
    let bigint_lt = chunk.add_import("ecma:bigint", "lt");
    let bigint_neg = chunk.add_import("ecma:bigint", "neg");

    lset(chunk, v_slot, line);

    lget(chunk, v_slot, line);
    chunk.emit_call(test_bigint, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    push_const(chunk, Value::bigint_i64(0), line);
    chunk.emit_call(bigint_lt, 2, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    chunk.emit_call(bigint_neg, 1, line);
    chunk.emit_else(line);
    lget(chunk, v_slot, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    lget(chunk, v_slot, line);
    chunk.emit_op(Op::F64_ABS, line);
    chunk.emit_end(line);
}

pub fn emit_php_intdiv(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let a_slot = alloc_local(chunk);

    lset(chunk, b_slot, line);
    lset(chunk, a_slot, line);

    lget(chunk, b_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::compiler::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::compiler::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    push_str(chunk, "Division by zero", line);
    vybe_compiler::compiler::errors::emit_exception_new_finalize(chunk, "DivisionByZeroError", line);
    vybe_compiler::compiler::errors::emit_throw(chunk, line);
    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_op(Op::F64_DIV, line);
    vybe_compiler::compiler::math::emit_trunc(chunk, line);
    chunk.emit_end(line);
}

fn coerce_to_str(chunk: &mut Chunk, slot: u16, line: u32) {
    push_str(chunk, "", line);
    lget(chunk, slot, line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
}

fn emit_numeric_fallback(
    chunk: &mut Chunk,
    source_slot: u16,
    parse_float: u16,
    plus: bool,
    line: u32,
) {
    lget(chunk, source_slot, line);
    chunk.emit_call(parse_float, 1, line);
    push_const(chunk, Value::F64(1.0), line);
    if plus {
        vybe_compiler::compiler::ops::emit_dyn_add(chunk, line)
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
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lget(chunk, width_slot, line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);
    push_str(chunk, "0", line);
    lget(chunk, out_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
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
        let s = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, s, line);
        s
    };

    let parse_float = chunks[0].add_import("ecma:number", "parseFloat");
    let parse_int = chunks[0].add_import("ecma:number", "parseInt");

    // typeof(v) === "string"?  if not, BR over the string-coerce arm.
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    let test_str_v = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_v, 1, line);
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
    {
        let idx = chunk.add_import("wasm:js-string", "length");
        chunk.emit_call(idx, 1, line);
    }
    lset(chunk, len_slot, line);

    push_const(chunk, Value::F64(-1.0), line);
    lset(chunk, suffix_start_slot, line);

    // Recognise a non-digit prefix followed by a one- or two-digit suffix.
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(47.0), line);
    vybe_compiler::compiler::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(58.0), line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(2.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(47.0), line);
    vybe_compiler::compiler::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(58.0), line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(3.0), line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    lget(chunk, len_slot, line);
    push_const(chunk, Value::F64(3.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, i_slot, line);

    lget(chunk, s_slot, line);
    lget(chunk, i_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "charCodeAt");
        chunk.emit_call(idx, 2, line);
    }
    lset(chunk, code_slot, line);

    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(47.0), line);
    vybe_compiler::compiler::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, code_slot, line);
    push_const(chunk, Value::F64(58.0), line);
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
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
    vybe_compiler::compiler::ops::emit_dyn_lt(chunk, line);
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
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, prefix_slot, line);

    lget(chunk, s_slot, line);
    lget(chunk, suffix_start_slot, line);
    lget(chunk, len_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "substring");
        chunk.emit_call(idx, 3, line);
    }
    lset(chunk, suffix_slot, line);

    lget(chunk, suffix_slot, line);
    push_const(chunk, Value::F64(10.0), line);
    chunk.emit_call(parse_int, 2, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(if plus { Op::F64_ADD } else { Op::F64_SUB }, line);

    push_str(chunk, "", line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    lset(chunk, out_slot, line);
    emit_pad_to_width_from_slots(chunk, out_slot, width_slot, line);

    lget(chunk, prefix_slot, line);
    lget(chunk, out_slot, line);
    {
        let idx = chunk.add_import("wasm:js-string", "concat");
        chunk.emit_call(idx, 2, line);
    }
    chunk.emit_end(line);
    chunk.emit_else(line);

    // Numeric case: v ± 1
    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    if plus {
        vybe_compiler::compiler::ops::emit_dyn_add(chunk, line)
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
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
    lset(chunk, range_slot, line);

    // result = min + floor( (abs(r) / 2^64) * range )
    lget(chunk, min_slot, line);
    chunk.emit_call(rand_idx, 0, line);
    chunk.emit_op(Op::F64_ABS, line);
    push_const(chunk, Value::F64(18446744073709551616.0), line); // 2^64
    chunk.emit_op(Op::F64_DIV, line);
    lget(chunk, range_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    vybe_compiler::compiler::ops::emit_dyn_add(chunk, line);
}

/// PHP `lcg_value()` — a pseudo-random float in the half-open range [0, 1).
/// Composes `wasi:random` (a raw u64) scaled by 2^64, matching the `rand`
/// family's entropy source. No arguments.
pub fn emit_lcg_value(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let rand_idx = chunks[0].add_import(
        "wasi:random/insecure".to_string(),
        "get-insecure-random-u64".to_string(),
    );
    let chunk = &mut chunks[current];
    // abs(r) / 2^64  ∈ [0, 1)
    chunk.emit_call(rand_idx, 0, line);
    chunk.emit_op(Op::F64_ABS, line);
    push_const(chunk, Value::F64(18446744073709551616.0), line); // 2^64
    chunk.emit_op(Op::F64_DIV, line);
}

/// PHP numeric comparison: coerce both sides to f64 via parseFloat,
/// then compare as numbers. Handles `'10' > '9'` → true.
/// Stack: [a, b] → [bool]
pub fn emit_php_compare_gt(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let a_slot = alloc_local(chunk);
    let number_fn = chunk.add_import("ecma:number", "Number");

    lset(chunk, b_slot, line);
    lset(chunk, a_slot, line);
    lget(chunk, a_slot, line);
    chunk.emit_call(number_fn, 1, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(number_fn, 1, line);
    chunk.emit_op(Op::F64_GT, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
}

/// PHP numeric comparison: less-than.
pub fn emit_php_compare_lt(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let a_slot = alloc_local(chunk);
    let number_fn = chunk.add_import("ecma:number", "Number");

    lset(chunk, b_slot, line);
    lset(chunk, a_slot, line);
    lget(chunk, a_slot, line);
    chunk.emit_call(number_fn, 1, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(number_fn, 1, line);
    chunk.emit_op(Op::F64_LT, line);
    vybe_compiler::compiler::ops::emit_i32_to_bool(chunk, line);
}
