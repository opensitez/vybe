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
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
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

/// PHP `floatval($v)` — and the `(float)` cast underneath it.
///
/// PHP's string→number rule is "read the longest numeric prefix, otherwise
/// **0**", and `true`/`false`/`null` are 1/0/0. ECMA's `parseFloat` agrees on
/// the prefix (`"12abc"` → 12) but answers `NaN` for everything else, so
/// `floatval('bad')` was NaN and one non-numeric element poisoned a whole
/// `array_sum`.
pub fn emit_php_floatval(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let parse_float = chunks[0].add_import("ecma:number", "parseFloat");
    let test_bool = chunks[0].add_import("wasm:js-boolean", "test");
    let is_nan = chunks[0].add_import("ecma:number", "isNaN");
    let chunk = &mut chunks[current];
    let v_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    lset(chunk, v_slot, line);

    lget(chunk, v_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);

    lget(chunk, v_slot, line);
    chunk.emit_call(test_bool, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    lget(chunk, v_slot, line);
    chunk.emit_call(parse_float, 1, line);
    lset(chunk, n_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_call(is_nan, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);
    lget(chunk, n_slot, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
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
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_op(Op::F64_ABS, line);
    push_const(chunk, Value::F64(9_223_372_036_854_774_784.0), line);
    chunk.emit_op(Op::F64_LE, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
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
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_struct_new(0, 0, line);
    chunk.emit_dup(line);
    push_str(chunk, "Division by zero", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        chunk,
        "DivisionByZeroError",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_op(Op::F64_DIV, line);
    vybe_compiler::primitives::math::emit_trunc(chunk, line);
    chunk.emit_end(line);
}

/// `__php_inc(v)` — the `[builtin_slots.string] inc` target and the shared
/// step's string arm. Stack: `[v]` → `[v incremented]`.
pub fn emit_php_inc(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_unary_arith(chunks, current, /*plus=*/ true, line);
}

/// `__php_dec(v)` — the `dec` twin. Stack: `[v]` → `[v decremented]`.
pub fn emit_php_dec(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_unary_arith(chunks, current, /*plus=*/ false, line);
}

/// PHP `$v++` / `$v--`.
///
/// The STRING rules are Zend's own (`increment_string` in
/// Zend/zend_operators.c): a numeric string steps numerically; a
/// non-numeric string INCREMENTS with Perl-style character carry
/// ("az"++ is "ba", "Zz"++ is "AAa", "a9"++ is "b0", "zz"++ is "aaa") where
/// carry stops at a non-alphanumeric character and a carry past the front
/// prepends by the first character's class — and DECREMENT of a non-numeric
/// string is a NO-OP. Non-strings step through the common primitive, which
/// keeps a BigInt exact.
fn emit_unary_arith(chunks: &mut [Chunk], current: usize, plus: bool, line: u32) {
    let chunk = &mut chunks[current];
    let v_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, v_slot, line);

    let test_str = chunk.add_import("wasm:js-string", "test");
    let number = chunk.add_import("ecma:number", "Number");
    let is_nan = chunk.add_import("ecma:number", "isNaN");

    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    chunk.emit_call(test_str, 1, line);
    chunk.emit_if(line);

    // A numeric string steps as its NUMBER (PHP checks is_numeric first).
    let n_slot = alloc_local(chunk);
    lget(chunk, v_slot, line);
    chunk.emit_call(number, 1, line);
    lset(chunk, n_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_call(is_nan, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(if plus { Op::F64_ADD } else { Op::F64_SUB }, line);
    chunk.emit_else(line);
    if plus {
        emit_zend_string_increment(chunk, v_slot, line);
    } else {
        // Zend: decrementing a non-numeric string changes NOTHING.
        lget(chunk, v_slot, line);
    }
    chunk.emit_end(line);

    chunk.emit_else(line);
    // Non-string: the common step (BigInt stays BigInt, null coerces).
    lget(chunk, v_slot, line);
    vybe_compiler::primitives::bigint::emit_step(chunk, plus, line);
    chunk.emit_end(line);
}

/// Zend `increment_string`, scanning right to left. Stack: `[]` → `[string]`
/// (reads the string from `v_slot`).
fn emit_zend_string_increment(chunk: &mut Chunk, v_slot: u16, line: u32) {
    let char_code_at = chunk.add_import("ecma:string", "charCodeAt");
    let from_char_code = chunk.add_import("ecma:string", "fromCharCode");
    let substring = chunk.add_import("wasm:js-string", "substring");
    let length = chunk.add_import("wasm:js-string", "length");

    let pos = alloc_local(chunk);
    let tail = alloc_local(chunk);
    let ch = alloc_local(chunk);
    let result = alloc_local(chunk);

    // pos = len - 1; tail accumulates the RESET characters ('z' → 'a', …).
    lget(chunk, v_slot, line);
    chunk.emit_call(length, 1, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, pos, line);
    push_str(chunk, "", line);
    lset(chunk, tail, line);

    // outer { prep { loop { … } } prepend-code } — carry off the front
    // branches to `prep`'s end (the prepend arm); a finished result
    // branches to `outer`'s end.
    let outer = chunk.emit_block(line);
    let prep = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);

    lget(chunk, pos, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_br_if(1, line);

    lget(chunk, v_slot, line);
    lget(chunk, pos, line);
    chunk.emit_call(char_code_at, 2, line);
    lset(chunk, ch, line);

    // a–y, A–Y, 0–8: increment in place, done.
    let mut first = true;
    for (lo, hi) in [(97.0, 121.0), (65.0, 89.0), (48.0, 56.0)] {
        lget(chunk, ch, line);
        push_const(chunk, Value::F64(lo), line);
        chunk.emit_op(Op::F64_GE, line);
        lget(chunk, ch, line);
        push_const(chunk, Value::F64(hi), line);
        chunk.emit_op(Op::F64_LE, line);
        chunk.emit_op(Op::I32_AND, line);
        if first {
            first = false;
        } else {
            chunk.emit_op(Op::I32_OR, line);
        }
    }
    chunk.emit_if(line);
    lget(chunk, v_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, pos, line);
    chunk.emit_call(substring, 3, line);
    lget(chunk, ch, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_call(from_char_code, 1, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    lget(chunk, tail, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    lset(chunk, result, line);
    chunk.emit_br(3, line);
    chunk.emit_end(line);

    // z / Z / 9: reset, carry one position left.
    for (code, reset) in [(122.0, "a"), (90.0, "A"), (57.0, "0")] {
        lget(chunk, ch, line);
        push_const(chunk, Value::F64(code), line);
        chunk.emit_op(Op::F64_EQ, line);
        chunk.emit_if(line);
        push_str(chunk, reset, line);
        lget(chunk, tail, line);
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
        lset(chunk, tail, line);
        lget(chunk, pos, line);
        push_const(chunk, Value::F64(1.0), line);
        chunk.emit_op(Op::F64_SUB, line);
        lset(chunk, pos, line);
        chunk.emit_br(1, line);
        chunk.emit_end(line);
    }

    // Non-alphanumeric: carry STOPS — the resets already made stay.
    lget(chunk, v_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    lget(chunk, pos, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_call(substring, 3, line);
    lget(chunk, tail, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    lset(chunk, result, line);
    chunk.emit_br(2, line);

    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(prep);

    // Carry past the front: prepend by the ORIGINAL first character's class
    // ('z…' → "a" + resets, 'Z…' → "A", '9…' → "1"; z→a and Z→A are both
    // −25 in char codes).
    lget(chunk, v_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_call(char_code_at, 2, line);
    lset(chunk, ch, line);
    lget(chunk, ch, line);
    push_const(chunk, Value::F64(57.0), line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);
    push_str(chunk, "1", line);
    chunk.emit_else(line);
    lget(chunk, ch, line);
    push_const(chunk, Value::F64(25.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_call(from_char_code, 1, line);
    chunk.emit_end(line);
    lget(chunk, tail, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    lset(chunk, result, line);

    chunk.emit_end(line);
    chunk.patch_block(outer);

    lget(chunk, result, line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
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
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
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
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
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
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}
