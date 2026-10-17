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
    let parse_float = chunks[current].add_import("ecma:number", "parseFloat");
    let test_bool = chunks[current].add_import("wasm:js-boolean", "test");
    let boolean_cast = chunks[current].add_import("wasm:js-boolean", "cast");
    let chunk = &mut chunks[current];
    let v_slot = alloc_local(chunk);
    let n_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_TEE, v_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::F64(0.0), line);
    chunk.emit_else(line);

    lget(chunk, v_slot, line);
    chunk.emit_call(test_bool, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    chunk.emit_call(boolean_cast, 1, line);
    // A successful Boolean test guarantees the cast is canonical 0 or 1.
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
    chunk.emit_else(line);

    lget(chunk, v_slot, line);
    chunk.emit_call(parse_float, 1, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, n_slot, line);
    // parseFloat produces a Number: only NaN differs from itself.
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_NE, line);
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
/// invoke. The PHP packing adapter reads the returned bytes with `chr($u[$j])`.
/// Stack in: `[v, bytes]`; out: `[Uint8Array]`.
pub fn emit_pack_float_bytes(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let v_slot = alloc_local(chunk);
    let buf_slot = alloc_local(chunk);

    lset(chunk, b_slot, line); // bytes (top)
    lset(chunk, v_slot, line); // value

    // buf = new ArrayBuffer(bytes)
    lget(chunk, b_slot, line);
    let ab_new = chunk.add_import("ecma:arraybuffer", "new");
    chunk.emit_call(ab_new, 1, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, buf_slot, line);

    // dv = new DataView(buf)
    let dv_new = chunk.add_import("ecma:dataview", "new");
    chunk.emit_call(dv_new, 1, line);

    // invokeMethod(dv, bytes==4 ? "setFloat32" : "setFloat64", 0, v, true)
    // Keep the DataView on the operand stack until the call consumes it.
    lget(chunk, b_slot, line);
    push_const(chunk, Value::F64(4.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);

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

    chunk.emit_op_u16(Op::LOCAL_TEE, v_slot, line);
    chunk.emit_call(test_bigint, 1, line);
    chunk.emit_if_i32(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);

    lget(chunk, v_slot, line);
    chunk.emit_call(number_is_integer, 1, line);
    let boolean_cast = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(boolean_cast, 1, line);
    chunk.emit_if_i32(line);
    lget(chunk, v_slot, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_op(Op::F64_ABS, line);
    push_const(chunk, Value::F64(9_223_372_036_854_774_784.0), line);
    chunk.emit_op(Op::F64_LE, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

pub fn emit_php_is_float(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    // PHP is_float: true for any Value::F64, false for everything else.
    // wasm:js-number.test returns true for F64 values.
    // BigInt is NOT float. I32 is NOT float.
    let chunk = &mut chunks[current];
    let v_slot = alloc_local(chunk);
    let test_number = chunk.add_import("wasm:js-number", "test");
    let test_bigint = chunk.add_import("wasm:js-bigint", "test");

    // Retain the input while testing it, without a store/reload pair.
    chunk.emit_op_u16(Op::LOCAL_TEE, v_slot, line);
    // BigInt → false
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

    chunk.emit_op_u16(Op::LOCAL_TEE, v_slot, line);
    chunk.emit_call(test_bigint, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, v_slot, line);
    // Exact Number zero preserves this guarded BigInt sign comparison.
    push_const(chunk, Value::I32(0), line);
    chunk.emit_call(bigint_lt, 2, line);
    let cast = chunk.add_import("wasm:js-boolean", "cast");
    chunk.emit_call(cast, 1, line);
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

    chunk.emit_if(line);
    super::type_guard::emit_throw_const(
        chunks,
        current,
        "DivisionByZeroError",
        "Division by zero",
        line,
    );
    let chunk = &mut chunks[current];
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

    chunk.emit_op_u16(Op::LOCAL_GET, v_slot, line);
    chunk.emit_call(test_str, 1, line);
    chunk.emit_if_value(line);

    // A numeric string steps as its NUMBER (PHP checks is_numeric first).
    let n_slot = alloc_local(chunk);
    lget(chunk, v_slot, line);
    chunk.emit_call(number, 1, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, n_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
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

/// String protocol Inc. The profile binds ONLY builtin_slots.string.inc here.
/// Its ordinary operator caller has already tested wasm:js-string.test; keep
/// generic php.inc available separately with its original complete semantics.
pub fn emit_php_string_inc_proven(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let helper = php_string_increment_helper(chunks, line);
    let chunk = &mut chunks[current];
    let v_slot = alloc_local(chunk);
    lset(chunk, v_slot, line);
    let number = chunk.add_import("ecma:number", "Number");
    let n_slot = alloc_local(chunk);
    lget(chunk, v_slot, line);
    chunk.emit_call(number, 1, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, n_slot, line);
    lget(chunk, n_slot, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    lget(chunk, n_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    if let Some(index) = helper {
        lget(chunk, v_slot, line);
        let _ = chunk;
        if !vybe_compiler::primitives::functions::try_emit_direct_helper_call(
            chunks, current, index, 1, line,
        ) {
            let chunk = &mut chunks[current];
            chunk.emit_op(Op::DROP, line);
            vybe_compiler::primitives::functions::emit_ref_func(chunk, index, 0, line);
            lget(chunk, v_slot, line);
            vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
        }
    } else {
        emit_zend_string_increment(chunk, v_slot, line);
    }
    chunks[current].emit_end(line);
}

const PHP_STRING_INCREMENT_BODY: &str = "primitive.php.string_increment";
const PHP_STRING_INCREMENT_PROVENANCE: &str = "<php-private-string-increment-body>";

fn php_string_increment_helper(chunks: &mut Vec<Chunk>, line: u32) -> Option<usize> {
    use vybe_runtime::chunk::{LocalName, LocalOrigin};
    if let Some(index) = chunks.iter().position(|c| {
        c.name == PHP_STRING_INCREMENT_BODY && c.arity == 1
            && c.local_names.iter().any(|n| n.slot == 0
                && n.origin == LocalOrigin::Compiler
                && n.name == PHP_STRING_INCREMENT_PROVENANCE)
    }) {
        return (index <= usize::from(u16::MAX)).then_some(index);
    }
    // Existing ordinary ref.func operand is u16. Preserve old inline behavior
    // for modules beyond that boundary; never truncate or alter runtime ABI.
    if chunks.len() > usize::from(u16::MAX) { return None; }
    let index = chunks.len();
    let mut body = vybe_compiler::primitives::functions::create_function_chunk(PHP_STRING_INCREMENT_BODY, 1);
    body.local_count = 1;
    body.param_count = 1;
    body.local_names.push(LocalName::new(0, PHP_STRING_INCREMENT_PROVENANCE, LocalOrigin::Compiler));
    emit_zend_string_increment(&mut body, 0, line);
    body.emit_op(Op::RETURN, line);
    chunks.push(body);
    Some(index)
}

/// Zend `increment_string`, scanning right to left. Stack: `[]` → `[string]`
/// (reads the string from `v_slot`).
fn emit_zend_string_increment(chunk: &mut Chunk, v_slot: u16, line: u32) {
    let char_code_at = chunk.add_import("wasm:js-string", "charCodeAt");
    let from_char_code = chunk.add_import("wasm:js-string", "fromCharCode");
    let substring = chunk.add_import("wasm:js-string", "substring");
    let length = chunk.add_import("wasm:js-string", "length");

    let pos = alloc_local(chunk);
    let tail = alloc_local(chunk);
    let ch = alloc_local(chunk);
    let result = alloc_local(chunk);

    // pos = len - 1; tail accumulates the RESET characters ('z' → 'a', …).
    lget(chunk, v_slot, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
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
    // -1 is the exhausted unsigned index, including an initial empty range.
    // Do not use a signed comparison that would reject a high-bit index.
    chunk.emit_i32_const(-1, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_br_if(1, line);

    lget(chunk, v_slot, line);
    lget(chunk, pos, line);
    chunk.emit_call(char_code_at, 2, line);
    lset(chunk, ch, line);

    // a–y, A–Y, 0–8: increment in place, done.
    let mut first = true;
    for (lo, hi) in [(97, 121), (65, 89), (48, 56)] {
        lget(chunk, ch, line);
        // Unsigned subtraction folds both bounds into one comparison. Codes
        // below lo wrap above the small accepted range instead of matching it.
        chunk.emit_i32_const(lo, line);
        chunk.emit_op(Op::I32_SUB, line);
        chunk.emit_i32_const(hi - lo, line);
        chunk.emit_op(Op::I32_LE_U, line);
        if first {
            first = false;
        } else {
            chunk.emit_op(Op::I32_OR, line);
        }
    }
    chunk.emit_if(line);
    lget(chunk, v_slot, line);
    chunk.emit_i32_const(0, line);
    lget(chunk, pos, line);
    chunk.emit_call(substring, 3, line);
    lget(chunk, ch, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_call(from_char_code, 1, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    lget(chunk, tail, line);
    vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
    lset(chunk, result, line);
    chunk.emit_br(3, line);
    chunk.emit_end(line);

    // z / Z / 9: reset, carry one position left.
    for (code, reset) in [(122, "a"), (90, "A"), (57, "0")] {
        lget(chunk, ch, line);
        chunk.emit_i32_const(code, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if(line);
        push_str(chunk, reset, line);
        lget(chunk, tail, line);
        vybe_compiler::primitives::strings::emit_concat(chunk, 2, line);
        lset(chunk, tail, line);
        lget(chunk, pos, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_SUB, line);
        lset(chunk, pos, line);
        chunk.emit_br(1, line);
        chunk.emit_end(line);
    }

    // Non-alphanumeric: carry STOPS — the resets already made stay.
    lget(chunk, v_slot, line);
    chunk.emit_i32_const(0, line);
    lget(chunk, pos, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
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
    chunk.emit_i32_const(0, line);
    chunk.emit_call(char_code_at, 2, line);
    lset(chunk, ch, line);
    lget(chunk, ch, line);
    chunk.emit_i32_const(57, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    push_str(chunk, "1", line);
    chunk.emit_else(line);
    lget(chunk, ch, line);
    chunk.emit_i32_const(25, line);
    chunk.emit_op(Op::I32_SUB, line);
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
/// Floating-point rounding can turn the largest `u64` into `2^64`, so clamp
/// the scaled offset to `range - 1` before adding the minimum. With no args,
/// PHP `rand()` spans `[0, getrandmax()]`
/// (2147483647). Composes only `wasi:random` + arithmetic opcodes — no new
/// host fns.
pub fn emit_rand(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let rand_idx = chunks[current].add_import(
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
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, range_slot, line);

    // result = min + floor( (unsigned r / 2^64) * range )
    lget(chunk, min_slot, line);
    chunk.emit_call(rand_idx, 0, line);
    chunk.emit_op(Op::F64_CONVERT_I64_U, line);
    push_const(chunk, Value::F64(18446744073709551616.0), line); // 2^64
    chunk.emit_op(Op::F64_DIV, line);
    lget(chunk, range_slot, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    lget(chunk, range_slot, line);
    push_const(chunk, Value::F64(1.0), line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_op(Op::F64_MIN, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
}

/// PHP `lcg_value()` — a pseudo-random float in the half-open range [0, 1).
/// Composes `wasi:random` (a raw u64) scaled by 2^64, matching the `rand`
/// family's entropy source. No arguments.
pub fn emit_lcg_value(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let rand_idx = chunks[current].add_import(
        "wasi:random/insecure".to_string(),
        "get-insecure-random-u64".to_string(),
    );
    let chunk = &mut chunks[current];
    // Convert the full unsigned bit pattern, then clamp the rounded endpoint.
    chunk.emit_call(rand_idx, 0, line);
    chunk.emit_op(Op::F64_CONVERT_I64_U, line);
    push_const(chunk, Value::F64(18446744073709551616.0), line); // 2^64
    chunk.emit_op(Op::F64_DIV, line);
    push_const(chunk, Value::F64(f64::from_bits(0x3fef_ffff_ffff_ffff)), line);
    chunk.emit_op(Op::F64_MIN, line);
}

/// PHP numeric comparison: coerce both sides to f64 via parseFloat,
/// then compare as numbers. Handles `'10' > '9'` → true.
/// Stack: [a, b] → [bool]
pub fn emit_php_compare_gt(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let number_fn = chunk.add_import("ecma:number", "Number");

    lset(chunk, b_slot, line);
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
    let number_fn = chunk.add_import("ecma:number", "Number");

    lset(chunk, b_slot, line);
    chunk.emit_call(number_fn, 1, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(number_fn, 1, line);
    chunk.emit_op(Op::F64_LT, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}
