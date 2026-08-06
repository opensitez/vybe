//! Kotlin `Double`/`Float` classification predicates.

use vybe_compiler::primitives::ops;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

#[derive(Debug, Clone, Copy)]
pub enum CompareZero {
    Lt,
    Gt,
    Le,
    Ge,
}

/// `println(x)` where `x` is statically a `Double`.
///
/// Kotlin's `Double.toString()` always carries a fraction — `2.0 * 3` prints
/// `6.0` — and an integral f64 renders as `6` through every generic path,
/// because at runtime `6.0` and `6` are the same value. Only the static type
/// distinguishes them, which is why this is reached through a
/// `[builtin_slots.double]` binding rather than by a test on the value.
///
/// Stack: `[value]` → `[]`.
pub fn emit_print_double(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    } else {
        emit_double_to_string(chunks, current, line);
    }
    let log = chunks[current].add_import("web:console", "log");
    chunks[current].emit_call(log, 1, line);
}

/// Kotlin `Double.toString()`. Stack: `[value]` → `[string]`.
///
/// `6.0` → `"6.0"`, `0.125` → `"0.125"`, `-2.0` → `"-2.0"`. Appends `.0` only
/// when the ECMA rendering produced no fraction and no exponent — `1e21`
/// renders as `1e+21` in both languages and must not gain a suffix.
pub fn emit_double_to_string(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let text = chunks[current].alloc_scratch(1);

    // Negative zero: ECMA's ToString renders it "0" (§6.1.6.1.20), Kotlin
    // renders the sign — "-0.0". Detected the only way f64 allows: it
    // compares equal to zero while 1/x is -Infinity. Emitted INLINE, so the
    // shape is one value-if with the ordinary rendering in its else arm.
    let v = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("-0.0", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("0.0", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    let to_str = chunks[current].add_import("ecma:string", "String");
    chunks[current].emit_call(to_str, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    emit_contains_non_integral_mark(chunks, current, text, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const(".0", line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 2, line);

    chunks[current].emit_end(line);
}

/// Push i32 `1` when the rendered number already shows a fraction or exponent
/// (`.`, `e`, `E`) or is not a finite decimal at all (`NaN`, `Infinity`).
fn emit_contains_non_integral_mark(chunks: &mut Vec<Chunk>, current: usize, text: u16, line: u32) {
    let mut first = true;
    for mark in [".", "e", "E", "N", "I"] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
        chunks[current].emit_string_const(mark, line);
        let idx = chunks[current].add_import("ecma:string", "includes");
        chunks[current].emit_call(idx, 2, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        if first {
            first = false;
        } else {
            chunks[current].emit_op(Op::I32_OR, line);
        }
    }
}

/// Kotlin `Double.isInfinite()`. Stack: `[value]` → `[bool]`.
///
/// ECMA has `Number.isFinite` and `Number.isNaN` but nothing for "infinite",
/// which is neither of them: `isInfinite(x)` is `!isFinite(x) && !isNaN(x)`.
/// Composed from the two host predicates rather than adding a host function.
pub fn emit_is_infinite(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    let is_finite = chunks[current].add_import("ecma:number", "isFinite");
    chunks[current].emit_call(is_finite, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    // finite → not infinite
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    // not finite → infinite unless it is NaN
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    let is_nan = chunks[current].add_import("ecma:number", "isNaN");
    chunks[current].emit_call(is_nan, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Kotlin `toInt()` with no radix: a STRING receiver parses strictly and
/// throws NumberFormatException on garbage (`"x".toInt()`), a numeric
/// receiver truncates (`3.9.toInt() == 3`). One spelling, two receivers —
/// decided at runtime.
pub fn emit_to_int_throwing(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, v, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    {
        chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
        crate::emitter::strings::emit_strict_int_or_null(chunks, current, 1, line);
        let parsed = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, parsed, line);
        // `typeof`, not REF_IS_NULL: the probe misreads numbers.
        chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
        chunks[current].emit_call(type_of, 1, line);
        chunks[current].emit_string_const("number", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
        crate::emitter::nullability::emit_exception(
            chunks,
            current,
            1,
            "NumberFormatException",
            line,
        );
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, v, line);
    let trunc = chunks[current].add_import("ecma:math", "trunc");
    chunks[current].emit_call(trunc, 1, line);
    chunks[current].emit_end(line);
}

/// Kotlin `String.toIntOrNull(radix?)`.
///
/// Stack: `[receiver]` or `[receiver, radix]` -> `[number|null]`.
pub fn emit_to_int_or_null(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let radix = if argc >= 2 {
        let slot = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
        Some(slot)
    } else {
        None
    };
    let value = chunk.alloc_scratch(1);
    if argc >= 1 {
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    } else {
        chunk.emit_string_const("", line);
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    if let Some(radix) = radix {
        chunk.emit_op_u16(Op::LOCAL_GET, radix, line);
    }
    let parse = chunk.add_import("ecma:number", "parseInt");
    chunk.emit_call(parse, if radix.is_some() { 2 } else { 1 }, line);
    emit_null_if_nan(chunks, current, line);
}

/// Kotlin `String.toDoubleOrNull()`.
///
/// Stack: `[receiver]` -> `[number|null]`.
pub fn emit_to_double_or_null(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    if argc >= 1 {
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    } else {
        chunk.emit_string_const("", line);
        chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let parse = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(parse, 1, line);
    emit_null_if_nan(chunks, current, line);
}

fn emit_null_if_nan(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let parsed = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parsed, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
    let is_nan = chunks[current].add_import("ecma:number", "isNaN");
    chunks[current].emit_call(is_nan, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
    chunks[current].emit_end(line);
}

/// Kotlin `compareTo` syntax lowering.
///
/// Stack: `[compare_result]` -> `[bool]`.
pub fn emit_compare_zero(chunks: &mut Vec<Chunk>, current: usize, op: CompareZero, line: u32) {
    chunks[current].emit_i32_const(0, line);
    match op {
        CompareZero::Lt => ops::emit_dyn_lt(&mut chunks[current], line),
        CompareZero::Gt => ops::emit_dyn_gt(&mut chunks[current], line),
        CompareZero::Le => ops::emit_dyn_le(&mut chunks[current], line),
        CompareZero::Ge => ops::emit_dyn_ge(&mut chunks[current], line),
    }
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Fixed-width integer casts: `toByte()`/`toShort()` sign-extend from the
/// low bits (255.toByte() is -1, like the JVM); `toUByte()`/`toUShort()`
/// mask unsigned. `math.trunc` alone never wrapped.
pub fn emit_wrap_int(
    chunks: &mut Vec<Chunk>,
    current: usize,
    bits: u8,
    signed: bool,
    line: u32,
) {
    let chunk = &mut chunks[current];
    // `Number` first: `"-128".toByte()` arrives with a STRING receiver, and
    // `toI32` traps on it. Numbers pass through unchanged.
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    let to_i32 = chunk.add_import("wasm:js-number", "toI32");
    chunk.emit_call(to_i32, 1, line);
    if signed {
        let shift = 32 - bits as i32;
        chunk.emit_i32_const(shift, line);
        chunk.emit_op(Op::I32_SHL, line);
        chunk.emit_i32_const(shift, line);
        chunk.emit_op(Op::I32_SHR_S, line);
    } else {
        let mask = if bits >= 32 { -1 } else { (1i64 << bits) as i32 - 1 };
        chunk.emit_i32_const(mask, line);
        chunk.emit_op(Op::I32_AND, line);
    }
}

/// Kotlin `toUInt` — the SAME 32 bits read unsigned: `(-1).toUInt()` is
/// 4294967295. Plain truncation kept the sign, which is only right for
/// values already in range.
///
/// Stack: `[value]` → `[f64]`.
pub fn emit_to_uint32(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let number = chunk.add_import("ecma:number", "Number");
    chunk.emit_call(number, 1, line);
    let to_i32 = chunk.add_import("wasm:js-number", "toI32");
    chunk.emit_call(to_i32, 1, line);
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
}

/// Kotlin `Long.toInt()` — the shared width primitive at 32 bits, handed
/// back as a NUMBER (Kotlin `Int` lives in the number model).
///
/// Stack: `[value]` → `[f64 in i32 range]`.
pub fn emit_long_to_int32(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    vybe_compiler::primitives::bigint::emit_as_int_n_number(&mut chunks[current], 32, line);
}

/// Kotlin `toLong()` — the shared truncate-then-BigInt widening.
///
/// Stack: `[value]` → `[bigint]`.
pub fn emit_to_long(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    vybe_compiler::primitives::bigint::emit_to_bigint_trunc(&mut chunks[current], line);
}

/// The three Kotlin `Long` shifts — the shared width primitive at 64 bits
/// (count masks &63 per JLS §15.19, result wraps, `ushr` reads unsigned).
///
/// Stack: `[value, count]` → `[bigint]`.
pub fn emit_long_shift(
    chunks: &mut Vec<Chunk>,
    current: usize,
    kind: vybe_compiler::primitives::bigint::ShiftKind,
    line: u32,
) {
    vybe_compiler::primitives::bigint::emit_wrapped_shift(&mut chunks[current], 64, kind, line);
}

/// Kotlin `Int / Int` — truncating division that THROWS ArithmeticException
/// on a zero divisor (JLS §15.17.2), where the raw wasm op would trap
/// uncatchably and the float fallback answered `Infinity`.
///
/// Stack: `[a, b]` → `[quotient]`.
pub fn emit_int_div(chunks: &mut Vec<Chunk>, current: usize, _argc: u8, line: u32) {
    let b = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_f64_const(0.0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("/ by zero", line);
    crate::emitter::nullability::emit_exception(chunks, current, 1, "ArithmeticException", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    vybe_compiler::primitives::math::emit_trunc(&mut chunks[current], line);
}
