//! The `System.Math` statics that are NOT a one-to-one `ecma:math` call.
//!
//! Everything with a direct ECMA counterpart (`Sin`, `Log`, `Cbrt`, `Hypot`, …)
//! is wired straight to its import in `dispatch`. What lands here is what has
//! no single import behind it: `DivRem` answers two numbers at once, `BigMul`
//! is a widening multiply, and `ILogB`/`ScaleB`/`FusedMultiplyAdd` are IEEE
//! operations ECMA never exposed.

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};

use super::object_fields::field_slot;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn field_set(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

fn emit_throw_argument_out_of_range(
    chunks: &mut [Chunk],
    current: usize,
    message: &str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    crate::emitter::core::exceptions::emit_new_typed(
        chunks,
        current,
        "ArgumentOutOfRangeException",
        ValueSource::ConstStr(message.to_string()),
        line,
    );
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

/// Normalise `MidpointRounding` in `slot` to its ORDINAL.
///
/// ⛔ THE MODE ARRIVES AS ITS .NET NAME. `static_member_constant` publishes
/// `MidpointRounding.AwayFromZero` as the string `"AwayFromZero"` — the
/// spelling, so a lowering can select on it without a second representation —
/// and the tie ladder below compares against 1.0. A string never equalled an
/// ordinal, so EVERY named mode fell through to `F64_NEAREST` and rounded
/// ties-to-even. A caller that already passes a number is left alone.
fn emit_mode_ordinal(chunk: &mut Chunk, slot: u16, line: u32) {
    for (name, ordinal) in [
        ("AwayFromZero", 1.0),
        ("ToZero", 2.0),
        ("ToNegativeInfinity", 3.0),
        ("ToPositiveInfinity", 4.0),
        ("ToEven", 0.0),
    ] {
        get(chunk, slot, line);
        chunk.emit_string_const(name, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        chunk.emit_f64_const(ordinal, line);
        set(chunk, slot, line);
        chunk.emit_end(line);
    }
}

/// `Math.Round(value)` / `(value, digits)` / `(value, digits, mode)`.
///
/// The mode is a `MidpointRounding` ordinal, and the five modes are the five
/// ways .NET breaks a tie. Verified against the .NET SDK over every mode × a
/// midpoint, a non-midpoint and a negative:
///
/// | mode | ordinal | rule |
/// |---|---|---|
/// | `ToEven` | 0 | ties to the even neighbour — the .NET DEFAULT |
/// | `AwayFromZero` | 1 | ties away from zero |
/// | `ToZero` | 2 | truncate |
/// | `ToNegativeInfinity` | 3 | floor |
/// | `ToPositiveInfinity` | 4 | ceil |
///
/// ⛔`ToZero`/`ToNegativeInfinity`/`ToPositiveInfinity` are NOT tie rules — they
/// apply to every value, not just a midpoint, which is why each is its own
/// opcode rather than a nudge before `F64_NEAREST`.
///
/// Scaling by `10^digits` first is what makes one implementation serve all
/// three arities: `Round(v)` is `digits = 0`, and the mode-only overload is the
/// three-argument one with `digits = 0` (measured identical for all 25
/// mode × value pairs), so no caller needs a fourth shape.
pub fn emit_round(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(4);
    let (mode, digits, value, factor) = (scratch, scratch + 1, scratch + 2, scratch + 3);

    if argc >= 3 {
        set(chunk, mode, line);
    } else {
        chunk.emit_f64_const(0.0, line);
        set(chunk, mode, line);
    }
    if argc >= 2 {
        set(chunk, digits, line);
    } else {
        chunk.emit_f64_const(0.0, line);
        set(chunk, digits, line);
    }
    set(chunk, value, line);

    // ⛔ AT TWO ARGUMENTS THE SECOND IS EITHER `digits` OR THE MODE.
    // `Round(2.5, MidpointRounding.AwayFromZero)` is a real .NET overload, and
    // a `MidpointRounding` is a NAME here, not a number — fed to `pow` as a
    // digit count it answered `NaN`. Only a number can be a digit count.
    if argc == 2 {
        get(chunk, digits, line);
        let is_number = chunk.add_import("wasm:js-number", "test");
        chunk.emit_call(is_number, 1, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        get(chunk, digits, line);
        set(chunk, mode, line);
        chunk.emit_f64_const(0.0, line);
        set(chunk, digits, line);
        chunk.emit_end(line);
    }

    emit_mode_ordinal(chunk, mode, line);

    // `digits` is bounded 0..=15 and .NET throws outside it. Guarded only where
    // a caller actually supplied one: the mode-only and no-argument forms
    // synthesize 0, which is always in range.
    if argc >= 2 {
        get(chunk, digits, line);
        chunk.emit_f64_const(0.0, line);
        chunk.emit_op(Op::F64_LT, line);
        get(chunk, digits, line);
        chunk.emit_f64_const(15.0, line);
        chunk.emit_op(Op::F64_GT, line);
        chunk.emit_op(Op::I32_OR, line);
        chunk.emit_if(line);
        emit_throw_argument_out_of_range(
            chunks,
            current,
            "Rounding digits must be between 0 and 15, inclusive. (Parameter 'digits')",
            line,
        );
        chunks[current].emit_end(line);
    }
    let chunk = &mut chunks[current];

    chunk.emit_f64_const(10.0, line);
    get(chunk, digits, line);
    let pow = chunk.add_import("ecma:math", "pow");
    chunk.emit_call(pow, 2, line);
    set(chunk, factor, line);

    // The scaled value, rounded by the mode, then unscaled.
    get(chunk, value, line);
    get(chunk, factor, line);
    chunk.emit_op(Op::F64_MUL, line);
    emit_round_scaled(chunk, mode, line);
    get(chunk, factor, line);
    chunk.emit_op(Op::F64_DIV, line);
}

/// The scaled value is on the stack; leave it rounded per the mode in `mode`.
fn emit_round_scaled(chunk: &mut Chunk, mode: u16, line: u32) {
    let scaled = chunk.alloc_scratch(1);
    set(chunk, scaled, line);

    // An if/else ladder over the four non-default ordinals; `ToEven` (0) and
    // anything unrecognised fall through to `F64_NEAREST`, which IS ties-to-even.
    //
    // `AwayFromZero` has no opcode: it is `trunc(x + copysign(0.5, x))`, which
    // pushes a tie outward in whichever direction the value already points and
    // leaves a non-tie where it was.
    for ordinal in [1.0, 2.0, 3.0, 4.0] {
        get(chunk, mode, line);
        chunk.emit_f64_const(ordinal, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if_value(line);
        get(chunk, scaled, line);
        match ordinal as i32 {
            1 => {
                chunk.emit_f64_const(0.5, line);
                get(chunk, scaled, line);
                chunk.emit_op(Op::F64_COPYSIGN, line);
                chunk.emit_op(Op::F64_ADD, line);
                chunk.emit_op(Op::F64_TRUNC, line);
            }
            2 => chunk.emit_op(Op::F64_TRUNC, line),
            3 => chunk.emit_op(Op::F64_FLOOR, line),
            _ => chunk.emit_op(Op::F64_CEIL, line),
        }
        chunk.emit_else(line);
    }
    get(chunk, scaled, line);
    chunk.emit_op(Op::F64_NEAREST, line);
    for _ in 0..4 {
        chunk.emit_end(line);
    }
}

/// `[a, b] → [{Quotient, Remainder, Item1, Item2}]`.
///
/// ⛔ .NET 7 added this TUPLE overload alongside the older out-param one, and
/// callers reach the pair by either spelling — `t.Quotient` or `t.Item1`. Both
/// names are published because a `ValueTuple` genuinely answers to both, not as
/// an alias for one preferred spelling.
///
/// The quotient TRUNCATES toward zero, which is what integer division means in
/// every .NET language; `Math.Floor` would answer −4 where `DivRem(-17, 5)`
/// owes −3.
pub fn emit_div_rem(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(5);
    let divisor = scratch;
    let dividend = scratch + 1;
    let quotient = scratch + 2;
    let remainder = scratch + 3;
    let pair = scratch + 4;

    set(chunk, divisor, line);
    set(chunk, dividend, line);

    get(chunk, dividend, line);
    get(chunk, divisor, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_TRUNC, line);
    set(chunk, quotient, line);

    get(chunk, dividend, line);
    get(chunk, quotient, line);
    get(chunk, divisor, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    set(chunk, remainder, line);

    class_slots::emit_class_alloc(chunk, line);
    set(chunk, pair, line);
    for (key, slot) in [
        ("Quotient", quotient),
        ("Item1", quotient),
        ("Remainder", remainder),
        ("Item2", remainder),
    ] {
        get(chunk, pair, line);
        get(chunk, slot, line);
        field_set(chunk, key, line);
    }
    get(chunk, pair, line);
}

/// `[a, b] → [a * b]` — `Math.BigMul`, the widening multiply.
///
/// The product of two Int32s needs up to 62 bits, so .NET returns an Int64.
/// This platform's Int64 is an f64, which is exact to 53 — the same ceiling
/// every other Int64 here already lives under, not a new one introduced by
/// this multiply.
pub fn emit_big_mul(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::F64_MUL, line);
}

/// `Math.FusedMultiplyAdd(a, b, c)` — `a * b + c`.
///
/// ⚠ .NET computes this with a SINGLE rounding (that is what "fused" means);
/// this emits two, because wasm has no `f64.fma` and neither does ECMA. The
/// results differ only where the exact product needs more than 53 bits before
/// the add — the same rounding this platform already applies to `a * b + c`
/// written out, so nothing here is worse than the spelling it replaces.
pub fn emit_fused_multiply_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let addend = chunk.alloc_scratch(1);
    set(chunk, addend, line);
    chunk.emit_op(Op::F64_MUL, line);
    get(chunk, addend, line);
    chunk.emit_op(Op::F64_ADD, line);
}

/// `Math.ILogB(x)` — the base-two exponent of `x`, as an Int32.
///
/// ⛔ `floor(log2(|x|))`, not `log2` rounded: `ILogB(8)` is 3 and so is
/// `ILogB(15)`. The three IEEE edge answers are .NET's own — `Int32.MinValue`
/// for zero, `Int32.MaxValue` for NaN and either infinity.
pub fn emit_ilogb(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    set(chunk, value, line);

    get(chunk, value, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(f64::from(i32::MIN), line);
    chunk.emit_else(line);

    // NaN fails every comparison with itself; the infinities are their own
    // absolute value. One arm covers all three.
    get(chunk, value, line);
    get(chunk, value, line);
    chunk.emit_op(Op::F64_NE, line);
    get(chunk, value, line);
    chunk.emit_op(Op::F64_ABS, line);
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(f64::from(i32::MAX), line);
    chunk.emit_else(line);
    get(chunk, value, line);
    chunk.emit_op(Op::F64_ABS, line);
    let log2 = chunk.add_import("ecma:math", "log2");
    chunk.emit_call(log2, 1, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `Math.ScaleB(x, n)` — `x * 2^n`, computed as the exact power of two.
pub fn emit_scaleb(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let exponent = chunk.alloc_scratch(1);
    set(chunk, exponent, line);
    chunk.emit_f64_const(2.0, line);
    get(chunk, exponent, line);
    let pow = chunk.add_import("ecma:math", "pow");
    chunk.emit_call(pow, 2, line);
    chunk.emit_op(Op::F64_MUL, line);
}

// ── System.Decimal's integer-bit surface ───────────────────────────────────
//
// A Decimal is an f64 in this tree, so it carries no stored scale of its own.
// `GetBits` therefore RECOVERS the scale — the fewest decimal places that make
// the value integral — and the pair round-trips exactly for every value an f64
// holds, which is the contract the corpus asserts. The word layout is .NET's:
// `[lo, mid, hi, flags]`, magnitude little-endian across the first three and
// the scale in bits 16..23 of `flags` with the sign in bit 31.

/// The largest scale .NET's Decimal admits.
const DECIMAL_MAX_SCALE: f64 = 28.0;

/// `Decimal.Negate(d)`.
pub fn emit_decimal_negate(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::math::emit_neg(&mut chunks[current], line);
}

/// `Decimal.Compare(a, b)` → −1 / 0 / 1, through the SHARED spaceship the
/// comparer surface already uses. A Decimal-only ordering would be a second
/// answer to a question that already has one.
pub fn emit_decimal_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    let abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    vybe_compiler::primitives::object::emit_compare(&mut chunks[current], abi, line);
}

/// Leave `10^n` on the stack for the value in `slot`.
fn emit_pow10(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_f64_const(10.0, line);
    get(chunk, slot, line);
    let pow = chunk.add_import("ecma:math", "pow");
    chunk.emit_call(pow, 2, line);
}

/// `Decimal.GetBits(d)` → `Integer()` of four words.
pub fn emit_decimal_get_bits(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let scratch = chunk.alloc_scratch(4);
    let (value, scale, mag, word) = (scratch, scratch + 1, scratch + 2, scratch + 3);
    set(chunk, value, line);

    // The scale: multiply by ten until the magnitude is integral, or until
    // .NET's own ceiling is reached.
    get(chunk, value, line);
    chunk.emit_op(Op::F64_ABS, line);
    set(chunk, mag, line);
    chunk.emit_f64_const(0.0, line);
    set(chunk, scale, line);

    chunk.emit_block(line);
    chunk.emit_loop_s(line);
    get(chunk, mag, line);
    get(chunk, mag, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_br_if(1, line);
    get(chunk, scale, line);
    chunk.emit_f64_const(DECIMAL_MAX_SCALE, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_br_if(1, line);
    get(chunk, mag, line);
    chunk.emit_f64_const(10.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    set(chunk, mag, line);
    get(chunk, scale, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    set(chunk, scale, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // Rounding the accumulated product removes the error the repeated ×10
    // introduces — `123.45 * 100` is `12344.999999999998` in f64.
    get(chunk, mag, line);
    chunk.emit_op(Op::F64_NEAREST, line);
    set(chunk, mag, line);

    // lo, mid, hi — the magnitude little-endian in 32-bit words.
    for shift in 0..3u32 {
        get(chunk, mag, line);
        if shift > 0 {
            chunk.emit_f64_const(4294967296f64.powi(shift as i32), line);
            chunk.emit_op(Op::F64_DIV, line);
            chunk.emit_op(Op::F64_FLOOR, line);
        }
        set(chunk, word, line);
        get(chunk, word, line);
        get(chunk, word, line);
        chunk.emit_f64_const(4294967296.0, line);
        chunk.emit_op(Op::F64_DIV, line);
        chunk.emit_op(Op::F64_FLOOR, line);
        chunk.emit_f64_const(4294967296.0, line);
        chunk.emit_op(Op::F64_MUL, line);
        chunk.emit_op(Op::F64_SUB, line);
    }

    // flags — the scale in bits 16..23, the sign in bit 31.
    get(chunk, scale, line);
    chunk.emit_f64_const(65536.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    get(chunk, value, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(2147483648.0, line);
    chunk.emit_else(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_end(line);
    chunk.emit_op(Op::F64_ADD, line);

    chunk.emit_array_new_fixed(0, 4, line);
}

/// `New Decimal(bits)` — the inverse of [`emit_decimal_get_bits`]. An ordinary
/// number argument is the scalar constructor and passes straight through.
pub fn emit_decimal_from_bits(chunks: &mut [Chunk], current: usize, line: u32) {
    let scratch = chunks[current].alloc_scratch(4);
    let (arg, mag, flags, scale) = (scratch, scratch + 1, scratch + 2, scratch + 3);
    set(&mut chunks[current], arg, line);

    get(&mut chunks[current], arg, line);
    let is_number = chunks[current].add_import("wasm:js-number", "test");
    chunks[current].emit_call(is_number, 1, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], arg, line);
    chunks[current].emit_else(line);

    // The magnitude, little-endian across the first three words. A word read
    // back as a negative Int32 is the same 32 bits, so it folds by +2^32.
    chunks[current].emit_f64_const(0.0, line);
    set(&mut chunks[current], mag, line);
    for word in 0..3u32 {
        get(&mut chunks[current], arg, line);
        chunks[current].emit_f64_const(word as f64, line);
        vybe_compiler::primitives::collections::emit_get(chunks, current, line);
        set(&mut chunks[current], flags, line);
        get(&mut chunks[current], flags, line);
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_op(Op::F64_LT, line);
        chunks[current].emit_if_value(line);
        get(&mut chunks[current], flags, line);
        chunks[current].emit_f64_const(4294967296.0, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        chunks[current].emit_else(line);
        get(&mut chunks[current], flags, line);
        chunks[current].emit_end(line);
        chunks[current].emit_f64_const(4294967296f64.powi(word as i32), line);
        chunks[current].emit_op(Op::F64_MUL, line);
        get(&mut chunks[current], mag, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        set(&mut chunks[current], mag, line);
    }

    get(&mut chunks[current], arg, line);
    chunks[current].emit_f64_const(3.0, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    set(&mut chunks[current], flags, line);

    get(&mut chunks[current], flags, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_f64_const(65536.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    set(&mut chunks[current], scale, line);
    get(&mut chunks[current], flags, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    chunks[current].emit_f64_const(65536.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    get(&mut chunks[current], scale, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    set(&mut chunks[current], scale, line);

    get(&mut chunks[current], mag, line);
    emit_pow10(&mut chunks[current], scale, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    set(&mut chunks[current], mag, line);

    // Bit 31 of `flags` is the sign, whichever way the word was read back.
    // ⛔ BOTH ARMS PUSH. The magnitude is parked in its slot first: negating a
    // value the `if` did not push leaves the arms at different heights.
    get(&mut chunks[current], flags, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    get(&mut chunks[current], flags, line);
    chunks[current].emit_f64_const(2147483648.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], mag, line);
    chunks[current].emit_op(Op::F64_NEG, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], mag, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}
