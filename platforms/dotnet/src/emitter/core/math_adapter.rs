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
