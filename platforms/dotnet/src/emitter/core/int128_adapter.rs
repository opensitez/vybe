//! `Int128` / `UInt128` width discipline — the four primitives the synthesized
//! classes need and cannot express in C# source.
//!
//! .NET's 128-bit integers are FIXED-WIDTH two's complement, verified on the
//! SDK: `1 << 128` is `1` (the count masks by 127), `Int128.MaxValue + 1` is
//! `Int128.MinValue`, and `UInt128.MaxValue + 1` is `0`. ECMA gives exactly
//! those semantics through `BigInt.asIntN`/`asUintN` (§21.2.2), and
//! `primitives::bigint` already parameterises them by width — so this file is
//! four thin entry points, not an arithmetic implementation.
//!
//! ⚠ WIDENING, NOT TRUNCATION, IS THE POINT OF THE `BigInt` CALL. A 128-bit
//! literal does NOT survive f64: `Int128.MinValue` is -2^127, which as a
//! `Number` is an approximation. Every value enters through `BigInt(v)`, which
//! is exact for a decimal STRING, so the constants are parsed from their
//! decimal spelling rather than computed in the number model.

use vybe_compiler::primitives::bigint::{self, ShiftKind};
use vybe_compiler::primitives::instructions::host;
use vybe_runtime::Chunk;

const WIDTH: u32 = 128;

/// `BigInt(v)` — exact for a string, a bigint, or an integral number.
fn to_bigint(chunk: &mut Chunk, line: u32) {
    let ctor = chunk.add_import("ecma:bigint", "BigInt");
    chunk.emit_call(ctor, 1, line);
}

/// `asIntN(128, BigInt(v))` — the signed wrap every `Int128` result takes.
pub fn emit_wrap_signed(chunks: &mut [Chunk], current: usize, line: u32) {
    to_bigint(&mut chunks[current], line);
    bigint::emit_as_int_n(&mut chunks[current], WIDTH, line);
}

/// `asUintN(128, BigInt(v))` — the unsigned wrap, for `UInt128`.
pub fn emit_wrap_unsigned(chunks: &mut [Chunk], current: usize, line: u32) {
    to_bigint(&mut chunks[current], line);
    bigint::emit_as_uint_n(&mut chunks[current], WIDTH, line);
}

/// `v << n` and `v >> n` with .NET's masking. Stack `[value, count]`.
///
/// ⛔ NOT the language's own `<<`. A plain shift on a BigInt answered `0` in
/// this runtime (`BigInteger.One << 100`), and .NET additionally masks the
/// count by 127 — `1 << 128` is `1`, not `0`. `emit_wrapped_shift` does both,
/// and it is the same helper Kotlin's `Long` uses at width 64.
pub fn emit_shift(chunks: &mut [Chunk], current: usize, left: bool, line: u32) {
    bigint::emit_wrapped_shift(
        &mut chunks[current],
        WIDTH,
        if left { ShiftKind::Shl } else { ShiftKind::Shr },
        line,
    );
}

/// One ECMA BigInt binary operation — `add`, `lt`, `xor`, … Stack `[a, b]`.
///
/// ⛔ THE LANGUAGE'S OWN `+` CANNOT DO THIS. A `+` between two BigInt payloads
/// compiled to a numeric add and trapped with `wasm:js-number.toF64 — not a
/// number`; C# has no BigInt arithmetic path of its own. §6.1.6.2's operations
/// are host functions (`platforms/ecma/src/bigint.rs`) and they are EXACT at
/// any magnitude — which also removes the need to compare by subtraction:
/// `ecma:bigint.lt` compares the mathematical values, so it is right for two
/// large operands where the generic `<` is not.
pub fn emit_binop(chunks: &mut [Chunk], current: usize, op: &'static str, line: u32) {
    host::emit(&mut chunks[current], "ecma:bigint", op, 2, line);
}

/// A small BigInt constant — `0n` / `2n`. Argc 0.
///
/// Needed because an integer LITERAL in the synthesized body is an f64, and
/// every ECMA BigInt operation refuses a Number operand implicitly. Emitted as
/// `BigInt(<n>)` rather than carried as a literal so there is one conversion
/// rule in this file.
pub fn emit_const(chunks: &mut [Chunk], current: usize, n: i32, line: u32) {
    chunks[current].emit_i32_const(n, line);
    to_bigint(&mut chunks[current], line);
}

// ── Composite operations ─────────────────────────────────────────────────
//
// ⛔ THESE EXIST BECAUSE **TWO** PROFILE-BUILTIN CALLS IN ONE SYNTHESIZED
// METHOD BODY OVERFLOW THE COMPILER'S STACK. Measured by bisection: a body of
// `return __vybe_bi_eq(v, 0n);` compiles, and
// `var a = __vybe_bi_eq(v, 0n); return __vybe_bi_eq(v, 0n);` — two calls to the
// SAME builtin, no nesting — aborts the process with
// `fatal runtime error: stack overflow` at COMPILE time, on a program that
// merely names the type. Hoisting into locals does not help; the count is what
// matters.
//
// So anything needing more than one operation is emitted HERE, as a single
// builtin, rather than composed in the synthesized C#. That is a workaround for
// a shared-compiler defect, not a design preference — when the defect is fixed,
// `Sign`/`Abs`/`Clamp` could equally be three lines of C# again.

use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// A BigInt constant on the stack.
fn big_const(chunk: &mut Chunk, n: i32, line: u32) {
    chunk.emit_i32_const(n, line);
    to_bigint(chunk, line);
}

/// `<slot> <op> <n>n` as an i32 condition.
fn cmp_const(chunk: &mut Chunk, slot: u16, op: &'static str, n: i32, line: u32) {
    get(chunk, slot, line);
    big_const(chunk, n, line);
    host::emit(chunk, "ecma:bigint", op, 2, line);
    ops::emit_dyn_to_bool(chunk, line);
}

/// `Sign(v)` — `-1`, `0` or `1`. Stack `[bigint] -> [i32]`.
pub fn emit_sign(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    cmp_const(&mut chunks[current], v, "lt", 0, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_else(line);
    cmp_const(&mut chunks[current], v, "eq", 0, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `Abs(v)`. Stack `[bigint] -> [bigint]`.
pub fn emit_abs(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    cmp_const(&mut chunks[current], v, "lt", 0, line);
    chunks[current].emit_if_value(line);
    big_const(&mut chunks[current], 0, line);
    get(&mut chunks[current], v, line);
    host::emit(&mut chunks[current], "ecma:bigint", "sub", 2, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], v, line);
    chunks[current].emit_end(line);
}

/// `v % 2n == 0n`. Stack `[bigint] -> [bool]`.
pub fn emit_is_even(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    get(&mut chunks[current], v, line);
    big_const(&mut chunks[current], 2, line);
    host::emit(&mut chunks[current], "ecma:bigint", "rem", 2, line);
    let r = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], r, line);
    cmp_const(&mut chunks[current], r, "eq", 0, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `CompareTo` — the sign of `a - b`. Stack `[a, b] -> [i32]`.
pub fn emit_cmp(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:bigint", "sub", 2, line);
    emit_sign(chunks, current, line);
}

/// `Clamp(v, min, max)`. Stack `[v, min, max] -> [bigint]`.
pub fn emit_clamp(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(3);
    let (v, lo, hi) = (base, base + 1, base + 2);
    set(&mut chunks[current], hi, line);
    set(&mut chunks[current], lo, line);
    set(&mut chunks[current], v, line);

    get(&mut chunks[current], v, line);
    get(&mut chunks[current], lo, line);
    host::emit(&mut chunks[current], "ecma:bigint", "lt", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], lo, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], v, line);
    get(&mut chunks[current], hi, line);
    host::emit(&mut chunks[current], "ecma:bigint", "gt", 2, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], hi, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], v, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `Min` / `Max`. Stack `[a, b] -> [bigint]`.
pub fn emit_pick(chunks: &mut [Chunk], current: usize, smallest: bool, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (a, b) = (base, base + 1);
    set(&mut chunks[current], b, line);
    set(&mut chunks[current], a, line);
    get(&mut chunks[current], a, line);
    get(&mut chunks[current], b, line);
    host::emit(
        &mut chunks[current],
        "ecma:bigint",
        if smallest { "lt" } else { "gt" },
        2,
        line,
    );
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], a, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], b, line);
    chunks[current].emit_end(line);
}

/// The decimal text of a BigInt. Stack `[bigint] -> [string]`.
///
/// ⛔ `"" + v` DOES NOT WORK. String concatenation with a BigInt payload trips
/// `wasm:js-number.toF64 — not a number`: the `+` lowers to a numeric add
/// before anything notices the operand is a BigInt. §21.2.3.3 `toString` is a
/// host function and is exact at any magnitude, which is the whole point —
/// a 39-digit `Int128.MaxValue` has no f64 spelling to fall back to.
pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:bigint", "toString", 1, line);
}

/// Whether a string is a parseable integer. Stack `[string] -> [bool]`.
///
/// `BigInt("invalid_number")` THROWS `SyntaxError` (§21.2.1.1), so `TryParse`
/// cannot just try it and look at the result — there is no result. `Number(s)`
/// is `NaN` for exactly the inputs `BigInt` rejects and is total, so it is the
/// validity test; the VALUE still comes from `BigInt`, which is exact where
/// `Number` would already have lost digits.
pub fn emit_is_numeric(chunks: &mut [Chunk], current: usize, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    let number = chunks[current].add_import("ecma:number", "Number");
    chunks[current].emit_call(number, 1, line);
    set(&mut chunks[current], n, line);
    // NaN is the only value that differs from itself.
    get(&mut chunks[current], n, line);
    get(&mut chunks[current], n, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `IsPow2(v)` — `v > 0 && (v & (v - 1)) == 0`. Stack `[bigint] -> [bool]`.
///
/// Emitted here rather than composed in the synthesized class because the
/// composite operations all live on this side; see the note above.
pub fn emit_is_pow2(chunks: &mut [Chunk], current: usize, line: u32) {
    let v = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], v, line);
    cmp_const(&mut chunks[current], v, "gt", 0, line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], v, line);
    get(&mut chunks[current], v, line);
    big_const(&mut chunks[current], 1, line);
    host::emit(&mut chunks[current], "ecma:bigint", "sub", 2, line);
    host::emit(&mut chunks[current], "ecma:bigint", "and", 2, line);
    let masked = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], masked, line);
    cmp_const(&mut chunks[current], masked, "eq", 0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `ToString("X")` / `ToString("x")` — hex. Stack `[bigint, upper] -> [string]`.
///
/// §21.2.3.3 `BigInt.prototype.toString(radix)` renders LOWERCASE, and .NET's
/// `"X"` is uppercase (`"x"` is the lowercase one), so the case is applied
/// here rather than left to the caller.
pub fn emit_to_hex(chunks: &mut [Chunk], current: usize, line: u32) {
    let upper = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], upper, line);
    chunks[current].emit_i32_const(16, line);
    host::emit(&mut chunks[current], "ecma:bigint", "toStringRadix", 2, line);
    let text = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], text, line);
    get(&mut chunks[current], upper, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], text, line);
    vybe_compiler::primitives::strings::emit_to_upper(&mut chunks[current], line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], text, line);
    chunks[current].emit_end(line);
}
