//! `System.Double` / `System.Single` static predicates, shared by .NET
//! languages.
//!
//! `Double.NaN`, `Double.PositiveInfinity` and `Double.NegativeInfinity` were
//! already known CONSTANTS (`core::types`), but the PREDICATES that read them
//! back — `IsNaN`, `IsInfinity`, `IsPositiveInfinity`, `IsNegativeInfinity`,
//! `IsFinite` — were registered nowhere, so `Double.IsNaN(x)` resolved to
//! nothing and answered `null` where a Boolean was expected.
//!
//! Each is IEEE arithmetic on the operand, not a host call: `x <> x` is the
//! definition of NaN, and the two infinities are the constants above. The
//! result is lifted to a real Boolean because .NET returns `Boolean` and VB
//! renders that as `True`/`False` — an i32 would print `1`.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// Park the operand so it can be read twice, and hand back its slot.
fn operand(chunk: &mut Chunk, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn as_bool(chunk: &mut Chunk, line: u32) {
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `Double.IsNaN(x)` — `x <> x`, true for NaN alone.
pub fn emit_is_nan(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let slot = operand(chunk, line);
    lget(chunk, slot, line);
    lget(chunk, slot, line);
    chunk.emit_op(Op::F64_NE, line);
    as_bool(chunk, line);
}

/// `Double.IsPositiveInfinity(x)`.
pub fn emit_is_positive_infinity(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    as_bool(chunk, line);
}

/// `Double.IsNegativeInfinity(x)`.
pub fn emit_is_negative_infinity(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_f64_const(f64::NEG_INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    as_bool(chunk, line);
}

/// `Double.IsInfinity(x)` — either sign.
pub fn emit_is_infinity(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let slot = operand(chunk, line);
    lget(chunk, slot, line);
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    lget(chunk, slot, line);
    chunk.emit_f64_const(f64::NEG_INFINITY, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_OR, line);
    as_bool(chunk, line);
}

/// `Double.IsFinite(x)` — neither NaN nor infinite.
pub fn emit_is_finite(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let slot = operand(chunk, line);
    // x = x  (not NaN)
    lget(chunk, slot, line);
    lget(chunk, slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    // and |x| <> +inf
    lget(chunk, slot, line);
    chunk.emit_op(Op::F64_ABS, line);
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_op(Op::I32_AND, line);
    as_bool(chunk, line);
}
