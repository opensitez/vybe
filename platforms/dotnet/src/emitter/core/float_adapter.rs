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

/// The smallest NORMAL double — `2^-1022`. Below it the exponent field is zero
/// and the significand loses its implicit leading one, which is what
/// "subnormal" names.
///
/// ⛔ Not `Double.Epsilon`. That is the smallest subnormal (`5e-324`) and sits
/// at the other end of the same range; using it as the boundary calls every
/// subnormal normal.
const MIN_NORMAL: f64 = f64::MIN_POSITIVE;

/// `[x] → [finite && x <> 0]` in the operand's slot, leaving the ABSOLUTE
/// value on the stack — the shared half of `IsNormal` and `IsSubnormal`.
fn emit_finite_nonzero_magnitude(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    lget(chunk, slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    lget(chunk, slot, line);
    chunk.emit_f64_const(f64::INFINITY, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, slot, line);
    chunk.emit_f64_const(f64::NEG_INFINITY, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_op(Op::I32_AND, line);
    lget(chunk, slot, line);
    chunk.emit_op(Op::F64_ABS, line);
}

/// `Double.IsNormal(x)` — finite, non-zero, and at or above the smallest
/// normal magnitude. Zero is NOT normal, which is the case the name hides.
///
/// ⚠ One emit serves `Double` and `Single`, as every predicate here does, but
/// this is the one whose boundary is width-dependent: `Single`'s smallest
/// normal is `1.17549435e-38`. A `Single` on this platform is stored as an
/// f64, so a value in between is not actually subnormal in the storage it
/// lives in, and the double boundary is the honest answer for it.
pub fn emit_is_normal(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let slot = operand(chunk, line);
    emit_finite_nonzero_magnitude(chunk, slot, line);
    chunk.emit_f64_const(MIN_NORMAL, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_op(Op::I32_AND, line);
    as_bool(chunk, line);
}

/// `Double.IsSubnormal(x)` — finite, non-zero, and below the smallest normal.
pub fn emit_is_subnormal(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let slot = operand(chunk, line);
    emit_finite_nonzero_magnitude(chunk, slot, line);
    chunk.emit_f64_const(MIN_NORMAL, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_op(Op::I32_AND, line);
    as_bool(chunk, line);
}
