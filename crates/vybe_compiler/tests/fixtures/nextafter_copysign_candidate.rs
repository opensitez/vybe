use vybe_ast::FloatLane;
use vybe_runtime::{Chunk,Op};
const MIN_SUBNORMAL: f64 = 5e-324;

/// Smallest positive f32 subnormal — bit pattern 1 in the 32-bit lane.
const MIN_SUBNORMAL_F32: f64 = 1.401298464324817e-45;

/// `nextUp(x)` (`up`) or `nextDown(x)`. Stack: `[x] → [adjacent]`.
pub fn emit_next_toward(chunk: &mut Chunk, up: bool, lane: FloatLane, line: u32) {
    let x = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);
    emit_next_toward_from_slot(chunk, x, up, lane, line);
}

fn emit_next_toward_from_slot(chunk: &mut Chunk, x: u16, up: bool, lane: FloatLane, line: u32) {
    // NaN has no neighbour.
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_NE, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_else(line);

    // The infinity we are travelling towards is its own successor.
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(if up { f64::INFINITY } else { f64::NEG_INFINITY }, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_else(line);

    emit_finite_next_toward_from_slot(chunk, x, up, lane, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

// x is ordered and is not the infinity in the travelled direction.
fn emit_finite_next_toward_from_slot(
    chunk: &mut Chunk, x: u16, up: bool, lane: FloatLane, line: u32,
) {
    // ±0 both step to the smallest subnormal of the travelled sign — the bit
    // pattern is NOT contiguous across the sign bit, so this cannot be a ±1.
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    let tiny = match lane {
        FloatLane::F32 => MIN_SUBNORMAL_F32,
        FloatLane::F64 => MIN_SUBNORMAL,
    };
    chunk.emit_f64_const(if up { tiny } else { -tiny }, line);
    chunk.emit_else(line);

    // Nonzero ordered x: signed unit controls the adjacent bit pattern.
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    match lane {
        FloatLane::F32 => vybe_compiler::primitives::bits::emit_reinterpret(chunk, vybe_ast::NumericRepr::I32, line),
        FloatLane::F64 => vybe_compiler::primitives::bits::emit_reinterpret(chunk, vybe_ast::NumericRepr::I64, line),
    }
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_COPYSIGN, line);
    match lane {
        FloatLane::F32 => {
            chunk.emit_op(Op::I32_FROM_F64, line);
            chunk.emit_op(if up { Op::I32_ADD } else { Op::I32_SUB }, line);
            vybe_compiler::primitives::bits::emit_reinterpret(chunk, vybe_ast::NumericRepr::F32, line);
        }
        FloatLane::F64 => {
            chunk.emit_op(Op::I64_TRUNC_F64_S, line);
            chunk.emit_op(if up { Op::I64_ADD } else { Op::I64_SUB }, line);
            vybe_compiler::primitives::bits::emit_reinterpret(chunk, vybe_ast::NumericRepr::F64, line);
        }
    }
    chunk.emit_end(line);
}

/// `nextafter(x, y)` — the neighbour of `x` in `y`'s direction, `y` when equal.
/// Stack: `[x, y] → [result]`.
pub fn emit_next_after(chunk: &mut Chunk, lane: FloatLane, line: u32) {
    chunk
        .local_count
        .checked_add(2)
        .expect("next_after full local budget");
    let base = chunk.alloc_scratch(2);
    chunk.emit_op_u16(Op::LOCAL_SET, base + 1, line); // y
    chunk.emit_op_u16(Op::LOCAL_SET, base, line); // x

    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if_value(line);
    // IEEE-754 §5.3.1: equal operands yield the SECOND, which settles the sign
    // for `nextafter(0.0, -0.0)`.
    chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if_value(line);
    // y > x proves x is ordered and cannot be positive infinity.
    emit_finite_next_toward_from_slot(chunk, base, true, lane, line);
    chunk.emit_else(line);
    emit_next_toward_from_slot(chunk, base, false, lane, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `ulp(x)` / Fortran `SPACING(x)` — the distance to the next representable
/// value away from zero. Derived from the neighbour rather than from the
/// exponent, so the binade boundary is right by construction.
pub fn emit_ulp(chunk: &mut Chunk, lane: FloatLane, line: u32) {
    let x = chunk.alloc_scratch(1);
    chunk.emit_op(Op::F64_ABS, line);
    chunk.emit_op_u16(Op::LOCAL_TEE, x, line);
    emit_next_toward(chunk, true, lane, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op(Op::F64_SUB, line);
}

