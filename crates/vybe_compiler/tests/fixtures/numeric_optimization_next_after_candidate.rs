// Isolated branch-preserving scalar-copy elimination. Production unchanged.
use vybe_ast::FloatLane;
use vybe_runtime::{Chunk, Op};
const MIN_SUBNORMAL: f64 = f64::from_bits(1);
const MIN_SUBNORMAL_F32: f64 = f32::from_bits(1) as f64;
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

    // Magnitude grows away from zero, so the step's sign follows the operand's.
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if_value(line);
    emit_bit_step(chunk, if up { 1 } else { -1 }, x, lane, line);
    chunk.emit_else(line);
    emit_bit_step(chunk, if up { -1 } else { 1 }, x, lane, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `local[x]`'s bit pattern, shifted by `step`, read back as an `f64`.
fn emit_bit_step(chunk: &mut Chunk, step: i64, x: u16, lane: FloatLane, line: u32) {
    use vybe_ast::NumericRepr;
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    match lane {
        FloatLane::F32 => {
            vybe_compiler::primitives::bits::emit_reinterpret(chunk, NumericRepr::I32, line);
            chunk.emit_i32_const(step as i32, line);
            chunk.emit_op(Op::I32_ADD, line);
            vybe_compiler::primitives::bits::emit_reinterpret(chunk, NumericRepr::F32, line);
        }
        FloatLane::F64 => {
            vybe_compiler::primitives::bits::emit_reinterpret(chunk, NumericRepr::I64, line);
            chunk.emit_i64_const(step, line);
            chunk.emit_op(Op::I64_ADD, line);
            vybe_compiler::primitives::bits::emit_reinterpret(chunk, NumericRepr::F64, line);
        }
    }
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
    emit_next_toward_from_slot(chunk, base, true, lane, line);
    chunk.emit_else(line);
    emit_next_toward_from_slot(chunk, base, false, lane, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
