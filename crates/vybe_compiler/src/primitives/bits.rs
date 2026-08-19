//! Integer bit operations — one home for a family every language has.
//!
//! `i32.popcnt` is ONE wasm instruction, and before this module it was reached
//! three different ways: `common:fortran.popcnt` through a Fortran-local
//! adapter (i32 lane), `common:go.bits_ones_count` through a Go-local adapter
//! (i64 lane), and `opcode:i32_popcnt` from wast. Two emitters, opposite lanes,
//! neither aware of the other — the shape `directives.md` §9.2 describes, where
//! N spellings drift and each copy looks locally correct.
//!
//! The operations are [`UnaryOp::PopCount`] / [`LeadingZeros`] /
//! [`TrailingZeros`] and [`BinOp::RotL`] / [`RotR`], each carrying its
//! [`BitLane`], because the lane is a property of the operand's declared type
//! and wasm spells the two lanes as different instructions.
//!
//! Values arrive and leave as f64 — the runtime's number representation — so
//! each helper converts in, works in the integer lane, and converts back.
//!
//! [`LeadingZeros`]: UnaryOp::LeadingZeros
//! [`TrailingZeros`]: UnaryOp::TrailingZeros
//! [`RotR`]: BinOp::RotR

use vybe_ast::{BitLane, NumericRepr, ShiftOverflow};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Bring the value on the stack into `lane`'s integer type.
///
/// The `I32_*` opcodes coerce their operands dynamically, so the 32-bit lane
/// only has to force the coercion (`| 0`, the same trick every i32 site in the
/// tree uses). The 64-bit lane needs a real narrowing.
fn narrow(chunk: &mut Chunk, lane: BitLane, line: u32) {
    match lane {
        BitLane::W32 => {
            chunk.emit_i32_const(0, line);
            chunk.emit_op(Op::I32_OR, line);
        }
        BitLane::W64 => chunk.emit_op(Op::I64_TRUNC_F64_S, line),
    }
}

/// Return `lane`'s integer to the number representation the runtime carries.
/// The 32-bit lane is already there.
fn widen(chunk: &mut Chunk, lane: BitLane, line: u32) {
    if lane == BitLane::W64 {
        chunk.emit_op(Op::F64_CONVERT_I64_S, line);
    }
}

/// One unary bit-count instruction, wrapped in the lane's conversions.
fn emit_count(chunk: &mut Chunk, lane: BitLane, w32: Op, w64: Op, line: u32) {
    narrow(chunk, lane, line);
    chunk.emit_op(if lane == BitLane::W32 { w32 } else { w64 }, line);
    widen(chunk, lane, line);
}

pub fn emit_pop_count(chunk: &mut Chunk, lane: BitLane, line: u32) {
    emit_count(chunk, lane, Op::I32_POPCNT, Op::I64_POPCNT, line);
}

pub fn emit_leading_zeros(chunk: &mut Chunk, lane: BitLane, line: u32) {
    emit_count(chunk, lane, Op::I32_CLZ, Op::I64_CLZ, line);
}

pub fn emit_trailing_zeros(chunk: &mut Chunk, lane: BitLane, line: u32) {
    emit_count(chunk, lane, Op::I32_CTZ, Op::I64_CTZ, line);
}

/// `rotl` / `rotr` — stack on entry is `[value, count]`.
///
/// wasm's rotate already takes the count modulo the width, which is what every
/// language spelling this wants: a rotation by the full width IS the identity,
/// so [`ShiftOverflow`] does not apply here the way it does to a shift.
pub fn emit_rotate(chunk: &mut Chunk, lane: BitLane, left: bool, line: u32) {
    let base = chunk.alloc_scratch(2);
    // The count is on top, so it pops first.
    narrow(chunk, lane, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base + 1, line);
    narrow(chunk, lane, line);
    chunk.emit_op_u16(Op::LOCAL_SET, base, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    chunk.emit_op_u16(Op::LOCAL_GET, base + 1, line);
    let op = match (lane, left) {
        (BitLane::W32, true) => Op::I32_ROTL,
        (BitLane::W32, false) => Op::I32_ROTR,
        (BitLane::W64, true) => Op::I64_ROTL,
        (BitLane::W64, false) => Op::I64_ROTR,
    };
    chunk.emit_op(op, line);
    widen(chunk, lane, line);
}

/// Read a value's STORAGE as another type of the same width — a bit cast.
///
/// The fifth member of the family, and the one every language spells
/// differently: Fortran `TRANSFER`, Go `math.Float32bits`, Java
/// `Float.floatToIntBits`, C#/VB `BitConverter.SingleToInt32Bits`, C `union`,
/// C++ `std::bit_cast`, Rust `transmute`, JS `DataView`. Before this there were
/// five implementations — two correct, one that returned its argument
/// unchanged (VB `BitConverter` answered `1` for `SingleToInt32Bits(1.0f)`
/// where .NET answers `1065353216`), and two that did not resolve at all.
///
/// The runtime carries numbers as f64, so the 32-bit lanes narrow on the way in
/// and widen on the way out; the 64-bit lanes are already the right width.
pub fn emit_reinterpret(chunk: &mut Chunk, to: NumericRepr, line: u32) {
    match to {
        // f32 → i32
        NumericRepr::I32 => {
            chunk.emit_op(Op::F32_DEMOTE_F64, line);
            chunk.emit_op(Op::I32_REINTERPRET_F32, line);
        }
        // i32 → f32
        NumericRepr::F32 => {
            chunk.emit_op(Op::F32_REINTERPRET_I32, line);
            chunk.emit_op(Op::F64_PROMOTE_F32, line);
        }
        // f64 → i64
        NumericRepr::I64 => chunk.emit_op(Op::I64_REINTERPRET_F64, line),
        // i64 → f64
        NumericRepr::F64 => chunk.emit_op(Op::F64_REINTERPRET_I64, line),
    }
}

/// Does this region's [`ShiftOverflow`] policy need a guard around a shift?
///
/// wasm masks the count, so `Mask` is already what the bare instruction does
/// and costs nothing. `Zero` is the one that needs code.
pub fn shift_needs_guard(policy: ShiftOverflow) -> bool {
    policy == ShiftOverflow::Zero
}
