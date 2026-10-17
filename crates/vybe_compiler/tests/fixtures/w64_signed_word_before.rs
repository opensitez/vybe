#![allow(dead_code, unused_imports)]
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
//! Number-backed operands are converted into the declared integer lane.
//! Arithmetic returns a number; comparisons return an i32 predicate.
//!
//! [`LeadingZeros`]: UnaryOp::LeadingZeros
//! [`TrailingZeros`]: UnaryOp::TrailingZeros
//! [`RotR`]: BinOp::RotR

use vybe_ast::{BitLane, IntOp, NumericRepr, ShiftOverflow};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Raw locals cannot reuse a slot that previously held a value: Wasm local
/// types cover the entire function, including earlier statements.
pub(crate) fn raw_i64_scratch(chunk: &mut Chunk) -> u16 {
    if chunk.i64_scratch_slot.is_none() {
        chunk.local_count = chunk.local_count.max(chunk.scratch_high_water);
    }
    chunk.i64_scratch()
}

/// Bring the value on the stack into `lane`'s integer type.
///
/// Convert the numeric value to its integer lane. A direct i32 coercion can
/// saturate values above INT32_MAX; wrapping through i64 preserves all 32 bits.
fn narrow(chunk: &mut Chunk, lane: BitLane, line: u32) {
    match lane {
        BitLane::W32 => {
            chunk.emit_op(Op::I64_TRUNC_F64_S, line);
            chunk.emit_op(Op::I32_WRAP_I64, line);
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

fn narrow_comparison_word(chunk: &mut Chunk, lane: BitLane, line: u32) {
    if lane == BitLane::W32 {
        narrow(chunk, lane, line);
        return;
    }
    let word = retain_wide_comparison_word(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, word, line);
}

// Keep the input below the converted operand live on the stack. No source
// expression runs during either conversion, and conversion order is unchanged.
fn retain_wide_comparison_word(chunk: &mut Chunk, line: u32) -> u16 {
    // Number-backed unsigned words can occupy the upper half of u64. Both
    // conversions produce the same i64 bit lane; signed ordering is selected
    // by the comparison opcode, not by lossy conversion through i32.
    let value = chunk.alloc_scratch(1);
    let word = raw_i64_scratch(chunk);
    chunk.emit_op_u16(Op::LOCAL_TEE, value, line);
    chunk.emit_f64_const(9_223_372_036_854_775_808.0, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::I64_TRUNC_F64_U, line);
    chunk.emit_op_u16(Op::LOCAL_SET, word, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::I64_TRUNC_F64_S, line);
    chunk.emit_op_u16(Op::LOCAL_SET, word, line);
    chunk.emit_end(line);
    word
}

fn emit_wide_word_comparison(chunk: &mut Chunk, op: IntOp, line: u32) {
    // The right operand must convert first, including any conversion trap.
    // Its word remains in the dedicated i64 slot while the left is converted.
    let right = retain_wide_comparison_word(chunk, line);
    let left = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_TEE, left, line);
    chunk.emit_f64_const(9_223_372_036_854_775_808.0, line);
    chunk.emit_op(Op::F64_GE, line);
    chunk.emit_if_i32(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left, line);
    chunk.emit_op(Op::I64_TRUNC_F64_U, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    chunk.emit_op(integer_opcode(BitLane::W64, op), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left, line);
    chunk.emit_op(Op::I64_TRUNC_F64_S, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    chunk.emit_op(integer_opcode(BitLane::W64, op), line);
    chunk.emit_end(line);
}

pub fn emit_normalize_integer(chunk: &mut Chunk, lane: BitLane, line: u32) {
    narrow(chunk, lane, line);
    widen(chunk, lane, line);
}

/// One unary bit-count instruction. Skip narrowing only for a proven W32
/// operand; the W64 number representation always retains its conversions.
pub fn emit_count_with_normalized_operand(
    chunk: &mut Chunk,
    lane: BitLane,
    w32: Op,
    w64: Op,
    normalized: bool,
    line: u32,
) {
    if lane != BitLane::W32 || !normalized {
        narrow(chunk, lane, line);
    }
    chunk.emit_op(if lane == BitLane::W32 { w32 } else { w64 }, line);
    widen(chunk, lane, line);
}

pub fn emit_pop_count(chunk: &mut Chunk, lane: BitLane, line: u32) {
    emit_count_with_normalized_operand(chunk, lane, Op::I32_POPCNT, Op::I64_POPCNT, false, line);
}

pub fn emit_leading_zeros(chunk: &mut Chunk, lane: BitLane, line: u32) {
    emit_count_with_normalized_operand(chunk, lane, Op::I32_CLZ, Op::I64_CLZ, false, line);
}

pub fn emit_trailing_zeros(chunk: &mut Chunk, lane: BitLane, line: u32) {
    emit_count_with_normalized_operand(chunk, lane, Op::I32_CTZ, Op::I64_CTZ, false, line);
}

/// `rotl` / `rotr` — stack on entry is `[value, count]`.
///
/// wasm's rotate already takes the count modulo the width, which is what every
/// language spelling this wants: a rotation by the full width IS the identity,
/// so [`ShiftOverflow`] does not apply here the way it does to a shift.
pub fn emit_rotate(chunk: &mut Chunk, lane: BitLane, left: bool, line: u32) {
    emit_rotate_with_normalized_operands(chunk, lane, left, (false, false), line);
}

/// Skip conversion only when the emitter proved an operand is already a W32 word.
/// A normalized left word can remain on the stack while the count is narrowed.
pub fn emit_rotate_with_normalized_operands(
    chunk: &mut Chunk,
    lane: BitLane,
    left: bool,
    normalized: (bool, bool),
    line: u32,
) {
    if lane == BitLane::W64 {
        // Only the count must leave the operand stack. Reserve a persistent raw
        // i64 local; ordinary scratch is externref and can be reclaimed later.
        let count = raw_i64_scratch(chunk);
        narrow(chunk, lane, line);
        chunk.emit_op_u16(Op::LOCAL_SET, count, line);
        narrow(chunk, lane, line);
        chunk.emit_op_u16(Op::LOCAL_GET, count, line);
        chunk.emit_op(if left { Op::I64_ROTL } else { Op::I64_ROTR }, line);
        widen(chunk, lane, line);
        return;
    }
    if normalized.0 {
        if !normalized.1 {
            narrow(chunk, lane, line);
        }
        chunk.emit_op(if left { Op::I32_ROTL } else { Op::I32_ROTR }, line);
        return;
    }
    let count = chunk.alloc_scratch(1);
    // Preserve the count while narrowing the left operand underneath it.
    if lane != BitLane::W32 || !normalized.1 {
        narrow(chunk, lane, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);
    if lane != BitLane::W32 || !normalized.0 {
        narrow(chunk, lane, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    let op = match (lane, left) {
        (BitLane::W32, true) => Op::I32_ROTL,
        (BitLane::W32, false) => Op::I32_ROTR,
        (BitLane::W64, true) => Op::I64_ROTL,
        (BitLane::W64, false) => Op::I64_ROTR,
    };
    chunk.emit_op(op, line);
    widen(chunk, lane, line);
}

/// Compare numeric word operands after preserving their complete 32-bit lanes.
pub fn emit_word_comparison(chunk: &mut Chunk, opcode: Op, line: u32) {
    emit_word_comparison_with_normalized_operands(chunk, opcode, (false, false), line);
}

pub fn emit_word_comparison_with_normalized_operands(
    chunk: &mut Chunk,
    opcode: Op,
    normalized: (bool, bool),
    line: u32,
) {
    if normalized.0 {
        if !normalized.1 {
            narrow(chunk, BitLane::W32, line);
        }
        chunk.emit_op(opcode, line);
        return;
    }
    let right = chunk.alloc_scratch(1);
    if !normalized.1 {
        narrow(chunk, BitLane::W32, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, right, line);
    if !normalized.0 {
        narrow(chunk, BitLane::W32, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    chunk.emit_op(opcode, line);
}

/// Stack `[left, right]` -> one numeric result, using the declared integer
/// lane instead of the profile's default arithmetic slot.
pub fn emit_integer_binary(chunk: &mut Chunk, lane: BitLane, op: IntOp, line: u32) {
    emit_integer_binary_with_normalized_operands(chunk, lane, op, (false, false), line);
}

/// Operand proofs come from explicit integer-lane expressions, not type names.
pub fn emit_integer_binary_with_normalized_operands(
    chunk: &mut Chunk,
    lane: BitLane,
    op: IntOp,
    normalized: (bool, bool),
    line: u32,
) {
    if lane == BitLane::W32 && normalized.0 {
        // The left word needs no conversion. Normalize only the top operand;
        // both words then remain in the instruction's original stack order.
        if !normalized.1 {
            narrow(chunk, lane, line);
        }
        chunk.emit_op(integer_opcode(lane, op), line);
        return;
    }
    if lane == BitLane::W64 && op.is_comparison() {
        emit_wide_word_comparison(chunk, op, line);
        return;
    }
    {
        let right = if lane == BitLane::W64 {
            raw_i64_scratch(chunk)
        } else {
            chunk.alloc_scratch(1)
        };
        let narrow_operand = |chunk: &mut Chunk| {
            if op.is_comparison() {
                narrow_comparison_word(chunk, lane, line);
            } else {
                narrow(chunk, lane, line);
            }
        };
        if !normalized.1 {
            narrow_operand(chunk);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, right, line);
        if !normalized.0 {
            narrow_operand(chunk);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, right, line);
    }
    chunk.emit_op(integer_opcode(lane, op), line);
    if !op.is_comparison() {
        widen(chunk, lane, line);
    }
}

pub(crate) fn integer_opcode(lane: BitLane, op: IntOp) -> Op {
    match (lane, op) {
        (BitLane::W32, IntOp::Add) => Op::I32_ADD,
        (BitLane::W32, IntOp::Sub) => Op::I32_SUB,
        (BitLane::W32, IntOp::Mul) => Op::I32_MUL,
        (BitLane::W32, IntOp::DivS) => Op::I32_DIV_S,
        (BitLane::W32, IntOp::DivU) => Op::I32_DIV_U,
        (BitLane::W32, IntOp::RemS) => Op::I32_REM_S,
        (BitLane::W32, IntOp::RemU) => Op::I32_REM_U,
        (BitLane::W32, IntOp::Shl) => Op::I32_SHL,
        (BitLane::W32, IntOp::ShrS) => Op::I32_SHR_S,
        (BitLane::W32, IntOp::ShrU) => Op::I32_SHR_U,
        (BitLane::W32, IntOp::And) => Op::I32_AND,
        (BitLane::W32, IntOp::Or) => Op::I32_OR,
        (BitLane::W32, IntOp::Xor) => Op::I32_XOR,
        (BitLane::W32, IntOp::Eq) => Op::I32_EQ,
        (BitLane::W32, IntOp::Ne) => Op::I32_NE,
        (BitLane::W32, IntOp::LtS) => Op::I32_LT_S,
        (BitLane::W32, IntOp::LtU) => Op::I32_LT_U,
        (BitLane::W32, IntOp::LeS) => Op::I32_LE_S,
        (BitLane::W32, IntOp::LeU) => Op::I32_LE_U,
        (BitLane::W32, IntOp::GtS) => Op::I32_GT_S,
        (BitLane::W32, IntOp::GtU) => Op::I32_GT_U,
        (BitLane::W32, IntOp::GeS) => Op::I32_GE_S,
        (BitLane::W32, IntOp::GeU) => Op::I32_GE_U,
        (BitLane::W64, IntOp::Add) => Op::I64_ADD,
        (BitLane::W64, IntOp::Sub) => Op::I64_SUB,
        (BitLane::W64, IntOp::Mul) => Op::I64_MUL,
        (BitLane::W64, IntOp::DivS) => Op::I64_DIV_S,
        (BitLane::W64, IntOp::DivU) => Op::I64_DIV_U,
        (BitLane::W64, IntOp::RemS) => Op::I64_REM_S,
        (BitLane::W64, IntOp::RemU) => Op::I64_REM_U,
        (BitLane::W64, IntOp::Shl) => Op::I64_SHL,
        (BitLane::W64, IntOp::ShrS) => Op::I64_SHR_S,
        (BitLane::W64, IntOp::ShrU) => Op::I64_SHR_U,
        (BitLane::W64, IntOp::And) => Op::I64_AND,
        (BitLane::W64, IntOp::Or) => Op::I64_OR,
        (BitLane::W64, IntOp::Xor) => Op::I64_XOR,
        (BitLane::W64, IntOp::Eq) => Op::I64_EQ,
        (BitLane::W64, IntOp::Ne) => Op::I64_NE,
        (BitLane::W64, IntOp::LtS) => Op::I64_LT_S,
        (BitLane::W64, IntOp::LtU) => Op::I64_LT_U,
        (BitLane::W64, IntOp::LeS) => Op::I64_LE_S,
        (BitLane::W64, IntOp::LeU) => Op::I64_LE_U,
        (BitLane::W64, IntOp::GtS) => Op::I64_GT_S,
        (BitLane::W64, IntOp::GtU) => Op::I64_GT_U,
        (BitLane::W64, IntOp::GeS) => Op::I64_GE_S,
        (BitLane::W64, IntOp::GeU) => Op::I64_GE_U,
    }
}
