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

// ── Elementwise bit operations over a boolean ARRAY ────────────────────────
//
// A bit set is a boolean sequence, and combining two of them elementwise is the
// same operation the scalar helpers above perform on one word: .NET's
// `BitArray`, Java's `BitSet`, Python's `bitarray`, C++'s `vector<bool>`.
//
// The operators mutate the receiver and return it, which is what those APIs
// document and what lets calls chain.

use crate::primitives::collections::emit_import_call as bits_import_call;

/// The combining instruction for one elementwise pass, or `None` to negate.
fn emit_elementwise(chunks: &mut [Chunk], current: usize, op: Option<Op>, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let me = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);

    if op.is_some() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, other, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, me, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, me, line);
    bits_import_call(chunks, current, "ecma:array", "length", 1, line);
    // `| 0` is the narrowing every i32 site in this file uses: the `I32_*`
    // instructions coerce their operands dynamically.
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);

    let done = chunks[current].emit_block(line);
    let (again, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, me, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, me, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    bits_import_call(chunks, current, "ecma:array", "get", 2, line);
    bits_import_call(chunks, current, "wasm:js-boolean", "cast", 1, line);
    match op {
        Some(instr) => {
            chunks[current].emit_op_u16(Op::LOCAL_GET, other, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
            bits_import_call(chunks, current, "ecma:array", "get", 2, line);
            bits_import_call(chunks, current, "wasm:js-boolean", "cast", 1, line);
            chunks[current].emit_op(instr, line);
        }
        None => chunks[current].emit_op(Op::I32_EQZ, line),
    }
    // A real boolean, not the i32: under `materialize_bool_results` an i32 1
    // does not compare equal to the language's `true`.
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);

    bits_import_call(chunks, current, "ecma:array", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(again);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done);

    chunks[current].emit_op_u16(Op::LOCAL_GET, me, line);
}

/// Elementwise AND, in place. Stack: `[a, b]` → `[a]`.
pub fn emit_array_and(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_elementwise(chunks, current, Some(Op::I32_AND), line);
}

/// Elementwise OR, in place. Stack: `[a, b]` → `[a]`.
pub fn emit_array_or(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_elementwise(chunks, current, Some(Op::I32_OR), line);
}

/// Elementwise XOR, in place. Stack: `[a, b]` → `[a]`.
pub fn emit_array_xor(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_elementwise(chunks, current, Some(Op::I32_XOR), line);
}

/// Elementwise NOT, in place. Stack: `[a]` → `[a]`.
pub fn emit_array_not(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_elementwise(chunks, current, None, line);
}

/// Position of the highest set bit — `31 - clz` in the 32-bit lane.
///
/// An integer operation, not a floating-point logarithm: `log2(64)` is exactly
/// 6 where `ln(64)/ln(2)` can land a fraction below it.
///
/// Stack: `[x]` → `[n]`.
pub fn emit_log2(chunks: &mut [Chunk], current: usize, line: u32) {
    let tmp = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op(Op::I32_CLZ, line);
    chunk.emit_op_u16(Op::LOCAL_SET, tmp, line);
    chunk.emit_i32_const(31, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tmp, line);
    chunk.emit_op(Op::I32_SUB, line);
}

/// Whether exactly one bit is set. Zero is NOT a power of two, which the
/// `x & (x - 1)` test alone would report as one.
///
/// Stack: `[x]` → `[bool]`.
pub fn emit_is_pow2(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

/// The smallest power of two at or above `x`.
///
/// 0 and 1 are returned unchanged: 0 has no power of two at or above it that
/// fits the lane, and both would otherwise shift by the lane's full 32, which
/// wasm takes modulo the width and turns into a shift of zero.
///
/// Stack: `[x]` → `[n]`.
pub fn emit_round_up_pow2(chunks: &mut [Chunk], current: usize, line: u32) {
    let x = chunks[current].alloc_scratch(1);
    let chunk = &mut chunks[current];
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);

    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_LE_S, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_i32_const(32, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::I32_CLZ, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op(Op::I32_SHL, line);
    chunk.emit_end(line);
}

/// Reinterpret the 32-bit lane's signed result as unsigned.
///
/// The lane's instructions are signed, so a value with the high bit set comes
/// back negative. Languages whose bit types are unsigned — .NET's `uint`,
/// Java's `Integer.toUnsignedLong`, JS `>>> 0` — need the other reading, which
/// is what `f64.convert_i32_u` performs.
///
/// Stack: `[i32]` → `[n]`.
pub fn emit_as_unsigned32(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::F64_CONVERT_I32_U, line);
}

/// 32-bit rotate with an UNSIGNED result. Stack: `[value, count]` → `[n]`.
pub fn emit_rotl32_unsigned(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_rotate(&mut chunks[current], BitLane::W32, true, line);
    emit_as_unsigned32(&mut chunks[current], line);
}

/// 32-bit rotate right with an UNSIGNED result. Stack: `[value, count]` → `[n]`.
pub fn emit_rotr32_unsigned(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_rotate(&mut chunks[current], BitLane::W32, false, line);
    emit_as_unsigned32(&mut chunks[current], line);
}
