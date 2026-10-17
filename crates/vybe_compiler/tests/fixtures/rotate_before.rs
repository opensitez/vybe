use vybe_ast::BitLane;
use vybe_runtime::{Chunk, Op};
fn narrow(chunk: &mut Chunk, lane: BitLane, line: u32) {
    chunk.emit_op(Op::I64_TRUNC_F64_S, line);
    if lane == BitLane::W32 {
        chunk.emit_op(Op::I32_WRAP_I64, line);
    }
}
fn widen(chunk: &mut Chunk, lane: BitLane, line: u32) {
    if lane == BitLane::W64 {
        chunk.emit_op(Op::F64_CONVERT_I64_S, line);
    }
}
fn raw_i64_scratch(chunk: &mut Chunk) -> u16 {
    if chunk.i64_scratch_slot.is_none() {
        chunk.local_count = chunk.local_count.max(chunk.scratch_high_water);
    }
    chunk.i64_scratch()
}
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
    let base = chunk.alloc_scratch(2);
    // The count is on top, so it pops first.
    if lane != BitLane::W32 || !normalized.1 {
        narrow(chunk, lane, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, base + 1, line);
    if lane != BitLane::W32 || !normalized.0 {
        narrow(chunk, lane, line);
    }
    chunk.emit_op_u16(Op::LOCAL_TEE, base, line);
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
