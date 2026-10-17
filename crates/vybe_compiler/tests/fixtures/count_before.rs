use vybe_ast::BitLane;
use vybe_runtime::{Chunk, Op};
fn narrow(chunk: &mut Chunk, lane: BitLane, line: u32) {
    match lane {
        BitLane::W32 => {
            chunk.emit_op(Op::I64_TRUNC_F64_S, line);
            chunk.emit_op(Op::I32_WRAP_I64, line);
        }
        BitLane::W64 => chunk.emit_op(Op::I64_TRUNC_F64_S, line),
    }
}
fn widen(chunk: &mut Chunk, lane: BitLane, line: u32) {
    if lane == BitLane::W64 {
        chunk.emit_op(Op::F64_CONVERT_I64_S, line);
    }
}
pub fn emit_count(chunk: &mut Chunk, lane: BitLane, w32: Op, w64: Op, line: u32) {
    narrow(chunk, lane, line);
    chunk.emit_op(if lane == BitLane::W32 { w32 } else { w64 }, line);
    widen(chunk, lane, line);
}
