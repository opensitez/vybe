use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

use super::support::stash_args;

pub fn emit_round_away_from_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let value_slot = base;

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::F64_CONST_0, line);
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    let non_negative = chunks[current].emit_jump(Op::BR_IF_FALSE, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    let done = chunks[current].emit_jump(Op::BR, line);

    chunks[current].patch_jump(non_negative);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::F64_CEIL, line);
    chunks[current].patch_jump(done);
}