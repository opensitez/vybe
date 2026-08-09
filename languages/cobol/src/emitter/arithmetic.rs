use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use super::support::stash_args;

pub fn emit_round_away_from_zero(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash_args(chunks, current, 1, line);
    let value_slot = base;

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::F64_CEIL, line);
    chunks[current].emit_end(line);
}
