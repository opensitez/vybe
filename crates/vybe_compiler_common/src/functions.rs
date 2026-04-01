//! Function compilation helpers — shared bytecode patterns for parameters and calls.
//!
//! All compilers emit the same pattern for default parameter values:
//! check if param is null → substitute default → patch jump.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Emit the start of a default parameter check.
/// If the parameter at `param_slot` is null, the caller should compile the default
/// expression, then call `emit_default_param_end`.
/// Returns a jump offset to patch.
///
/// Stack: unchanged
pub fn emit_default_param_start(chunk: &mut Chunk, param_slot: u16, line: u32) -> usize {
    chunk.emit_op_u16(Op::local_get, param_slot, line);
    chunk.emit_op(Op::ref_is_null, line);
    chunk.emit_jump(Op::br_if_false, line)
}

/// Emit the end of a default parameter check.
/// Caller must have compiled the default expression onto the stack.
/// Stack before: [default_value]  Stack after: [] (stored in param_slot)
pub fn emit_default_param_end(chunk: &mut Chunk, param_slot: u16, skip_jump: usize, line: u32) {
    chunk.emit_op_u16(Op::local_set, param_slot, line);
    chunk.emit_op(Op::drop, line);
    chunk.patch_jump(skip_jump);
}
