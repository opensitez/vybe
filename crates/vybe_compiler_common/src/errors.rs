//! Exception handling helpers — shared try/catch/finally bytecode patterns.
//!
//! All compilers emit the same opcodes for exception handling:
//! - try_start → body → try_end → handler
//! - try_table for typed multi-catch

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// Emit the start of a try block. Returns the catch_jump offset to patch later.
/// Stack: unchanged
pub fn emit_try_start(chunk: &mut Chunk, line: u32) -> usize {
    let catch_jump = chunk.emit_jump(Op::try_start, line);
    chunk.emit(0u8, line); // reserved for finally offset
    catch_jump
}

/// Emit the end of the try body (normal exit path).
pub fn emit_try_end(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::try_end, line);
}

/// Patch the catch handler offset after the handler code has been emitted.
pub fn patch_catch(chunk: &mut Chunk, catch_jump: usize) {
    chunk.patch_jump(catch_jump);
}

/// Emit a throw — takes the exception value from TOS.
/// Stack before: [exception_value]  Stack after: diverges
pub fn emit_throw(chunk: &mut Chunk, line: u32) {
    chunk.emit_op(Op::throw, line);
}
