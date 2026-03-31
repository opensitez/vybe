//! I/O compilation — WASI-compatible print, input, file operations.
//!
//! Print uses `wasi:cli/log` (standard WASI import).
//! File I/O uses `wasi:filesystem/*` imports.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
// Target-aware variants can be added here when needed.
// use crate::Target;

/// Emit print/log. Stack: [arg1, arg2, ..., argN] → []
/// `arg_count` values are consumed from the stack.
pub fn emit_print(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("wasi:cli", "log");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(arg_count, line);
}

/// Emit readline (input). Stack: [] → [string]
pub fn emit_input(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:cli", "readLine");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(0, line);
}

/// Emit print + input (prompt then read). Stack: [prompt_string] → [string]
pub fn emit_prompt_input(chunk: &mut Chunk, line: u32) {
    emit_print(chunk, 1, line);
    chunk.emit_op(Op::drop, line);
    emit_input(chunk, line);
}

/// Read file contents. Stack: [filename] → [contents_string]
pub fn emit_read_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "readFile");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(1, line);
}

/// Write file contents. Stack: [filename, contents] → [null]
pub fn emit_write_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "writeFile");
    chunk.emit_op_u16(Op::call_import, idx, line);
    chunk.emit(2, line);
}
