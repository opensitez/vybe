//! I/O compilation — WASI-compatible print, input, file operations.
//!
//! Print uses `wasi:cli/log` (standard WASI import).
//! File I/O uses `wasi:filesystem/*` imports.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
// Target-aware variants can be added here when needed.
// use crate::Target;

/// Emit print/log. Adds import to the given chunk.
/// Use `emit_print_with_import` if your compiler requires imports in chunk 0.
/// Stack: [arg1, arg2, ..., argN] → []
pub fn emit_print(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("wasi:cli", "log");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(arg_count, line);
}

/// Emit print using a pre-resolved import index.
/// Use this when your compiler routes imports to chunk 0 separately.
/// Stack: [arg1, arg2, ..., argN] → []
pub fn emit_print_with_import(chunk: &mut Chunk, import_idx: u16, arg_count: u8, line: u32) {
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, line);
    chunk.emit(arg_count, line);
}

/// Emit readline (input). Adds import to the given chunk.
pub fn emit_input(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:cli", "readLine");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(0, line);
}

/// Emit readline using a pre-resolved import index.
pub fn emit_input_with_import(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, line);
    chunk.emit(0, line);
}

/// Emit print + input (prompt then read). Stack: [prompt_string] → [string]
pub fn emit_prompt_input(chunk: &mut Chunk, line: u32) {
    emit_print(chunk, 1, line);
    chunk.emit_op(Op::DROP, line);
    emit_input(chunk, line);
}

/// Read file contents. Stack: [filename] → [contents_string]
pub fn emit_read_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "readFile");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// Write to file. Stack: [target, data...] → [null]
/// For whole-file: argc=2 (filename, contents). For handle: argc=N (handle, items...).
pub fn emit_write_file(chunk: &mut Chunk, argc: u8, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "writeFile");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}

// ── File handle I/O (wasi:filesystem) ─────────────────────────────────

/// Open file handle. Stack: [filename, mode] → [handle]
pub fn emit_open_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "openFile");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(2, line);
}

/// Close file handle. Stack: [handle] → []
pub fn emit_close_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "closeFile");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// Print to file handle. Stack: [handle, data...] → []
pub fn emit_print_file(chunk: &mut Chunk, argc: u8, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "printFile");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(argc, line);
}


/// Read from file handle (Input). Stack: [handle] → [string]
pub fn emit_input_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "inputFile");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}

/// Read line from file handle. Stack: [handle] → [string]
pub fn emit_line_input(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "lineInput");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(1, line);
}
