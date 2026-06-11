//! I/O compilation — WASI-compatible print, input, file operations.
//!
//! Print uses `wasi:logging/logging.log` (WASI logging proposal).
//! Input uses `wasi:cli/stdin.get-stdin` + `[method]input-stream.blocking-read`.
//! File I/O uses `wasi:filesystem/*` imports.

use vybe_bytecode::Chunk;
use vybe_bytecode::Value;
use vybe_bytecode::opcode::Op;

/// Emit print/log. Stack: [arg1, ..., argN] → []
/// Routes to wasi:logging/logging.log. Flexible arity:
///   1 arg → host treats as (info, "", message)
///   N args → host joins them (info, "", joined)
pub fn emit_print(chunk: &mut Chunk, arg_count: u8, line: u32) {
    let idx = chunk.add_import("wasi:logging/logging", "log");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(arg_count, line);
}

/// Emit print using a pre-resolved import index.
pub fn emit_print_with_import(chunk: &mut Chunk, import_idx: u16, arg_count: u8, line: u32) {
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, line);
    chunk.emit(arg_count, line);
}

/// Emit print to stderr (warn/error level). Stack: [message] → []
pub fn emit_print_error(chunk: &mut Chunk, line: u32) {
    let level_idx = chunk.add_constant(Value::String(std::sync::Arc::from("error")));
    chunk.emit_op_u16(Op::CONST, level_idx, line);
    let idx = chunk.add_import("wasi:logging/logging", "log");
    chunk.emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunk.emit(2, line); // (level, message)
}

/// Emit readline. Stack: [] → [string]
/// Calls wasi:cli/stdin.get-stdin → [method]input-stream.blocking-read(stream, 65536).
/// Host returns a string for fd=0 (stdin line input).
pub fn emit_input(chunk: &mut Chunk, line: u32) {
    let stdin_idx = chunk.add_import("wasi:cli/stdin", "get-stdin");
    chunk.emit_op_u16(Op::CALL_IMPORT, stdin_idx, line);
    chunk.emit(0, line);
    let max_idx = chunk.add_constant(Value::I32(65536));
    chunk.emit_op_u16(Op::CONST, max_idx, line);
    let read_idx = chunk.add_import("wasi:io/streams", "[method]input-stream.blocking-read");
    chunk.emit_op_u16(Op::CALL_IMPORT, read_idx, line);
    chunk.emit(2, line);
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
