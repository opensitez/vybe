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
    chunk.emit_call(idx, arg_count, line);
}

/// Emit print using a pre-resolved import index.
pub fn emit_print_with_import(chunk: &mut Chunk, import_idx: u16, arg_count: u8, line: u32) {
    chunk.emit_call(import_idx, arg_count, line);
}

/// Emit a raw byte write to stdout — NO implicit newline, unlike
/// `wasi:logging/logging.log` which is one line-oriented record per call.
///
/// Composes the WASI 0.3 surface: `canon stream.new` (STREAM_NEW) creates
/// a `stream<u8>` as (readable, writable) i32 handles, the contents go in
/// via STREAM_WRITE, the writable end closes (EOF), and the readable end
/// is passed to `wasi:cli/stdout.write-via-stream(data: stream<u8>)`.
/// The returned `future<result<_, error-code>>` is discarded; both handle
/// table entries are dropped afterwards per the canonical ABI.
///
/// Stack: [] → []. `push_contents` emits the string to write while the
/// writable handle is on the stack. `rd_slot`/`wr_slot` are caller-defined
/// scratch locals for the two canon handles; `write_idx` is the resolved
/// `wasi:cli/stdout.write-via-stream` import.
pub fn emit_write_stdout_with_imports(
    chunk: &mut Chunk,
    write_idx: u16,
    rd_slot: u16,
    wr_slot: u16,
    line: u32,
    push_contents: impl FnOnce(&mut Chunk),
) {
    // canon stream.new → [rd, wr]
    chunk.emit_op(Op::STREAM_NEW, line);
    chunk.emit_op_u16(Op::LOCAL_SET, wr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, rd_slot, line);
    // stream.write(wr, contents)
    chunk.emit_op_u16(Op::LOCAL_GET, wr_slot, line);
    push_contents(chunk);
    chunk.emit_op(Op::STREAM_WRITE, line);
    // stream.drop-writable(wr) — signals EOF
    chunk.emit_op_u16(Op::LOCAL_GET, wr_slot, line);
    chunk.emit_op(Op::STREAM_DROP_WR, line);
    // wasi:cli/stdout.write-via-stream(rd) → future (discard)
    chunk.emit_op_u16(Op::LOCAL_GET, rd_slot, line);
    chunk.emit_call(write_idx, 1, line);
    chunk.emit_op(Op::DROP, line);
    // stream.drop-readable(rd)
    chunk.emit_op_u16(Op::LOCAL_GET, rd_slot, line);
    chunk.emit_op(Op::STREAM_DROP_RD, line);
}

/// Emit print to stderr (warn/error level). Stack: [message] → []
pub fn emit_print_error(chunk: &mut Chunk, line: u32) {
    chunk.emit_string_const("error", line);
    let idx = chunk.add_import("wasi:logging/logging", "log");
    chunk.emit_call(idx, 2, line); // (level, message)
}

/// Emit readline. Stack: [] → [string]
/// Calls wasi:cli/stdin.get-stdin → [method]input-stream.blocking-read(stream, 65536).
/// Host returns a string for fd=0 (stdin line input).
pub fn emit_input(chunk: &mut Chunk, line: u32) {
    let stdin_idx = chunk.add_import("wasi:cli/stdin", "get-stdin");
    chunk.emit_call(stdin_idx, 0, line);
    chunk.emit_i32_const(65536, line);
    let read_idx = chunk.add_import("wasi:io/streams", "[method]input-stream.blocking-read");
    chunk.emit_call(read_idx, 2, line);
}

/// Emit readline using a pre-resolved import index.
pub fn emit_input_with_import(chunk: &mut Chunk, import_idx: u16, line: u32) {
    chunk.emit_call(import_idx, 0, line);
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
    chunk.emit_call(idx, 1, line);
}

/// Write to file. Stack: [target, data...] → [null]
pub fn emit_write_file(chunk: &mut Chunk, argc: u8, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "writeFile");
    chunk.emit_call(idx, argc, line);
}

// ── File handle I/O (wasi:filesystem) ─────────────────────────────────

/// Open file handle. Stack: [filename, mode] → [handle]
pub fn emit_open_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "openFile");
    chunk.emit_call(idx, 2, line);
}

/// Close file handle. Stack: [handle] → []
pub fn emit_close_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "closeFile");
    chunk.emit_call(idx, 1, line);
}

/// Print to file handle. Stack: [handle, data...] → []
pub fn emit_print_file(chunk: &mut Chunk, argc: u8, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "printFile");
    chunk.emit_call(idx, argc, line);
}

/// Read from file handle (Input). Stack: [handle] → [string]
pub fn emit_input_file(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "inputFile");
    chunk.emit_call(idx, 1, line);
}

/// Read line from file handle. Stack: [handle] → [string]
pub fn emit_line_input(chunk: &mut Chunk, line: u32) {
    let idx = chunk.add_import("wasi:filesystem", "lineInput");
    chunk.emit_call(idx, 1, line);
}
