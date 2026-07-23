//! .NET `System.IO.StreamReader` / `StreamWriter` adapter — bytecode-only.
//!
//! .NET's text I/O wrappers around an underlying byte stream:
//!   - `StreamReader(path)` — opens file for reading; `ReadLine` / `ReadToEnd`
//!     give buffered text access.
//!   - `StreamWriter(path)` — opens file for writing; `Write` / `WriteLine`
//!     append; `Flush`/`Close` persists.
//!
//! Implementation strategy: load-whole-file model (matches the pre-migration
//! host fns in `vybe_host`). Construction reads the entire file via
//! `node:fs.readFileSync(path, "utf8")` and stashes the string on the
//! reader's `__content` field; subsequent `ReadLine`/`ReadToEnd` walk the
//! cached content. The writer accumulates into `__buf` and flushes via
//! `node:fs.writeFileSync(__path, __buf)` on `Flush`/`Close`.
//!
//! The cleaner streaming model would compose `wasi:io/streams` (read/write
//! 1 byte at a time) but the load-whole-file shape preserves the existing
//! ABI contract — no test exercises true streaming behaviour.
//!
//! Conventions:
//!   * Stack on entry for static ctors: `[arg0, arg1, ...]`
//!   * Stack on entry for instance methods: `[receiver, arg0, ...]`
//!   * Receiver-shape is a plain `Object` with `__type` stamped to
//!     `"StreamReader"` or `"StreamWriter"` plus the per-class fields.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_emitter::instructions::{core_wasm, host};

const TYPE_KEY: &str = "__type";
const CONTENT_KEY: &str = "__content";
const POS_KEY: &str = "__pos";
const PATH_KEY: &str = "__path";
const BUF_KEY: &str = "__buf";

const READER_TYPE: &str = "StreamReader";
const WRITER_TYPE: &str = "StreamWriter";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

/// `new StreamReader(path)` — load file into a `__content` string and
/// initialise `__pos = 0`. Stack: `[path]` → `[reader]`.
pub fn emit_stream_reader_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let read_idx = chunks[current].add_import("node:fs", "readFileSync");
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let content_key = chunk.add_constant(Value::String(Arc::from(CONTENT_KEY)));
    let pos_key = chunk.add_constant(Value::String(Arc::from(POS_KEY)));

    let path_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    // content = node:fs.readFileSync(path, "utf8")
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    push_const(chunk, Value::String(Arc::from("utf8")), line);
    chunk.emit_op_u16(Op::CALL_IMPORT, read_idx, line);
    chunk.emit(2, line);
    let content_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);

    // STRUCT_NEW → [obj]
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);

    // __type = "StreamReader"
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from(READER_TYPE)), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);

    // __content = content
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, content_key, line);
    chunk.emit_op(Op::DROP, line);

    // __pos = 0
    core_wasm::dup(chunk, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::STRUCT_SET, pos_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// `reader.ReadLine()` — return next line up to (excluding) `\n`,
/// advancing `__pos` past it. Returns `null` if `__pos >= len`.
/// Stack: `[reader]` → `[line_or_null]`.
pub fn emit_stream_reader_read_line(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let content_key = chunk.add_constant(Value::String(Arc::from(CONTENT_KEY)));
    let pos_key = chunk.add_constant(Value::String(Arc::from(POS_KEY)));

    let reader_slot = reserve_slot(chunk);
    let content_slot = reserve_slot(chunk);
    let pos_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let end_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    // reader_slot = pop reader
    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);

    // content = reader.__content
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, content_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);

    // pos = reader.__pos
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pos_key, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pos_slot, line);

    // len = wasm:js-string.length(content)
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    // result = null (default for end-of-stream)
    chunk.emit_op(Op::NULL, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let done_block = chunk.emit_block(line);

    // if pos >= len: result stays null → exit
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    vybe_emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    // end = pos
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);

    // Scan loop: advance `end` until `\n` or end-of-string.
    let scan_block = chunk.emit_block(line);
    let (scan_loop, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    vybe_emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line); // exit scan_block

    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
    chunk.emit_i32_const(10, line); // '\n'
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    chunk.emit_br_if(1, line); // newline → exit scan_block

    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(scan_loop);
    chunk.emit_end(line);
    chunk.patch_block(scan_block);

    // result = content.substring(pos, end)
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    // reader.__pos = end + 1 (skip past `\n`); clamp to len.
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::STRUCT_SET, pos_key, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.patch_block(done_block);

    // Push result.
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `reader.ReadToEnd()` — return remaining content from `__pos` to end,
/// advancing `__pos` to end. Stack: `[reader]` → `[remaining]`.
pub fn emit_stream_reader_read_to_end(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let content_key = chunk.add_constant(Value::String(Arc::from(CONTENT_KEY)));
    let pos_key = chunk.add_constant(Value::String(Arc::from(POS_KEY)));

    let reader_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);

    // content = reader.__content
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, content_key, line);

    // pos = reader.__pos (i32)
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pos_key, line);
    chunk.emit_op(Op::I32_FROM_F64, line);

    // end = wasm:js-string.length(content). Need to dup content first since substring consumes it.
    let content_slot = reserve_slot(chunk);
    let pos_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);

    // [content, pos, len] for substring
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);

    // reader.__pos = wasm:js-string.length(content) — fully consumed.
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op_u16(Op::STRUCT_SET, pos_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// `reader.AtEndOfStream` — `__pos >= wasm:js-string.length(__content)`.
/// Stack: `[reader]` → `[bool]`.
pub fn emit_stream_reader_at_end(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let content_key = chunk.add_constant(Value::String(Arc::from(CONTENT_KEY)));
    let pos_key = chunk.add_constant(Value::String(Arc::from(POS_KEY)));

    let reader_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);

    // pos
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, pos_key, line);
    // len
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, content_key, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    // pos < len → DYN_NOT → at-end
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    vybe_emitter::ops::emit_dyn_not(chunk, line);
}

/// `new StreamWriter(path)` — initialise `__path` + empty `__buf`.
/// Stack: `[path]` → `[writer]`.
pub fn emit_stream_writer_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let path_key = chunk.add_constant(Value::String(Arc::from(PATH_KEY)));
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));

    let path_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);

    // __type = "StreamWriter"
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from(WRITER_TYPE)), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);

    // __path = path
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_op_u16(Op::STRUCT_SET, path_key, line);
    chunk.emit_op(Op::DROP, line);

    // __buf = ""
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, buf_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// `writer.Write(s)` — append `s` to `__buf`. Stack: `[writer, s]` → `[null]`.
pub fn emit_stream_writer_write(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));
    let s_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    // [writer] → DUP → [writer, writer] → STRUCT_GET __buf → [writer, buf]
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buf_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, buf_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
}

/// `writer.WriteLine(s)` — append `s + "\n"` to `__buf`.
/// Stack: `[writer, s]` → `[null]`.
pub fn emit_stream_writer_write_line(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));
    let s_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buf_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::STRUCT_SET, buf_key, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
}

/// `writer.Flush()` / `writer.Close()` — `node:fs.writeFileSync(__path, __buf)`.
/// Stack: `[writer]` → `[null]`.
pub fn emit_stream_writer_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    let write_idx = chunks[current].add_import("node:fs", "writeFileSync");
    let chunk = &mut chunks[current];
    let path_key = chunk.add_constant(Value::String(Arc::from(PATH_KEY)));
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));

    let writer_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, writer_slot, line);

    // node:fs.writeFileSync(__path, __buf)
    chunk.emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, path_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buf_key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, write_idx, line);
    chunk.emit(2, line);
    chunk.emit_op(Op::DROP, line); // discard writeFileSync result

    chunk.emit_op(Op::NULL, line);
}

/// `reader.Close()` / `writer.Close()` — no-op for readers, flush for writers.
/// Stack: `[stream]` → `[null]`.
pub fn emit_stream_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let write_idx = chunks[current].add_import("node:fs", "writeFileSync");
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let path_key = chunk.add_constant(Value::String(Arc::from(PATH_KEY)));
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));

    let stream_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, stream_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, type_key, line);
    push_const(chunk, Value::String(Arc::from(WRITER_TYPE)), line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);

    let skip_flush = chunk.emit_block(line);
    vybe_emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, path_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, buf_key, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, write_idx, line);
    chunk.emit(2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.patch_block(skip_flush);
    chunk.emit_op(Op::NULL, line);
}

/// `reader.Close()` — no-op (no resource to release in load-whole-file model).
/// Stack: `[reader]` → `[null]`.
pub fn emit_stream_reader_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line); // drop receiver
    chunk.emit_op(Op::NULL, line);
}
