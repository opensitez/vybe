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
use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::object::emit_bind_method_with_slot;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const TYPE_KEY: &str = "__type";
const CONTENT_KEY: &str = "__content";
const POS_KEY: &str = "__pos";
const PATH_KEY: &str = "__path";
const BUF_KEY: &str = "__buf";
const BUILDER_KEY: &str = "__builder";
const SB_BUFFER_KEY: &str = "__buffer";
const DISPOSED_KEY: &str = "__disposed";

const READER_TYPE: &str = "StreamReader";
const WRITER_TYPE: &str = "StreamWriter";
const STRING_READER_TYPE: &str = "StringReader";
const STRING_WRITER_TYPE: &str = "StringWriter";

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

fn emit_throw_object_disposed(chunk: &mut Chunk, line: u32) {
    chunk.emit_struct_new(0, 0, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const("Cannot read from a closed TextReader.", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        chunk,
        "ObjectDisposedException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

fn emit_throw_if_disposed(chunk: &mut Chunk, reader_slot: u16, line: u32) {
    let disposed_key = chunk.add_constant(Value::String(Arc::from(DISPOSED_KEY)));
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, disposed_key, line);
    chunk.emit_if_value(line);
    emit_throw_object_disposed(chunk, line);
    chunk.emit_end(line);
}

fn bind_string_writer_to_string(
    chunks: &mut Vec<Chunk>,
    current: usize,
    this_slot: u16,
    line: u32,
) {
    let mut method = create_function_chunk("__string_writer_tostring", 1);
    let builder_key = method.add_constant(Value::String(Arc::from(BUILDER_KEY)));
    let sb_buffer_key = method.add_constant(Value::String(Arc::from(SB_BUFFER_KEY)));
    method.emit_op_u16(Op::LOCAL_GET, 0, line);
    method.emit_struct_field_op(Op::STRUCT_GET, 0, builder_key, line);
    method.emit_struct_field_op(Op::STRUCT_GET, 0, sb_buffer_key, line);
    method.emit_op(Op::RETURN, line);
    method.local_count = 1;
    chunks.push(method);
    let method_idx = chunks.len() - 1;

    for name in ["tostring", "ToString"] {
        emit_bind_method_with_slot(
            &mut chunks[current],
            this_slot,
            name,
            Some(vybe_ast::ProtocolSlot::ToString),
            method_idx,
            None,
            line,
        );
    }
}

fn emit_set_writer_buffer(chunk: &mut Chunk, writer_slot: u16, new_buf_slot: u16, line: u32) {
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));
    let builder_key = chunk.add_constant(Value::String(Arc::from(BUILDER_KEY)));
    let sb_buffer_key = chunk.add_constant(Value::String(Arc::from(SB_BUFFER_KEY)));
    let builder_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, new_buf_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, buf_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, builder_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, builder_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, builder_slot, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, builder_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, new_buf_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, sb_buffer_key, line);
    chunk.emit_end(line);
}

/// `new StreamReader(path)` — load file into a `__content` string and
/// initialise `__pos = 0`. Stack: `[path]` → `[reader]`.
pub fn emit_stream_reader_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let content_key = chunk.add_constant(Value::String(Arc::from(CONTENT_KEY)));
    let pos_key = chunk.add_constant(Value::String(Arc::from(POS_KEY)));
    let disposed_key = chunk.add_constant(Value::String(Arc::from(DISPOSED_KEY)));

    let path_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    // content = the file, read through `read-via-stream` and decoded as UTF-8.
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    vybe_compiler::primitives::fs_path::emit_read_file(chunk, line);
    let content_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);

    // STRUCT_NEW → [obj]
    chunk.emit_struct_new(0, 0, line);

    // __type = "StreamReader"
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from(READER_TYPE)), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);

    // __content = content
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, content_key, line);

    // __pos = 0
    core_wasm::dup(chunk, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, pos_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, disposed_key, line);
}

/// `new StringReader(text)` — cache the supplied string and initialise
/// `__pos = 0`. Stack: `[text]` → `[reader]`.
pub fn emit_string_reader_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let content_key = chunk.add_constant(Value::String(Arc::from(CONTENT_KEY)));
    let pos_key = chunk.add_constant(Value::String(Arc::from(POS_KEY)));
    let disposed_key = chunk.add_constant(Value::String(Arc::from(DISPOSED_KEY)));

    let content_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);
    chunk.emit_struct_new(0, 0, line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from(STRING_READER_TYPE)), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, content_key, line);

    core_wasm::dup(chunk, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, pos_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, disposed_key, line);
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
    let result_end_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    // reader_slot = pop reader
    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);
    emit_throw_if_disposed(chunk, reader_slot, line);

    // content = reader.__content
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, content_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);

    // pos = reader.__pos
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pos_key, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pos_slot, line);

    // len = wasm:js-string.length(content)
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    // result = null (default for end-of-stream)
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let done_block = chunk.emit_block(line);

    // if pos >= len: result stays null → exit
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    // end = pos
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);

    // Scan loop: advance `end` until `\n` or end-of-string.
    let scan_block = chunk.emit_block(line);
    let (scan_loop, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line); // exit scan_block

    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
    chunk.emit_i32_const(10, line); // '\n'
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
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

    // result_end = end, except CRLF returns the line without the '\r'.
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_end_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
    chunk.emit_i32_const(13, line); // '\r'
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_SUB, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_end_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // result = content.substring(pos, result_end)
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_end_slot, line);
    host::emit(chunk, "wasm:js-string", "substring", 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    // reader.__pos = end + 1 (skip past `\n`); clamp to len.
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, pos_key, line);

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
    emit_throw_if_disposed(chunk, reader_slot, line);

    // content = reader.__content
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, content_key, line);

    // pos = reader.__pos (i32)
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pos_key, line);
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
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, pos_key, line);
}

/// `reader.AtEndOfStream` — `__pos >= wasm:js-string.length(__content)`.
/// Stack: `[reader]` → `[bool]`.
pub fn emit_stream_reader_at_end(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let content_key = chunk.add_constant(Value::String(Arc::from(CONTENT_KEY)));
    let pos_key = chunk.add_constant(Value::String(Arc::from(POS_KEY)));

    let reader_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);
    emit_throw_if_disposed(chunk, reader_slot, line);

    // pos
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pos_key, line);
    // len
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, content_key, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    // pos < len → DYN_NOT → at-end
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `new StreamWriter(path)` — initialise `__path` + empty `__buf`.
/// Stack: `[path]` → `[writer]`.
pub fn emit_stream_writer_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let path_key = chunk.add_constant(Value::String(Arc::from(PATH_KEY)));
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));
    let nl_key = chunk.add_constant(Value::String(Arc::from("NewLine")));
    let nl_lower_key = chunk.add_constant(Value::String(Arc::from("newline")));

    let path_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    chunk.emit_struct_new(0, 0, line);

    // __type = "StreamWriter"
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from(WRITER_TYPE)), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);

    // __path = path
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, path_key, line);

    // __buf = ""
    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("")), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, buf_key, line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, nl_key, line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, nl_lower_key, line);
}

/// `new StringWriter()` / `new StringWriter(StringBuilder)` — initialise an
/// in-memory buffer. Stack: `[]` or `[builder]` → `[writer]`.
pub fn emit_string_writer_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));
    let builder_key = chunk.add_constant(Value::String(Arc::from(BUILDER_KEY)));
    let sb_type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let sb_buffer_key = chunk.add_constant(Value::String(Arc::from(SB_BUFFER_KEY)));
    let nl_key = chunk.add_constant(Value::String(Arc::from("NewLine")));
    let nl_lower_key = chunk.add_constant(Value::String(Arc::from("newline")));
    let encoding_key = chunk.add_constant(Value::String(Arc::from("Encoding")));
    let encoding_lower_key = chunk.add_constant(Value::String(Arc::from("encoding")));
    let web_name_key = chunk.add_constant(Value::String(Arc::from("WebName")));
    let web_name_lower_key = chunk.add_constant(Value::String(Arc::from("webname")));

    let initial_slot = reserve_slot(chunk);
    let builder_slot = reserve_slot(chunk);
    if argc > 0 {
        chunk.emit_op_u16(Op::LOCAL_SET, builder_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, builder_slot, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, 0, sb_buffer_key, line);
        chunk.emit_op_u16(Op::LOCAL_SET, initial_slot, line);
    } else {
        push_const(chunk, Value::String(Arc::from("")), line);
        chunk.emit_op_u16(Op::LOCAL_SET, initial_slot, line);

        chunk.emit_struct_new(0, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, builder_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, builder_slot, line);
        push_const(chunk, Value::String(Arc::from("StringBuilder")), line);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, sb_type_key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, builder_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, initial_slot, line);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, sb_buffer_key, line);
    }

    chunk.emit_struct_new(0, 0, line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from(STRING_WRITER_TYPE)), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, type_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, initial_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, buf_key, line);

    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, builder_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, builder_key, line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, nl_key, line);

    core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("\n")), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, nl_lower_key, line);

    let encoding_slot = reserve_slot(chunk);
    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, encoding_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, encoding_slot, line);
    push_const(chunk, Value::String(Arc::from("utf-16")), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, web_name_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, encoding_slot, line);
    push_const(chunk, Value::String(Arc::from("utf-16")), line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, web_name_lower_key, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, encoding_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, encoding_key, line);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, encoding_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, encoding_lower_key, line);

    let obj_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    let _ = chunk;

    bind_string_writer_to_string(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `writer.Write(s)` — append `s` to `__buf`. Stack: `[writer, s]` → `[null]`.
pub fn emit_stream_writer_write(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));
    let s_slot = reserve_slot(chunk);
    let writer_slot = reserve_slot(chunk);
    let new_buf_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, writer_slot, line);
    let left_slot = reserve_slot(chunk);
    let right_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    super::console_adapter::emit_dotnet_stringify(chunk, left_slot, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    super::console_adapter::emit_dotnet_stringify(chunk, right_slot, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, new_buf_slot, line);
    emit_set_writer_buffer(chunk, writer_slot, new_buf_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `writer.Write(fmt, a, b)` or `writer.Write(chars, index, count)`.
/// Stack: `[writer, arg0, arg1, arg2]` → `[null]`.
pub fn emit_stream_writer_write_3(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let a2_slot = reserve_slot(chunk);
    let a1_slot = reserve_slot(chunk);
    let a0_slot = reserve_slot(chunk);
    let writer_slot = reserve_slot(chunk);
    let text_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, a2_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a1_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a0_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, writer_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, a0_slot, line);
    host::emit(chunk, "ecma:value", "typeof", 1, line);
    push_const(chunk, Value::String(Arc::from("string")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, a0_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a1_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a2_slot, line);
    super::string_format_adapter::emit_string_format(chunks, current, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a0_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a1_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a1_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a2_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    host::emit(&mut chunks[current], "ecma:array", "slice", 3, line);
    push_const(&mut chunks[current], Value::String(Arc::from("")), line);
    host::emit(&mut chunks[current], "ecma:array", "join", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    emit_stream_writer_write(chunks, current, line);
}

/// `writer.WriteLine(s)` — append `s + "\n"` to `__buf`.
/// Stack: `[writer, s]` → `[null]`.
pub fn emit_stream_writer_write_line(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));
    let nl_key = chunk.add_constant(Value::String(Arc::from("newline")));
    let s_slot = reserve_slot(chunk);
    let writer_slot = reserve_slot(chunk);
    let new_buf_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, writer_slot, line);
    let left_slot = reserve_slot(chunk);
    let right_slot = reserve_slot(chunk);
    let nl_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);
    super::console_adapter::emit_dotnet_stringify(chunk, left_slot, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    super::console_adapter::emit_dotnet_stringify(chunk, right_slot, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, nl_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, nl_slot, line);
    super::console_adapter::emit_dotnet_stringify(chunk, nl_slot, nl_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, nl_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, new_buf_slot, line);
    emit_set_writer_buffer(chunk, writer_slot, new_buf_slot, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_stream_writer_write_line_async(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_stream_writer_write_line(chunks, current, line);
    host::emit(&mut chunks[current], "ecma:promise", "resolve", 1, line);
}

/// `StringWriter.ToString()` — return accumulated buffer.
pub fn emit_string_writer_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let builder_key = chunk.add_constant(Value::String(Arc::from(BUILDER_KEY)));
    let sb_buffer_key = chunk.add_constant(Value::String(Arc::from(SB_BUFFER_KEY)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, builder_key, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, sb_buffer_key, line);
}

pub fn emit_string_writer_get_string_builder(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let builder_key = chunk.add_constant(Value::String(Arc::from(BUILDER_KEY)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, builder_key, line);
}

/// `StringWriter.Flush/Close/Dispose()` — no-op for in-memory writers.
pub fn emit_string_writer_noop(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `StringReader.Peek()` — next UTF-16 code unit or -1 at end.
pub fn emit_string_reader_peek(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_string_reader_read_char(chunks, current, line, false);
}

/// `StringReader.Read()` — next UTF-16 code unit and advance, or -1.
pub fn emit_string_reader_read(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_string_reader_read_char(chunks, current, line, true);
}

/// `StringReader.Read(buffer, index, count)` / `ReadBlock(...)`.
/// Stack: `[reader, buffer, index, count]` -> `[read_count]`.
pub fn emit_string_reader_read_buffer(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let content_key = chunk.add_constant(Value::String(Arc::from(CONTENT_KEY)));
    let pos_key = chunk.add_constant(Value::String(Arc::from(POS_KEY)));
    let reader_slot = reserve_slot(chunk);
    let buffer_slot = reserve_slot(chunk);
    let index_slot = reserve_slot(chunk);
    let count_slot = reserve_slot(chunk);
    let content_slot = reserve_slot(chunk);
    let pos_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let read_slot = reserve_slot(chunk);
    let ch_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buffer_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);
    emit_throw_if_disposed(chunk, reader_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, content_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pos_key, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pos_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, read_slot, line);

    let outer = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, read_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, read_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, read_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
    host::emit(chunk, "wasm:js-string", "fromCharCode", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ch_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, buffer_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, read_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, ch_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);

    chunk.emit_op_u16(Op::LOCAL_GET, read_slot, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, read_slot, line);
    chunk.emit_br(0, line);

    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, read_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, pos_key, line);

    chunk.emit_op_u16(Op::LOCAL_GET, read_slot, line);
}

fn emit_string_reader_read_char(chunks: &mut [Chunk], current: usize, line: u32, advance: bool) {
    let chunk = &mut chunks[current];
    let content_key = chunk.add_constant(Value::String(Arc::from(CONTENT_KEY)));
    let pos_key = chunk.add_constant(Value::String(Arc::from(POS_KEY)));
    let reader_slot = reserve_slot(chunk);
    let content_slot = reserve_slot(chunk);
    let pos_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);
    emit_throw_if_disposed(chunk, reader_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, content_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, pos_key, line);
    chunk.emit_op(Op::I32_FROM_F64, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    if advance {
        chunk.emit_op_u16(Op::LOCAL_GET, reader_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_struct_field_op(Op::STRUCT_SET, 0, pos_key, line);
    }
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `writer.Flush()` / `writer.Close()` — the buffered text through
/// `write-via-stream`. Stack: `[writer]` → `[null]`.
pub fn emit_stream_writer_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let path_key = chunk.add_constant(Value::String(Arc::from(PATH_KEY)));
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));

    let writer_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, writer_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, path_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, writer_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    vybe_compiler::primitives::fs_path::emit_write_file(chunk, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `reader.Close()` / `writer.Close()` — no-op for readers, flush for writers.
/// Stack: `[stream]` → `[null]`.
pub fn emit_stream_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let path_key = chunk.add_constant(Value::String(Arc::from(PATH_KEY)));
    let buf_key = chunk.add_constant(Value::String(Arc::from(BUF_KEY)));

    let stream_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, stream_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, type_key, line);
    push_const(chunk, Value::String(Arc::from(WRITER_TYPE)), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);

    let skip_flush = chunk.emit_block(line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, path_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, stream_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, buf_key, line);
    vybe_compiler::primitives::fs_path::emit_write_file(chunk, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.patch_block(skip_flush);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `reader.Close()` — no-op (no resource to release in load-whole-file model).
/// Stack: `[reader]` → `[null]`.
pub fn emit_stream_reader_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let disposed_key = chunk.add_constant(Value::String(Arc::from(DISPOSED_KEY)));
    chunk.emit_bool_const(true, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, disposed_key, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
