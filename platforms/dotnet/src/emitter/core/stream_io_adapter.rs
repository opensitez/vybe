//! .NET `System.IO.StreamReader` / `StreamWriter` adapter — bytecode-only.
//!
//! .NET's text I/O wrappers around an underlying byte stream:
//!   - `StreamReader(path)` — opens file for reading; `ReadLine` / `ReadToEnd`
//!     give buffered text access.
//!   - `StreamWriter(path)` — opens file for writing; `Write` / `WriteLine`
//!     append; `Flush`/`Close` persists.
//!
//! Implementation strategy: load-whole-file model. Construction reads the
//! entire file through `fs_path::emit_read_file` — `open-at` +
//! `read-via-stream` + `canon stream.read`, decoded as UTF-8 — and stashes the
//! string on the reader's `__content` field; subsequent `ReadLine`/`ReadToEnd`
//! walk the cached content. The writer accumulates into `__buf` and flushes
//! through `fs_path::emit_write_file` on `Flush`/`Close`.
//!
//! ⚠It is still LOAD-WHOLE-FILE, and that is the remaining gap: a `StreamReader`
//! over a 2GB file materialises 2GB. The transport underneath is now a real
//! Component Model stream, so incremental reads are expressible — what is
//! missing is a reader that keeps the stream handle alive across `ReadLine`
//! calls, and a stream end is an index into the VM's handle table, so that is
//! a lifetime question rather than a call.
//!
//! It used to say the cleaner model would compose `wasi:io/streams`. That
//! package does not exist in WASI 0.3.1 — streams became Component Model
//! built-ins — so the sentence outlived the thing it described.
//!
//! Conventions:
//!   * Stack on entry for static ctors: `[arg0, arg1, ...]`
//!   * Stack on entry for instance methods: `[receiver, arg0, ...]`
//!   * Receiver-shape is a plain `Object` with `__type` stamped to
//!     `"StreamReader"` or `"StreamWriter"` plus the per-class fields.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::object::emit_bind_method_with_slot;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

const TYPE_KEY: &str = "__type";
const CONTENT_KEY: &str = "__content";
const POS_KEY: &str = "__pos";
const PATH_KEY: &str = "__path";
const BUF_KEY: &str = "__buf";
const BUILDER_KEY: &str = "__builder";
const SB_BUFFER_KEY: &str = "__buffer";
const DISPOSED_KEY: &str = "__disposed";
/// The underlying byte stream, when the reader/writer was built over one
/// (`new StreamReader(ms)`) rather than over a file path.
const STREAM_KEY: &str = "__stream";

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
    vybe_compiler::primitives::errors::emit_exception_new(
        chunk,
        "ObjectDisposedException",
        class_slots::ValueSource::ConstStr("Cannot read from a closed TextReader.".to_string()),
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

fn emit_throw_if_disposed(chunk: &mut Chunk, reader_slot: u16, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(DISPOSED_KEY),
        Dest::Stack,
        line,
    );
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
    class_slots::emit_class_get(
        &mut method,
        ObjSource::Local(0),
        &field_slot(BUILDER_KEY),
        Dest::Stack,
        line,
    );
    class_slots::emit_class_get(
        &mut method,
        ObjSource::Stack,
        &field_slot(SB_BUFFER_KEY),
        Dest::Stack,
        line,
    );
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
    let builder_slot = reserve_slot(chunk);

    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(writer_slot),
        &field_slot(BUF_KEY),
        ValueSource::Local(new_buf_slot),
        line,
    );

    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(writer_slot),
        &field_slot(BUILDER_KEY),
        Dest::Local(builder_slot),
        line,
    );

    chunk.emit_op_u16(Op::LOCAL_GET, builder_slot, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(builder_slot),
        &field_slot(SB_BUFFER_KEY),
        ValueSource::Local(new_buf_slot),
        line,
    );
    chunk.emit_end(line);
}

/// The stream fields a `MemoryStream` carries — read directly, because a text
/// reader over a stream needs the bytes and there is no name-dispatch here.
const MS_BUF: &str = "__ms_buf";
const MS_POS: &str = "__ms_pos";
const MS_LEN: &str = "__ms_len";

/// `[]` → `[text]` — the bytes from `__ms_pos` to `__ms_len` of the stream in
/// `stream_slot`, decoded as UTF-8, with the stream's cursor left at the end.
///
/// ⛔ A `StreamReader` over a stream reads it EAGERLY here, the same
/// load-whole-file model the path constructor uses; the cursor moves so a
/// second reader over the same stream sees nothing left, which is what .NET's
/// buffering does in practice.
fn emit_decode_stream_tail(chunk: &mut Chunk, stream_slot: u16, line: u32) {
    // decoder first: `decode` is receiver-first.
    host::emit(chunk, "web:encoding", "decoderNew", 0, line);

    for key in [MS_BUF, MS_POS, MS_LEN] {
        class_slots::emit_class_get(
            chunk,
            ObjSource::Local(stream_slot),
            &field_slot(key),
            Dest::Stack,
            line,
        );
    }
    host::emit(chunk, "ecma:array", "slice", 3, line);
    // ⛔ `decode` takes a BufferSource: a plain Array decodes to the EMPTY
    // STRING rather than failing.
    host::emit(chunk, "ecma:uint8array", "newFromIterable", 1, line);
    host::emit(chunk, "web:encoding", "decode", 2, line);

    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(stream_slot),
        &field_slot(MS_LEN),
        Dest::Stack,
        line,
    );
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(stream_slot),
        &field_slot(MS_POS),
        ValueSource::Stack,
        line,
    );
}

/// `[…, argN]` → `[arg0]` — drop the trailing constructor arguments .NET
/// accepts and this model does not need (encoding, buffer size, `leaveOpen`).
fn emit_drop_extra_args(chunk: &mut Chunk, argc: u8, line: u32) {
    for _ in 1..argc.max(1) {
        chunk.emit_op(Op::DROP, line);
    }
}

/// `new StreamReader(path)` / `new StreamReader(stream)` — load the source
/// into a `__content` string and initialise `__pos = 0`.
/// Stack: `[source, …]` → `[reader]`.
pub fn emit_stream_reader_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    emit_drop_extra_args(chunk, argc, line);
    let path_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    // content = the SOURCE as text. A string is a file path, read through
    // `read-via-stream`; anything else is a byte stream, decoded from its
    // cursor to its length.
    let content_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    host::emit(chunk, "wasm:js-string", "test", 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    vybe_compiler::primitives::fs_path::emit_read_file(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);
    chunk.emit_else(line);
    emit_decode_stream_tail(chunk, path_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, content_slot, line);
    chunk.emit_end(line);

    class_slots::emit_class_construct(
        chunk,
        READER_TYPE,
        &[
            (field_slot(CONTENT_KEY), ValueSource::Local(content_slot)),
            (field_slot(POS_KEY), ValueSource::ConstI32(0)),
            (field_slot(DISPOSED_KEY), ValueSource::ConstBool(false)),
        ],
        line,
    );
}

/// `new StringReader(text)` — cache the supplied string and initialise
/// `__pos = 0`. Stack: `[text]` → `[reader]`.
pub fn emit_string_reader_new(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    // `__content` comes straight off the stack — the construct owner spills it
    // for us, so the adapter no longer reserves a slot to hold it.
    class_slots::emit_class_construct(
        chunk,
        STRING_READER_TYPE,
        &[
            (field_slot(CONTENT_KEY), ValueSource::Stack),
            (field_slot(POS_KEY), ValueSource::ConstI32(0)),
            (field_slot(DISPOSED_KEY), ValueSource::ConstBool(false)),
        ],
        line,
    );
}

/// `reader.ReadLine()` — return next line up to (excluding) `\n`,
/// advancing `__pos` past it. Returns `null` if `__pos >= len`.
/// Stack: `[reader]` → `[line_or_null]`.
pub fn emit_stream_reader_read_line(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
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
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(CONTENT_KEY),
        Dest::Local(content_slot),
        line,
    );

    // pos = reader.__pos
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(POS_KEY),
        Dest::Stack,
        line,
    );
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
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_ADD, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(POS_KEY),
        ValueSource::Stack,
        line,
    );

    chunk.emit_end(line);
    chunk.patch_block(done_block);

    // Push result.
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `reader.ReadToEnd()` — return remaining content from `__pos` to end,
/// advancing `__pos` to end. Stack: `[reader]` → `[remaining]`.
pub fn emit_stream_reader_read_to_end(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let reader_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);
    emit_throw_if_disposed(chunk, reader_slot, line);

    // content = reader.__content
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(CONTENT_KEY),
        Dest::Stack,
        line,
    );

    // pos = reader.__pos (i32)
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(POS_KEY),
        Dest::Stack,
        line,
    );
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
    chunk.emit_op_u16(Op::LOCAL_GET, content_slot, line);
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(POS_KEY),
        ValueSource::Stack,
        line,
    );
}

/// `reader.AtEndOfStream` — `__pos >= wasm:js-string.length(__content)`.
/// Stack: `[reader]` → `[bool]`.
pub fn emit_stream_reader_at_end(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let reader_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);
    emit_throw_if_disposed(chunk, reader_slot, line);

    // pos
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(POS_KEY),
        Dest::Stack,
        line,
    );
    // len
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(CONTENT_KEY),
        Dest::Stack,
        line,
    );
    host::emit(chunk, "wasm:js-string", "length", 1, line);
    // pos < len → DYN_NOT → at-end
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}

/// `new StreamWriter(path)` / `new StreamWriter(stream)` — initialise the
/// destination + an empty `__buf`. Stack: `[destination, …]` → `[writer]`.
pub fn emit_stream_writer_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    emit_drop_extra_args(chunk, argc, line);
    let path_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, path_slot, line);

    class_slots::emit_class_construct(
        chunk,
        WRITER_TYPE,
        &[
            (field_slot(BUF_KEY), ValueSource::ConstStr(String::new())),
            (field_slot("NewLine"), ValueSource::ConstStr("\n".to_string())),
            (field_slot("newline"), ValueSource::ConstStr("\n".to_string())),
        ],
        line,
    );

    // A string destination is a file PATH; anything else is a byte STREAM the
    // flush appends to. Which one it is decides where `Flush` writes.
    // ⛔ The `if` arms are stack-NEUTRAL: the writer is re-read from a local
    // inside each arm rather than left on the stack across the branch.
    let writer_slot = reserve_slot(chunk);
    core_wasm::dup(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, writer_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path_slot, line);
    host::emit(chunk, "wasm:js-string", "test", 1, line);
    chunk.emit_if(line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(writer_slot),
        &field_slot(PATH_KEY),
        ValueSource::Local(path_slot),
        line,
    );
    chunk.emit_else(line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(writer_slot),
        &field_slot(STREAM_KEY),
        ValueSource::Local(path_slot),
        line,
    );
    chunk.emit_end(line);
}

/// `new StringWriter()` / `new StringWriter(StringBuilder)` — initialise an
/// in-memory buffer. Stack: `[]` or `[builder]` → `[writer]`.
pub fn emit_string_writer_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let initial_slot = reserve_slot(chunk);
    let builder_slot = reserve_slot(chunk);
    if argc > 0 {
        chunk.emit_op_u16(Op::LOCAL_SET, builder_slot, line);
        class_slots::emit_class_get(
            chunk,
            ObjSource::Local(builder_slot),
            &field_slot(SB_BUFFER_KEY),
            Dest::Local(initial_slot),
            line,
        );
    } else {
        push_const(chunk, Value::String(Arc::from("")), line);
        chunk.emit_op_u16(Op::LOCAL_SET, initial_slot, line);

        class_slots::emit_class_construct(
            chunk,
            "StringBuilder",
            &[(field_slot(SB_BUFFER_KEY), ValueSource::Local(initial_slot))],
            line,
        );
        chunk.emit_op_u16(Op::LOCAL_SET, builder_slot, line);
    }

    class_slots::emit_class_construct(
        chunk,
        STRING_WRITER_TYPE,
        &[
            (field_slot(BUF_KEY), ValueSource::Local(initial_slot)),
            (field_slot(BUILDER_KEY), ValueSource::Local(builder_slot)),
            (field_slot("NewLine"), ValueSource::ConstStr("\n".to_string())),
            (field_slot("newline"), ValueSource::ConstStr("\n".to_string())),
        ],
        line,
    );

    let encoding_slot = reserve_slot(chunk);
    class_slots::emit_class_construct(
        chunk,
        "Encoding",
        &[
            (field_slot("WebName"), ValueSource::ConstStr("utf-16".to_string())),
            (field_slot("webname"), ValueSource::ConstStr("utf-16".to_string())),
        ],
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, encoding_slot, line);
    for key in ["Encoding", "encoding"] {
        core_wasm::dup(chunk, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(key),
            ValueSource::Local(encoding_slot),
            line,
        );
    }

    let obj_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    let _ = chunk;

    bind_string_writer_to_string(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `writer.Write(s)` — append `s` to `__buf`. Stack: `[writer, s]` → `[null]`.
pub fn emit_stream_writer_write(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let s_slot = reserve_slot(chunk);
    let writer_slot = reserve_slot(chunk);
    let new_buf_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, writer_slot, line);
    let left_slot = reserve_slot(chunk);
    let right_slot = reserve_slot(chunk);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(writer_slot),
        &field_slot(BUF_KEY),
        Dest::Local(left_slot),
        line,
    );
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
pub fn emit_stream_writer_write_3(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
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
    let s_slot = reserve_slot(chunk);
    let writer_slot = reserve_slot(chunk);
    let new_buf_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, writer_slot, line);
    let left_slot = reserve_slot(chunk);
    let right_slot = reserve_slot(chunk);
    let nl_slot = reserve_slot(chunk);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(writer_slot),
        &field_slot(BUF_KEY),
        Dest::Local(left_slot),
        line,
    );
    super::console_adapter::emit_dotnet_stringify(chunk, left_slot, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, s_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    super::console_adapter::emit_dotnet_stringify(chunk, right_slot, right_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(writer_slot),
        &field_slot("newline"),
        Dest::Local(nl_slot),
        line,
    );
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
    for key in [BUILDER_KEY, SB_BUFFER_KEY] {
        class_slots::emit_class_get(chunk, ObjSource::Stack, &field_slot(key), Dest::Stack, line);
    }
}

pub fn emit_string_writer_get_string_builder(chunks: &mut [Chunk], current: usize, line: u32) {
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Stack,
        &field_slot(BUILDER_KEY),
        Dest::Stack,
        line,
    );
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

    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(CONTENT_KEY),
        Dest::Local(content_slot),
        line,
    );

    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(POS_KEY),
        Dest::Stack,
        line,
    );
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

    chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, read_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(POS_KEY),
        ValueSource::Stack,
        line,
    );

    chunk.emit_op_u16(Op::LOCAL_GET, read_slot, line);
}

fn emit_string_reader_read_char(chunks: &mut [Chunk], current: usize, line: u32, advance: bool) {
    let chunk = &mut chunks[current];
    let reader_slot = reserve_slot(chunk);
    let content_slot = reserve_slot(chunk);
    let pos_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, reader_slot, line);
    emit_throw_if_disposed(chunk, reader_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(CONTENT_KEY),
        Dest::Local(content_slot),
        line,
    );
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(reader_slot),
        &field_slot(POS_KEY),
        Dest::Stack,
        line,
    );
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
        chunk.emit_op_u16(Op::LOCAL_GET, pos_slot, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Local(reader_slot),
            &field_slot(POS_KEY),
            ValueSource::Stack,
            line,
        );
    }
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `[]` → `[]` — persist the writer's buffered text to wherever it belongs:
/// appended to the underlying byte stream when it was built over one, written
/// to `__path` otherwise. The buffer is CLEARED for the stream case, so two
/// flushes do not write the text twice.
fn emit_writer_persist(chunks: &mut [Chunk], current: usize, writer_slot: u16, line: u32) {
    let chunk = &mut chunks[current];
    let obj = ObjSource::Local(writer_slot);

    class_slots::emit_class_get(chunk, obj, &field_slot(STREAM_KEY), Dest::Stack, line);
    host::emit(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);

    // Stream-backed: the buffered text, UTF-8 encoded, through the stream's
    // own write — which is what grows it and moves its cursor.
    class_slots::emit_class_get(chunk, obj, &field_slot(STREAM_KEY), Dest::Stack, line);
    host::emit(chunk, "web:encoding", "encoderNew", 0, line);
    class_slots::emit_class_get(chunk, obj, &field_slot(BUF_KEY), Dest::Stack, line);
    host::emit(chunk, "web:encoding", "encode", 2, line);
    super::memory_stream_adapter::emit_write(chunks, current, 1, line);
    let chunk = &mut chunks[current];
    class_slots::emit_class_set(
        chunk,
        obj,
        &field_slot(BUF_KEY),
        ValueSource::ConstStr(String::new()),
        line,
    );

    chunk.emit_else(line);

    class_slots::emit_class_get(chunk, obj, &field_slot(PATH_KEY), Dest::Stack, line);
    class_slots::emit_class_get(chunk, obj, &field_slot(BUF_KEY), Dest::Stack, line);
    vybe_compiler::primitives::fs_path::emit_write_file(chunk, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
}

/// `writer.Flush()` / `writer.Close()` — the buffered text to the file or the
/// underlying stream. Stack: `[writer]` → `[null]`.
pub fn emit_stream_writer_flush(chunks: &mut [Chunk], current: usize, line: u32) {
    let writer_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, writer_slot, line);
    emit_writer_persist(chunks, current, writer_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `reader.Close()` / `writer.Close()` — no-op for readers, flush for writers.
/// Stack: `[stream]` → `[null]`.
pub fn emit_stream_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let stream_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, stream_slot, line);

    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(stream_slot),
        &field_slot(TYPE_KEY),
        Dest::Stack,
        line,
    );
    push_const(chunk, Value::String(Arc::from(WRITER_TYPE)), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);

    let skip_flush = chunk.emit_block(line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(0, line);

    emit_writer_persist(chunks, current, stream_slot, line);

    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.patch_block(skip_flush);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `reader.Close()` — no-op (no resource to release in load-whole-file model).
/// Stack: `[reader]` → `[null]`.
pub fn emit_stream_reader_close(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(DISPOSED_KEY),
        ValueSource::ConstBool(true),
        line,
    );
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
