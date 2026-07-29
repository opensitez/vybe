use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::instructions::host;

const TYPE_KEY: &str = "__type";
const ENCODING_KEY: &str = "__encoding";

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

fn emit_set_string_field(chunk: &mut Chunk, key: &str, value: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(key)));
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from(value)), line);
    chunk.emit_op_u16(Op::STRUCT_SET, key_idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn encoding_web_name(encoding: &str) -> &str {
    match encoding.to_ascii_lowercase().as_str() {
        "utf8" | "utf-8" => "utf-8",
        "utf16le" | "unicode" => "utf-16",
        "utf16be" => "utf-16BE",
        "utf32" | "utf-32" => "utf-32",
        "ascii" | "us-ascii" => "us-ascii",
        "latin1" | "iso-8859-1" => "iso-8859-1",
        _ => encoding,
    }
}

pub fn emit_encoding_value(chunks: &mut [Chunk], current: usize, encoding: &str, line: u32) {
    let chunk = &mut chunks[current];
    let web_name = encoding_web_name(encoding);
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    emit_set_string_field(chunk, TYPE_KEY, "Encoding", line);
    emit_set_string_field(chunk, ENCODING_KEY, encoding, line);
    emit_set_string_field(chunk, "WebName", web_name, line);
    emit_set_string_field(chunk, "webname", web_name, line);
    emit_set_string_field(chunk, "HeaderName", web_name, line);
    emit_set_string_field(chunk, "headername", web_name, line);
    let read_only_idx = chunk.add_constant(Value::String(Arc::from("IsReadOnly")));
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::STRUCT_SET, read_only_idx, line);
    chunk.emit_op(Op::DROP, line);
    let read_only_lower_idx = chunk.add_constant(Value::String(Arc::from("isreadonly")));
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::STRUCT_SET, read_only_lower_idx, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_encoding_name_from_receiver(chunk: &mut Chunk, recv_slot: u16, fallback: &str, line: u32) {
    let key_idx = chunk.add_constant(Value::String(Arc::from(ENCODING_KEY)));
    chunk.emit_op_u16(Op::LOCAL_GET, recv_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, key_idx, line);
    let enc_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, enc_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from(fallback)), line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    chunk.emit_end(line);
}

fn stash_receiver_text(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    fallback: &str,
    line: u32,
) -> (u16, u16) {
    let chunk = &mut chunks[current];
    let text_slot = reserve_slot(chunk);
    let enc_slot = reserve_slot(chunk);

    if argc > 1 {
        let recv_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, recv_slot, line);
        emit_encoding_name_from_receiver(chunk, recv_slot, fallback, line);
        chunk.emit_op_u16(Op::LOCAL_SET, enc_slot, line);
    } else {
        chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
        push_const(chunk, Value::String(Arc::from(fallback)), line);
        chunk.emit_op_u16(Op::LOCAL_SET, enc_slot, line);
    }

    (enc_slot, text_slot)
}

fn stash_receiver_bytes(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    fallback: &str,
    line: u32,
) -> (u16, u16) {
    let chunk = &mut chunks[current];
    let bytes_slot = reserve_slot(chunk);
    let enc_slot = reserve_slot(chunk);

    if argc > 1 {
        let recv_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, recv_slot, line);
        emit_encoding_name_from_receiver(chunk, recv_slot, fallback, line);
        chunk.emit_op_u16(Op::LOCAL_SET, enc_slot, line);
    } else {
        chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
        push_const(chunk, Value::String(Arc::from(fallback)), line);
        chunk.emit_op_u16(Op::LOCAL_SET, enc_slot, line);
    }

    (enc_slot, bytes_slot)
}

fn emit_char_array_to_string(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let char_code_idx = chunks[current].add_import("wasm:js-string", "charCodeAt");
    let from_chars_idx = chunks[current].add_import("wasm:js-string", "fromCharCodeArray");
    let chunk = &mut chunks[current];
    let units_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let elem_slot = reserve_slot(chunk);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, units_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    host::emit(&mut chunks[current], "ecma:string", "String", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, char_code_idx, line);
    chunks[current].emit(2, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, from_chars_idx, line);
    chunks[current].emit(3, line);
}

fn emit_text_value(chunks: &mut [Chunk], current: usize, text_slot: u16, line: u32) {
    let string_test_idx = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, string_test_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_else(line);
    emit_char_array_to_string(chunks, current, text_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_encoding_get_bytes(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    fallback: &str,
    line: u32,
) {
    let (enc_slot, text_slot) = stash_receiver_text(chunks, current, argc, fallback, line);
    let from_idx = chunks[current].add_import("node:buffer", "from");
    let byte_len_idx = chunks[current].add_import("node:buffer", "byteLength");
    let str_len_idx = chunks[current].add_import("wasm:js-string", "length");
    let value_slot = reserve_slot(&mut chunks[current]);
    emit_text_value(chunks, current, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("utf32")),
        line,
    );
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_utf32_text_to_bytes(chunks, current, value_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("us-ascii:throw")),
        line,
    );
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("utf-8")),
        line,
    );
    chunks[current].emit_op_u16(Op::CALL_IMPORT, byte_len_idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, str_len_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Unable to encode character.", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "EncoderFallbackException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("us-ascii")),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, enc_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, from_idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_end(line);
}

pub fn emit_encoding_get_byte_count(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    fallback: &str,
    line: u32,
) {
    let (enc_slot, text_slot) = stash_receiver_text(chunks, current, argc, fallback, line);
    let byte_len_idx = chunks[current].add_import("node:buffer", "byteLength");
    emit_text_value(chunks, current, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, byte_len_idx, line);
    chunks[current].emit(2, line);
}

pub fn emit_encoding_get_string(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    fallback: &str,
    line: u32,
) {
    let (enc_slot, bytes_slot) = stash_receiver_bytes(chunks, current, argc, fallback, line);
    let to_string_idx = chunks[current].add_import("node:buffer", "toString");

    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("utf16le")),
        line,
    );
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_utf16_bytes_to_string(chunks, current, bytes_slot, false, 2, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("utf32")),
        line,
    );
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_utf16_bytes_to_string(chunks, current, bytes_slot, false, 4, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("utf-8:throw")),
        line,
    );
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_throw_on_invalid_utf8_bytes(chunks, current, bytes_slot, line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("utf-8")),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, enc_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, enc_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, to_string_idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_throw_on_invalid_utf8_bytes(
    chunks: &mut [Chunk],
    current: usize,
    bytes_slot: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let len_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    core_wasm::i32_const(&mut chunks[current], line, 247);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("Unable to decode bytes.", line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "DecoderFallbackException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
}

fn emit_utf16_bytes_to_string(
    chunks: &mut [Chunk],
    current: usize,
    bytes_slot: u16,
    big_endian: bool,
    stride: i32,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let units_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, units_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    let (first_offset, second_offset) = if big_endian { (1, 0) } else { (0, 1) };
    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    if first_offset != 0 {
        core_wasm::i32_const(&mut chunks[current], line, first_offset);
        chunks[current].emit_op(Op::I32_ADD, line);
    }
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    if second_offset != 0 {
        core_wasm::i32_const(&mut chunks[current], line, second_offset);
        chunks[current].emit_op(Op::I32_ADD, line);
    }
    chunks[current].emit_op(Op::ARRAY_GET, line);
    core_wasm::i32_const(&mut chunks[current], line, 256);
    chunks[current].emit_op(Op::I32_MUL, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, stride);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    let from_chars_idx = chunks[current].add_import("wasm:js-string", "fromCharCodeArray");
    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, units_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, from_chars_idx, line);
    chunks[current].emit(3, line);
}

pub fn emit_encoding_unicode_get_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (_enc_slot, bytes_slot) = stash_receiver_bytes(chunks, current, argc, "utf16le", line);
    emit_utf16_bytes_to_string(chunks, current, bytes_slot, false, 2, line);
}

pub fn emit_encoding_utf32_get_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (_enc_slot, bytes_slot) = stash_receiver_bytes(chunks, current, argc, "utf32", line);
    emit_utf16_bytes_to_string(chunks, current, bytes_slot, false, 4, line);
}

fn emit_utf16be_get_bytes(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (_enc_slot, text_slot) = stash_receiver_text(chunks, current, argc, "utf16be", line);
    let len_idx = chunks[current].add_import("wasm:js-string", "length");
    let char_code_idx = chunks[current].add_import("wasm:js-string", "charCodeAt");
    let chunk = &mut chunks[current];
    let bytes_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let code_slot = reserve_slot(chunk);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, len_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, char_code_idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 8);
    chunks[current].emit_op(Op::I32_SHR_U, line);
    core_wasm::i32_const(&mut chunks[current], line, 255);
    chunks[current].emit_op(Op::I32_AND, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, code_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 255);
    chunks[current].emit_op(Op::I32_AND, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
}

pub fn emit_encoding_big_endian_unicode_get_bytes(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    emit_utf16be_get_bytes(chunks, current, argc, line);
}

fn emit_utf32_text_to_bytes(chunks: &mut [Chunk], current: usize, text_slot: u16, line: u32) {
    let len_idx = chunks[current].add_import("wasm:js-string", "length");
    let char_code_idx = chunks[current].add_import("wasm:js-string", "charCodeAt");
    let chunk = &mut chunks[current];
    let bytes_slot = reserve_slot(chunk);
    let len_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);
    let code_slot = reserve_slot(chunk);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, len_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, char_code_idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_slot, line);
    for shift in [0, 8, 16, 24] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, code_slot, line);
        if shift != 0 {
            core_wasm::i32_const(&mut chunks[current], line, shift);
            chunks[current].emit_op(Op::I32_SHR_U, line);
        }
        core_wasm::i32_const(&mut chunks[current], line, 255);
        chunks[current].emit_op(Op::I32_AND, line);
        vybe_compiler::primitives::collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
}

pub fn emit_encoding_utf32_get_bytes(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (_enc_slot, text_slot) = stash_receiver_text(chunks, current, argc, "utf32", line);
    emit_utf32_text_to_bytes(chunks, current, text_slot, line);
}

pub fn emit_encoding_get_preamble(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
}

pub fn emit_encoding_utf8_get_preamble(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    core_wasm::i32_const(&mut chunks[current], line, 239);
    core_wasm::i32_const(&mut chunks[current], line, 187);
    core_wasm::i32_const(&mut chunks[current], line, 191);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 3, line);
}

pub fn emit_encoding_get_max_byte_count(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    factor: i32,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let char_count_slot = reserve_slot(chunk);
    let factor_slot = reserve_slot(chunk);
    if argc > 1 {
        let recv_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, char_count_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, recv_slot, line);
        emit_encoding_name_from_receiver(chunk, recv_slot, "utf-8", line);
        push_const(chunk, Value::String(Arc::from("utf16le")), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if_value(line);
        push_const(chunk, Value::I32(2), line);
        chunk.emit_else(line);
        push_const(chunk, Value::I32(factor), line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_SET, factor_slot, line);
    } else {
        chunk.emit_op_u16(Op::LOCAL_SET, char_count_slot, line);
        push_const(chunk, Value::I32(factor), line);
        chunk.emit_op_u16(Op::LOCAL_SET, factor_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, char_count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, factor_slot, line);
    chunk.emit_op(Op::I32_MUL, line);
}

pub fn emit_encoding_get_max_char_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_encoding_get_max_byte_count(chunks, current, argc, 4, line);
}

pub fn emit_encoding_utf32_get_byte_count(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let (_enc_slot, text_slot) = stash_receiver_text(chunks, current, argc, "utf32", line);
    let len_idx = chunks[current].add_import("wasm:js-string", "length");
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, len_idx, line);
    chunks[current].emit(1, line);
    core_wasm::i32_const(&mut chunks[current], line, 4);
    chunks[current].emit_op(Op::I32_MUL, line);
}

pub fn emit_encoding_get_char_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc >= 3 {
        let count_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_op(Op::DROP, line);
        if argc > 3 {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    } else {
        let bytes_slot = reserve_slot(chunk);
        if argc > 1 {
            chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
            chunk.emit_op(Op::DROP, line);
        } else {
            chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
        }
        chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
        chunk.emit_op(Op::ARRAY_LENGTH, line);
    }
}

pub fn emit_encoding_get_chars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let char_index_slot = reserve_slot(chunk);
    let chars_slot = reserve_slot(chunk);
    let count_slot = reserve_slot(chunk);
    let byte_index_slot = reserve_slot(chunk);
    let bytes_slot = reserve_slot(chunk);
    let i_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, char_index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, chars_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, byte_index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    if argc > 5 {
        chunk.emit_op(Op::DROP, line);
    }
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, char_index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    let one_slot = reserve_slot(&mut chunks[current]);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, byte_index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    let from_chars_idx = chunks[current].add_import("wasm:js-string", "fromCharCodeArray");
    chunks[current].emit_op_u16(Op::LOCAL_GET, one_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, from_chars_idx, line);
    chunks[current].emit(3, line);
    chunks[current].emit_op(Op::ARRAY_SET, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
}

pub fn emit_encoding_convert(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let bytes_slot = reserve_slot(chunk);
    let dst_slot = reserve_slot(chunk);
    let src_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dst_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, src_slot, line);

    let src_enc_slot = reserve_slot(chunk);
    let dst_enc_slot = reserve_slot(chunk);
    emit_encoding_name_from_receiver(chunk, src_slot, "utf8", line);
    chunk.emit_op_u16(Op::LOCAL_SET, src_enc_slot, line);
    emit_encoding_name_from_receiver(chunk, dst_slot, "utf8", line);
    chunk.emit_op_u16(Op::LOCAL_SET, dst_enc_slot, line);

    let from_idx = chunks[current].add_import("node:buffer", "from");
    let to_string_idx = chunks[current].add_import("node:buffer", "toString");
    let text_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src_enc_slot, line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("utf16le")),
        line,
    );
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_utf16_bytes_to_string(chunks, current, bytes_slot, false, 2, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src_enc_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, to_string_idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dst_enc_slot, line);
    push_const(
        &mut chunks[current],
        Value::String(Arc::from("utf32")),
        line,
    );
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    emit_encoding_utf32_get_bytes(chunks, current, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dst_enc_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, from_idx, line);
    chunks[current].emit(2, line);
    chunks[current].emit_end(line);
}

pub fn emit_encoding_get_encoding(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let name_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    host::emit(chunk, "ecma:string", "String", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    push_const(chunk, Value::String(Arc::from("65001")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from("utf-8")), line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
    push_const(chunk, Value::String(Arc::from("utf8")), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);
    push_const(chunk, Value::String(Arc::from("utf-8")), line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_end(line);
    if argc == 3 {
        chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
        push_const(chunk, Value::String(Arc::from("us-ascii")), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::String(Arc::from("us-ascii:throw")), line);
        chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
        push_const(chunk, Value::String(Arc::from("utf-8")), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
        push_const(chunk, Value::String(Arc::from("utf-8:throw")), line);
        chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    let type_key = chunk.add_constant(Value::String(Arc::from(TYPE_KEY)));
    let enc_key = chunk.add_constant(Value::String(Arc::from(ENCODING_KEY)));
    let web_key = chunk.add_constant(Value::String(Arc::from("WebName")));
    let web_lower_key = chunk.add_constant(Value::String(Arc::from("webname")));
    let header_key = chunk.add_constant(Value::String(Arc::from("HeaderName")));
    let header_lower_key = chunk.add_constant(Value::String(Arc::from("headername")));
    let readonly_key = chunk.add_constant(Value::String(Arc::from("IsReadOnly")));
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Encoding")), line);
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);
    for key in [
        enc_key,
        web_key,
        web_lower_key,
        header_key,
        header_lower_key,
    ] {
        vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
        chunk.emit_op_u16(Op::LOCAL_GET, name_slot, line);
        chunk.emit_op_u16(Op::STRUCT_SET, key, line);
        chunk.emit_op(Op::DROP, line);
    }
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::STRUCT_SET, readonly_key, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_encoding_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let other_slot = reserve_slot(chunk);
    let recv_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, recv_slot, line);
    emit_encoding_name_from_receiver(chunk, recv_slot, "utf8", line);
    emit_encoding_name_from_receiver(chunk, other_slot, "utf8", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
}

pub fn emit_object_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let right_slot = reserve_slot(chunk);
    let left_slot = reserve_slot(chunk);
    let left_enc_slot = reserve_slot(chunk);
    let right_enc_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_slot, line);

    let enc_key = chunk.add_constant(Value::String(Arc::from(ENCODING_KEY)));
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, enc_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, left_enc_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_op_u16(Op::STRUCT_GET, enc_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, right_enc_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, left_enc_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::object::emit_equals(chunk, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_enc_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::object::emit_equals(chunk, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, left_enc_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_enc_slot, line);
    vybe_compiler::primitives::object::emit_equals(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
