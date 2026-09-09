use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};

use vybe_compiler::primitives::functions::create_function_chunk;
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::object::emit_bind_method_with_slot;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

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
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::ConstStr(value.to_string()),
        line,
    );
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

fn encoding_code_page(encoding: &str) -> i32 {
    match encoding.to_ascii_lowercase().as_str() {
        "utf8" | "utf-8" | "utf-8:throw" => 65001,
        "utf16le" | "unicode" => 1200,
        "utf16be" => 1201,
        "utf32" | "utf-32" => 12000,
        "ascii" | "us-ascii" | "us-ascii:throw" => 20127,
        "latin1" | "iso-8859-1" => 28591,
        _ => 65001,
    }
}

pub fn emit_encoding_value(chunks: &mut [Chunk], current: usize, encoding: &str, line: u32) {
    let chunk = &mut chunks[current];
    let web_name = encoding_web_name(encoding);
    let code_page = encoding_code_page(encoding);
    class_slots::emit_class_construct(
        chunk,
        "Encoding",
        &[
            (
                field_slot(ENCODING_KEY),
                ValueSource::ConstStr(encoding.to_string()),
            ),
            (
                field_slot("WebName"),
                ValueSource::ConstStr(web_name.to_string()),
            ),
            (
                field_slot("webname"),
                ValueSource::ConstStr(web_name.to_string()),
            ),
            (
                field_slot("HeaderName"),
                ValueSource::ConstStr(web_name.to_string()),
            ),
            (
                field_slot("headername"),
                ValueSource::ConstStr(web_name.to_string()),
            ),
            (field_slot("CodePage"), ValueSource::ConstI32(code_page)),
            (field_slot("codepage"), ValueSource::ConstI32(code_page)),
            (
                field_slot("EncoderFallback"),
                ValueSource::ConstStr("EncoderFallback".to_string()),
            ),
            (
                field_slot("encoderfallback"),
                ValueSource::ConstStr("EncoderFallback".to_string()),
            ),
            (
                field_slot("DecoderFallback"),
                ValueSource::ConstStr("DecoderFallback".to_string()),
            ),
            (
                field_slot("decoderfallback"),
                ValueSource::ConstStr("DecoderFallback".to_string()),
            ),
            (field_slot("EmitBom"), ValueSource::ConstBool(false)),
            (field_slot("emitbom"), ValueSource::ConstBool(false)),
            (field_slot("IsReadOnly"), ValueSource::ConstBool(false)),
            (field_slot("isreadonly"), ValueSource::ConstBool(false)),
        ],
        line,
    );
}

pub fn emit_utf8encoding_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let bom_slot = reserve_slot(&mut chunks[current]);
    match argc {
        0 => chunks[current].emit_bool_const(false, line),
        1 => {}
        _ => chunks[current].emit_op(Op::DROP, line),
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, bom_slot, line);
    for _ in 2..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_encoding_value(chunks, current, "utf-8", line);
    for key in ["EmitBom", "emitbom"] {
        core_wasm::dup(&mut chunks[current], line);
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Stack,
            &field_slot(key),
            ValueSource::Local(bom_slot),
            line,
        );
    }
}

fn emit_encoding_name_from_receiver(chunk: &mut Chunk, recv_slot: u16, fallback: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(recv_slot),
        &field_slot(ENCODING_KEY),
        Dest::Stack,
        line,
    );
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
    chunks[current].emit_call(char_code_idx, 2, line);
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
    chunks[current].emit_call(from_chars_idx, 3, line);
}

fn emit_text_value(chunks: &mut [Chunk], current: usize, text_slot: u16, line: u32) {
    let string_test_idx = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_call(string_test_idx, 1, line);
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
    if argc >= 6 {
        let chunk = &mut chunks[current];
        let byte_index_slot = reserve_slot(chunk);
        let dest_slot = reserve_slot(chunk);
        let char_count_slot = reserve_slot(chunk);
        let char_index_slot = reserve_slot(chunk);
        let text_slot = reserve_slot(chunk);
        let recv_slot = reserve_slot(chunk);
        let enc_slot = reserve_slot(chunk);
        let value_slot = reserve_slot(chunk);
        let bytes_slot = reserve_slot(chunk);
        let written_slot = reserve_slot(chunk);
        let i_slot = reserve_slot(chunk);

        chunk.emit_op_u16(Op::LOCAL_SET, byte_index_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, dest_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, char_count_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, char_index_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, recv_slot, line);
        for _ in 6..argc {
            chunk.emit_op(Op::DROP, line);
        }

        emit_encoding_name_from_receiver(chunk, recv_slot, fallback, line);
        chunk.emit_op_u16(Op::LOCAL_SET, enc_slot, line);
        emit_text_value(chunks, current, text_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, char_index_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, char_count_slot, line);
        host::emit(&mut chunks[current], "ecma:string", "substr", 3, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

        let from_idx = chunks[current].add_import("node:buffer", "from");
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, enc_slot, line);
        chunks[current].emit_call(from_idx, 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, written_slot, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

        let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, written_slot, line);
        chunks[current].emit_op(Op::I32_LT_S, line);
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, dest_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, byte_index_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[current].emit_op(Op::ARRAY_GET, line);
        vybe_compiler::primitives::collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, written_slot, line);
        return;
    }

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
    chunks[current].emit_call(byte_len_idx, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_call(str_len_idx, 1, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    crate::emitter::core::exceptions::emit_new_typed(
        chunks,
        current,
        "EncoderFallbackException",
        class_slots::ValueSource::ConstStr("Unable to encode character.".to_string()),
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
    chunks[current].emit_call(from_idx, 2, line);
    chunks[current].emit_end(line);
}

fn emit_bytes_slot_to_string(
    chunks: &mut [Chunk],
    current: usize,
    enc_slot: u16,
    bytes_slot: u16,
    line: u32,
) {
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
    chunks[current].emit_call(to_string_idx, 2, line);
    chunks[current].emit_end(line);
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
    chunks[current].emit_call(byte_len_idx, 2, line);
}

pub fn emit_encoding_get_string(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    fallback: &str,
    line: u32,
) {
    if argc >= 4 {
        let chunk = &mut chunks[current];
        let count_slot = reserve_slot(chunk);
        let index_slot = reserve_slot(chunk);
        let bytes_slot = reserve_slot(chunk);
        let recv_slot = reserve_slot(chunk);
        let enc_slot = reserve_slot(chunk);
        let end_slot = reserve_slot(chunk);
        let slice_slot = reserve_slot(chunk);

        chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, recv_slot, line);
        for _ in 4..argc {
            chunk.emit_op(Op::DROP, line);
        }
        emit_encoding_name_from_receiver(chunk, recv_slot, fallback, line);
        chunk.emit_op_u16(Op::LOCAL_SET, enc_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
        vybe_compiler::primitives::collections::emit_slice(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slice_slot, line);
        emit_bytes_slot_to_string(chunks, current, enc_slot, slice_slot, line);
        return;
    }

    let (enc_slot, bytes_slot) = stash_receiver_bytes(chunks, current, argc, fallback, line);
    emit_bytes_slot_to_string(chunks, current, enc_slot, bytes_slot, line);
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
    crate::emitter::core::exceptions::emit_new_typed(
        chunks,
        current,
        "DecoderFallbackException",
        class_slots::ValueSource::ConstStr("Unable to decode bytes.".to_string()),
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
    chunks[current].emit_call(from_chars_idx, 3, line);
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
    chunks[current].emit_call(len_idx, 1, line);
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
    chunks[current].emit_call(char_code_idx, 2, line);
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
    chunks[current].emit_call(len_idx, 1, line);
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
    chunks[current].emit_call(char_code_idx, 2, line);
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
    if argc > 0 {
        let recv_slot = reserve_slot(&mut chunks[current]);
        chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        class_slots::emit_class_get(
            &mut chunks[current],
            ObjSource::Local(recv_slot),
            &field_slot("EmitBom"),
            Dest::Stack,
            line,
        );
        chunks[current].emit_if_value(line);
        core_wasm::i32_const(&mut chunks[current], line, 239);
        core_wasm::i32_const(&mut chunks[current], line, 187);
        core_wasm::i32_const(&mut chunks[current], line, 191);
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 3, line);
        chunks[current].emit_else(line);
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_end(line);
        return;
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
    chunks[current].emit_call(len_idx, 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 4);
    chunks[current].emit_op(Op::I32_MUL, line);
}

pub fn emit_encoding_get_char_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 4 {
        let chunk = &mut chunks[current];
        let count_slot = reserve_slot(chunk);
        let index_slot = reserve_slot(chunk);
        let bytes_slot = reserve_slot(chunk);
        let recv_slot = reserve_slot(chunk);
        let enc_slot = reserve_slot(chunk);
        let end_slot = reserve_slot(chunk);
        let slice_slot = reserve_slot(chunk);
        chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
        chunk.emit_op_u16(Op::LOCAL_SET, recv_slot, line);
        for _ in 4..argc {
            chunk.emit_op(Op::DROP, line);
        }
        emit_encoding_name_from_receiver(chunk, recv_slot, "utf-8", line);
        chunk.emit_op_u16(Op::LOCAL_SET, enc_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
        vybe_compiler::primitives::collections::emit_slice(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slice_slot, line);
        emit_bytes_slot_to_string(chunks, current, enc_slot, slice_slot, line);
    } else {
        let (enc_slot, bytes_slot) = stash_receiver_bytes(chunks, current, argc, "utf-8", line);
        emit_bytes_slot_to_string(chunks, current, enc_slot, bytes_slot, line);
    }
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    host::emit(&mut chunks[current], "wasm:js-number", "toI32", 1, line);
}

pub fn emit_encoding_get_chars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 5 {
        let (enc_slot, bytes_slot) = stash_receiver_bytes(chunks, current, argc, "utf-8", line);
        emit_bytes_slot_to_string(chunks, current, enc_slot, bytes_slot, line);
        chunks[current].emit_string_const("", line);
        vybe_compiler::primitives::strings::emit_split(&mut chunks[current], line);
        return;
    }

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
    chunks[current].emit_call(from_chars_idx, 3, line);
    chunks[current].emit_op(Op::ARRAY_SET, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
}

fn bind_decoder_get_chars(chunks: &mut Vec<Chunk>, current: usize, this_slot: u16, line: u32) {
    let mut method = create_function_chunk("__decoder_getchars", 6);
    for local in 0..6 {
        method.emit_op_u16(Op::LOCAL_GET, local, line);
    }
    let mut method_chunks = vec![method];
    emit_decoder_get_chars(&mut method_chunks, 0, 6, line);
    method_chunks[0].emit_op(Op::RETURN, line);
    let method = method_chunks.pop().unwrap();
    chunks.push(method);
    let method_idx = chunks.len() - 1;
    for name in ["getchars", "GetChars"] {
        emit_bind_method_with_slot(
            &mut chunks[current],
            this_slot,
            name,
            None,
            method_idx,
            None,
            line,
        );
    }
}

pub fn emit_encoding_get_decoder(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let recv_slot = reserve_slot(&mut chunks[current]);
    let enc_slot = reserve_slot(&mut chunks[current]);
    let obj_slot = reserve_slot(&mut chunks[current]);
    if argc > 0 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, recv_slot, line);
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        emit_encoding_name_from_receiver(&mut chunks[current], recv_slot, "utf-8", line);
    } else {
        chunks[current].emit_string_const("utf-8", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, enc_slot, line);
    class_slots::emit_class_construct(
        &mut chunks[current],
        "Decoder",
        &[
            (field_slot(ENCODING_KEY), ValueSource::Local(enc_slot)),
            (field_slot("PendingByte"), ValueSource::ConstI32(0)),
            (field_slot("pendingbyte"), ValueSource::ConstI32(0)),
        ],
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_slot, line);
    bind_decoder_get_chars(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

pub fn emit_decoder_get_chars(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 6 {
        emit_encoding_get_chars(chunks, current, argc, line);
        return;
    }

    let from_chars_idx = chunks[current].add_import("wasm:js-string", "fromCharCodeArray");
    let chunk = &mut chunks[current];
    let char_index_slot = reserve_slot(chunk);
    let chars_slot = reserve_slot(chunk);
    let count_slot = reserve_slot(chunk);
    let byte_index_slot = reserve_slot(chunk);
    let bytes_slot = reserve_slot(chunk);
    let recv_slot = reserve_slot(chunk);
    let byte_slot = reserve_slot(chunk);
    let pending_slot = reserve_slot(chunk);

    chunk.emit_op_u16(Op::LOCAL_SET, char_index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, chars_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, byte_index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, recv_slot, line);
    for _ in 6..argc {
        chunk.emit_op(Op::DROP, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, byte_index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, byte_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(recv_slot),
        &field_slot("PendingByte"),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, pending_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, byte_slot, line);
    chunk.emit_i32_const(195, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(recv_slot),
        &field_slot("PendingByte"),
        ValueSource::Local(byte_slot),
        line,
    );
    chunk.emit_i32_const(0, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, chars_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, char_index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pending_slot, line);
    chunk.emit_i32_const(195, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_op_u16(Op::LOCAL_GET, byte_slot, line);
    chunk.emit_i32_const(169, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("\u{00E9}", line);
    chunk.emit_else(line);
    let one_slot = reserve_slot(&mut chunks[current]);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, byte_slot, line);
    vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, one_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_call(from_chars_idx, 3, line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(0, line);
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Local(recv_slot),
        &field_slot("PendingByte"),
        ValueSource::Stack,
        line,
    );
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_end(line);
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
    chunks[current].emit_call(to_string_idx, 2, line);
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
    chunks[current].emit_call(from_idx, 2, line);
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

    class_slots::emit_class_alloc(chunk, line);
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    push_const(chunk, Value::String(Arc::from("Encoding")), line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(TYPE_KEY),
        ValueSource::Stack,
        line,
    );
    for key in [
        ENCODING_KEY,
        "WebName",
        "webname",
        "HeaderName",
        "headername",
    ] {
        vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
        class_slots::emit_class_set(
            chunk,
            ObjSource::Stack,
            &field_slot(key),
            ValueSource::Local(name_slot),
            line,
        );
    }
    vybe_compiler::primitives::instructions::core_wasm::dup(chunk, line);
    chunk.emit_bool_const(false, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot("IsReadOnly"),
        ValueSource::Stack,
        line,
    );
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

    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(ENCODING_KEY),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, left_enc_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(ENCODING_KEY),
        Dest::Stack,
        line,
    );
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
