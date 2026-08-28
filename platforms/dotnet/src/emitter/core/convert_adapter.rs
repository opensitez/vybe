use vybe_compiler::primitives::instructions::host;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_compiler::primitives::class_slots;

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn emit_throw_dotnet_exception(chunk: &mut Chunk, exception_name: &str, message: &str, line: u32) {
    vybe_compiler::primitives::errors::emit_exception_new(
        chunk,
        exception_name,
        class_slots::ValueSource::ConstStr(message.to_string()),
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

fn emit_throw_if_null(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "ArgumentNullException",
        "Value cannot be null.",
        line,
    );
    chunk.emit_end(line);
}

fn emit_throw_if_slice_out_of_range(
    chunk: &mut Chunk,
    bytes_slot: u16,
    offset_slot: u16,
    length_slot: u16,
    line: u32,
) {
    let array_len_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, array_len_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, offset_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_op_u16(Op::LOCAL_GET, length_slot, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_LT_S, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op_u16(Op::LOCAL_GET, offset_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, length_slot, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_GET, array_len_slot, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "ArgumentOutOfRangeException",
        "Offset and length were out of bounds for the array.",
        line,
    );
    chunk.emit_end(line);
}

fn emit_buffer_to_base64(
    chunks: &mut [Chunk],
    current: usize,
    bytes_slot: u16,
    start_slot: Option<u16>,
    end_slot: Option<u16>,
    line: u32,
) {
    let binary_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunks[current].emit_else(line);
    vybe_compiler::primitives::base64::emit_byte_array_slot_to_binary_string(
        chunks,
        current,
        Some(bytes_slot),
        start_slot,
        end_slot,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, binary_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, binary_slot, line);
    vybe_compiler::primitives::base64::emit_encode_binary_string(chunks, current, line);
}

fn emit_base64_char_valid(chunk: &mut Chunk, code_slot: u16, line: u32) {
    let ok_slot = reserve_slot(chunk);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ok_slot, line);
    for (lo, hi) in [(65, 90), (97, 122), (48, 57)] {
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        chunk.emit_i32_const(lo, line);
        chunk.emit_op(Op::I32_GE_S, line);
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        chunk.emit_i32_const(hi, line);
        chunk.emit_op(Op::I32_LE_S, line);
        chunk.emit_op(Op::I32_AND, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ok_slot, line);
        chunk.emit_end(line);
    }
    for code in [43, 47, 61] {
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        chunk.emit_i32_const(code, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_if(line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, ok_slot, line);
        chunk.emit_end(line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, ok_slot, line);
}

fn emit_base64_whitespace(chunk: &mut Chunk, code_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunk.emit_i32_const(9, line);
    chunk.emit_op(Op::I32_EQ, line);
    for code in [10, 13, 32] {
        chunk.emit_op_u16(Op::LOCAL_GET, code_slot, line);
        chunk.emit_i32_const(code, line);
        chunk.emit_op(Op::I32_EQ, line);
        chunk.emit_op(Op::I32_OR, line);
    }
}

fn emit_filter_dotnet_base64(
    chunks: &mut [Chunk],
    current: usize,
    text_slot: u16,
    throw_on_invalid: bool,
    line: u32,
) -> u16 {
    let out_slot = reserve_slot(&mut chunks[current]);
    let i_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    let code_slot = reserve_slot(&mut chunks[current]);
    let invalid_slot = reserve_slot(&mut chunks[current]);
    let pad_seen_slot = reserve_slot(&mut chunks[current]);
    let pad_count_slot = reserve_slot(&mut chunks[current]);
    let filtered_len_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, invalid_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pad_seen_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pad_count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "wasm:js-number", "toI32", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "charCodeAt",
        2,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, code_slot, line);

    emit_base64_whitespace(&mut chunks[current], code_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    emit_base64_char_valid(&mut chunks[current], code_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, invalid_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, code_slot, line);
    chunks[current].emit_i32_const(61, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pad_seen_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pad_count_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pad_count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pad_count_slot, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, invalid_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pad_seen_slot, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, invalid_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "charAt", 2, line);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    vybe_compiler::primitives::strings::emit_length(&mut chunks[current], line);
    host::emit(&mut chunks[current], "wasm:js-number", "toI32", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, filtered_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, filtered_len_slot, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_REM_S, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, invalid_slot, line);
    chunks[current].emit_op(Op::I32_OR, line);
    if throw_on_invalid {
        chunks[current].emit_if(line);
        emit_throw_dotnet_exception(
            &mut chunks[current],
            "FormatException",
            "The input is not a valid Base-64 string.",
            line,
        );
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_SET, invalid_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    invalid_slot
}

fn emit_insert_line_breaks(chunks: &mut [Chunk], current: usize, input_slot: u16, line: u32) {
    let out_slot = reserve_slot(&mut chunks[current]);
    let i_slot = reserve_slot(&mut chunks[current]);
    let len_slot = reserve_slot(&mut chunks[current]);
    let take_slot = reserve_slot(&mut chunks[current]);
    let remaining_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, remaining_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, remaining_slot, line);
    chunks[current].emit_i32_const(76, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(76, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, take_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, remaining_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, take_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, take_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "substr", 3, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, take_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_string_const("\r\n", line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_end(line);

    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

pub fn emit_convert_to_base64_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let bytes_slot = reserve_slot(&mut chunks[current]);
    let offset_slot = reserve_slot(&mut chunks[current]);
    let length_slot = reserve_slot(&mut chunks[current]);
    let end_slot = reserve_slot(&mut chunks[current]);
    let options_slot = reserve_slot(&mut chunks[current]);
    let result_slot = reserve_slot(&mut chunks[current]);

    match argc {
        3 => {
            chunks[current].emit_op_u16(Op::LOCAL_SET, length_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, offset_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
            emit_throw_if_null(&mut chunks[current], bytes_slot, line);
            emit_throw_if_slice_out_of_range(
                &mut chunks[current],
                bytes_slot,
                offset_slot,
                length_slot,
                line,
            );
            chunks[current].emit_op_u16(Op::LOCAL_GET, offset_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, length_slot, line);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
            emit_buffer_to_base64(
                chunks,
                current,
                bytes_slot,
                Some(offset_slot),
                Some(end_slot),
                line,
            );
        }
        2 => {
            chunks[current].emit_op_u16(Op::LOCAL_SET, options_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
            emit_throw_if_null(&mut chunks[current], bytes_slot, line);
            emit_buffer_to_base64(chunks, current, bytes_slot, None, None, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, options_slot, line);
            chunks[current].emit_string_const("__dotnet_base64_insertlinebreaks", line);
            vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
            chunks[current].emit_if_value(line);
            emit_insert_line_breaks(chunks, current, result_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
            chunks[current].emit_end(line);
        }
        _ => {
            chunks[current].emit_op_u16(Op::LOCAL_SET, bytes_slot, line);
            emit_throw_if_null(&mut chunks[current], bytes_slot, line);
            emit_buffer_to_base64(chunks, current, bytes_slot, None, None, line);
        }
    }
}

pub fn emit_convert_from_base64_string(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let text_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);
    emit_throw_if_null(&mut chunks[current], text_slot, line);
    emit_filter_dotnet_base64(chunks, current, text_slot, true, line);
    vybe_compiler::primitives::base64::emit_decode_binary_string(chunks, current, line);
    vybe_compiler::primitives::base64::emit_binary_string_to_byte_array(chunks, current, line);
}

fn emit_bool_int_pair(chunks: &mut [Chunk], current: usize, ok: bool, value_slot: u16, line: u32) {
    if ok {
        vybe_compiler::primitives::instructions::core_wasm::bool_const(
            &mut chunks[current],
            line,
            true,
        );
    } else {
        vybe_compiler::primitives::instructions::core_wasm::bool_const(
            &mut chunks[current],
            line,
            false,
        );
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::primitives::collections::emit_array_pair(chunks, current, line);
}

fn emit_false_zero_pair(chunks: &mut [Chunk], current: usize, line: u32) {
    let zero_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, zero_slot, line);
    emit_bool_int_pair(chunks, current, false, zero_slot, line);
}

fn emit_chars_to_string(chunks: &mut [Chunk], current: usize, source_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, source_slot, line);
    chunks[current].emit_string_const("", line);
    vybe_compiler::primitives::collections::emit_join(chunks, current, line);
    chunks[current].emit_end(line);
}

pub fn emit_convert_try_from_base64_chars(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let source_slot = reserve_slot(&mut chunks[current]);
    let dest_slot = reserve_slot(&mut chunks[current]);
    let text_slot = reserve_slot(&mut chunks[current]);
    let filtered_slot = reserve_slot(&mut chunks[current]);
    let decoded_slot = reserve_slot(&mut chunks[current]);
    let decoded_len_slot = reserve_slot(&mut chunks[current]);
    let dest_len_slot = reserve_slot(&mut chunks[current]);
    let i_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_op_u16(Op::LOCAL_SET, dest_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, source_slot, line);

    emit_chars_to_string(chunks, current, source_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, text_slot, line);

    let invalid_slot = emit_filter_dotnet_base64(chunks, current, text_slot, false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, filtered_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, invalid_slot, line);
    chunks[current].emit_if_value(line);
    emit_false_zero_pair(chunks, current, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, filtered_slot, line);
    vybe_compiler::primitives::base64::emit_decode_binary_string(chunks, current, line);
    vybe_compiler::primitives::base64::emit_binary_string_to_byte_array(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, decoded_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, decoded_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, decoded_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dest_len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, decoded_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest_len_slot, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    emit_false_zero_pair(chunks, current, line);
    chunks[current].emit_else(line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, decoded_len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dest_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, decoded_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    emit_bool_int_pair(chunks, current, true, decoded_len_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_convert_to_base64_char_array(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    let input_slot = reserve_slot(&mut chunks[current]);
    let offset_slot = reserve_slot(&mut chunks[current]);
    let length_slot = reserve_slot(&mut chunks[current]);
    let out_slot = reserve_slot(&mut chunks[current]);
    let out_offset_slot = reserve_slot(&mut chunks[current]);
    let end_slot = reserve_slot(&mut chunks[current]);
    let b64_slot = reserve_slot(&mut chunks[current]);
    let count_slot = reserve_slot(&mut chunks[current]);
    let i_slot = reserve_slot(&mut chunks[current]);

    chunks[current].emit_op_u16(Op::LOCAL_SET, out_offset_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, length_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, offset_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, input_slot, line);
    emit_throw_if_null(&mut chunks[current], input_slot, line);
    emit_throw_if_null(&mut chunks[current], out_slot, line);
    emit_throw_if_slice_out_of_range(
        &mut chunks[current],
        input_slot,
        offset_slot,
        length_slot,
        line,
    );

    chunks[current].emit_op_u16(Op::LOCAL_GET, offset_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, length_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, end_slot, line);
    emit_buffer_to_base64(
        chunks,
        current,
        input_slot,
        Some(offset_slot),
        Some(end_slot),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, b64_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b64_slot, line);
    host::emit(&mut chunks[current], "wasm:js-string", "length", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_offset_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b64_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    host::emit(&mut chunks[current], "ecma:string", "charAt", 2, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
}
