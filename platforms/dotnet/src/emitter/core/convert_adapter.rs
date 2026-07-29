use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;
use vybe_compiler::primitives::instructions::host;

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn emit_throw_dotnet_exception(chunk: &mut Chunk, exception_name: &str, message: &str, line: u32) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(message, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(chunk, exception_name, line);
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
    chunk: &mut Chunk,
    bytes_slot: u16,
    start_slot: Option<u16>,
    end_slot: Option<u16>,
    line: u32,
) {
    let buffer_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    host::emit(chunk, "wasm:js-string", "test", 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_string_const("utf-8", line);
    host::emit(chunk, "node:buffer", "from", 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buffer_slot, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, bytes_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, buffer_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, buffer_slot, line);
    chunk.emit_string_const("base64", line);
    if let (Some(start_slot), Some(end_slot)) = (start_slot, end_slot) {
        chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
        host::emit(chunk, "node:buffer", "toString", 4, line);
    } else {
        host::emit(chunk, "node:buffer", "toString", 2, line);
    }
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
                &mut chunks[current],
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
            emit_buffer_to_base64(&mut chunks[current], bytes_slot, None, None, line);
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
            emit_buffer_to_base64(&mut chunks[current], bytes_slot, None, None, line);
        }
    }
}

pub fn emit_convert_from_base64_string(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let text_slot = reserve_slot(chunk);
    let result_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, text_slot, line);
    emit_throw_if_null(chunk, text_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, text_slot, line);
    host::emit(chunk, "node:buffer", "fromBase64Strict", 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "FormatException",
        "The input is not a valid Base-64 string.",
        line,
    );
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
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
        &mut chunks[current],
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
