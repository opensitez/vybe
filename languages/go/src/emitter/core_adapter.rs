use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::{canon_marshal, collections, convert, loops, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.core_bytes_to_string" if argc == 1 => emit_bytes_to_string(chunks, current, line),
        "go.core_string_to_bytes" if argc == 1 => emit_string_to_bytes(chunks, current, line),
        "go.core_string_byte_len" if argc == 1 => emit_string_byte_len(chunks, current, line),
        "go.core_string_to_runes" if argc == 1 => emit_string_to_runes(chunks, current, line),
        "go.core_runes_to_string" if argc == 1 => emit_runes_to_string(chunks, current, line),
        "go.core_rune_value" if argc == 1 => emit_rune_value(chunks, current, line),
        _ => return false,
    }
    true
}

fn emit_string_to_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let s = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    host::emit(&mut chunks[current], "web:encoding", "encoderNew", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    host::emit(&mut chunks[current], "web:encoding", "encode", 2, line);
    host::emit(&mut chunks[current], "ecma:array", "from", 1, line);
}

fn emit_string_byte_len(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_string_to_bytes(chunks, current, line);
    collections::emit_len(chunks, current, line);
    convert::emit_to_int(&mut chunks[current], line);
}

fn emit_bytes_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    canon_marshal::emit_decode_utf8(&mut chunks[current], line);
}

fn emit_string_to_runes(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let s = base;
    let chars = base + 1;
    let out = base + 2;
    let i = base + 3;
    let len = base + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, s, line);
    host::emit(&mut chunks[current], "ecma:array", "from", 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, chars, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars, line);
    collections::emit_len(chunks, current, line);
    convert::emit_to_int(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let loop_state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, chars, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "codePointAt",
        2,
        line,
    );
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    loops::emit_loop_end(chunks, current, loop_state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_runes_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let runes = base;
    let out = base + 1;
    let i = base + 2;
    let len = base + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, runes, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, runes, line);
    collections::emit_len(chunks, current, line);
    convert::emit_to_int(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let loop_state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, runes, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    host::emit(
        &mut chunks[current],
        "wasm:js-string",
        "fromCodePoint",
        1,
        line,
    );
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    loops::emit_loop_end(chunks, current, loop_state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_rune_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "wasm:js-string", "test", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    {
        chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
        chunks[current].emit_i32_const(0, line);
        host::emit(
            &mut chunks[current],
            "wasm:js-string",
            "codePointAt",
            2,
            line,
        );
    }
    chunks[current].emit_else(line);
    {
        chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
        convert::emit_to_int(&mut chunks[current], line);
    }
    chunks[current].emit_end(line);
}
