//! Python `ascii()` adapter.
//!
//! The surface is Python-specific, but the representation step remains shared
//! with `repr()`. This adapter only adds Python's non-ASCII escape pass.

use vybe_compiler::primitives::{loops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn append_stack_piece(chunks: &mut [Chunk], current: usize, out: u16, line: u32) {
    let piece = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], piece, line);
    lget(&mut chunks[current], out, line);
    lget(&mut chunks[current], piece, line);
    strings::emit_concat(&mut chunks[current], 2, line);
    lset(&mut chunks[current], out, line);
}

fn emit_hex_digit(chunk: &mut Chunk, code: u16, shift: i32, line: u32) {
    let nibble = chunk.alloc_scratch(1);
    lget(chunk, code, line);
    if shift > 0 {
        chunk.emit_i32_const(shift, line);
        chunk.emit_op(Op::I32_SHR_U, line);
    }
    chunk.emit_i32_const(15, line);
    chunk.emit_op(Op::I32_AND, line);
    lset(chunk, nibble, line);

    lget(chunk, nibble, line);
    chunk.emit_i32_const(9, line);
    chunk.emit_op(Op::I32_LE_U, line);
    chunk.emit_if_value(line);
    lget(chunk, nibble, line);
    chunk.emit_i32_const(48, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_else(line);
    lget(chunk, nibble, line);
    chunk.emit_i32_const(87, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_end(line);

    let from_char = chunk.add_import("wasm:js-string", "fromCharCode");
    chunk.emit_call(from_char, 1, line);
}

fn emit_hex_escape(
    chunks: &mut [Chunk],
    current: usize,
    prefix: &str,
    code: u16,
    digits: usize,
    line: u32,
) {
    chunks[current].emit_string_const(prefix, line);
    for digit in (0..digits).rev() {
        emit_hex_digit(&mut chunks[current], code, (digit * 4) as i32, line);
    }
    strings::emit_concat(&mut chunks[current], digits + 1, line);
}

pub fn emit_ascii(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }

    crate::emitter::runtime_adapter::emit_repr(chunks, current, 1, line);

    let text = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let code = chunks[current].alloc_scratch(1);

    lset(&mut chunks[current], text, line);
    chunks[current].emit_string_const("", line);
    lset(&mut chunks[current], out, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i, line);
    lget(&mut chunks[current], text, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    lset(&mut chunks[current], len, line);

    let loop_state = loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i, line);
    lget(&mut chunks[current], len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], text, line);
    lget(&mut chunks[current], i, line);
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    lset(&mut chunks[current], code, line);

    lget(&mut chunks[current], code, line);
    chunks[current].emit_i32_const(127, line);
    chunks[current].emit_op(Op::I32_LE_U, line);
    chunks[current].emit_if(line);
    {
        lget(&mut chunks[current], text, line);
        lget(&mut chunks[current], i, line);
        call_import(chunks, current, "ecma:string", "charAt", 2, line);
        append_stack_piece(chunks, current, out, line);
    }
    chunks[current].emit_else(line);
    {
        lget(&mut chunks[current], code, line);
        chunks[current].emit_i32_const(255, line);
        chunks[current].emit_op(Op::I32_LE_U, line);
        chunks[current].emit_if(line);
        emit_hex_escape(chunks, current, "\\x", code, 2, line);
        append_stack_piece(chunks, current, out, line);
        chunks[current].emit_else(line);
        emit_hex_escape(chunks, current, "\\u", code, 4, line);
        append_stack_piece(chunks, current, out, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);

    lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i, line);
    loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], out, line);
}
