//! Java UUID adapters over the existing ECMA string/crypto surface.

use crate::emitter::{instructions::host, strings};
use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn substring(chunk: &mut Chunk, line: u32, value: u16, start: i32, end: i32) {
    get(chunk, value, line);
    chunk.emit_i32_const(start, line);
    chunk.emit_i32_const(end, line);
    host::emit(chunk, "ecma:string", "substring", 3, line);
}

fn concat(chunk: &mut Chunk, line: u32) {
    strings::emit_str_concat(chunk, line);
}

fn char_code_at(chunk: &mut Chunk, line: u32, value: u16, index: i32) {
    get(chunk, value, line);
    chunk.emit_i32_const(index, line);
    host::emit(chunk, "wasm:js-string", "charCodeAt", 2, line);
}

fn signed_bit_is_set(chunk: &mut Chunk, line: u32, value: u16, index: i32) {
    char_code_at(chunk, line, value, index);
    chunk.emit_i32_const('8' as i32, line);
    chunk.emit_op(Op::I32_GE_S, line);
}

pub fn emit_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_string_const("00000000-0000-0000-0000-000000000000", line);
}

pub fn emit_from_string(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(&mut chunks[current], "ecma:string", "toLowerCase", 1, line);
}

pub fn emit_name_from_bytes(chunks: &mut [Chunk], current: usize, line: u32) {
    let bytes = chunks[current].alloc_scratch(1);
    let hex = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], bytes, line);
    get(&mut chunks[current], bytes, line);
    host::emit(&mut chunks[current], "node:crypto", "md5", 1, line);
    set(&mut chunks[current], hex, line);

    substring(&mut chunks[current], line, hex, 0, 8);
    chunks[current].emit_string_const("-", line);
    concat(&mut chunks[current], line);
    substring(&mut chunks[current], line, hex, 8, 12);
    concat(&mut chunks[current], line);
    chunks[current].emit_string_const("-3", line);
    concat(&mut chunks[current], line);
    substring(&mut chunks[current], line, hex, 13, 16);
    concat(&mut chunks[current], line);
    chunks[current].emit_string_const("-a", line);
    concat(&mut chunks[current], line);
    substring(&mut chunks[current], line, hex, 17, 20);
    concat(&mut chunks[current], line);
    chunks[current].emit_string_const("-", line);
    concat(&mut chunks[current], line);
    substring(&mut chunks[current], line, hex, 20, 32);
    concat(&mut chunks[current], line);
}

pub fn emit_version(chunks: &mut [Chunk], current: usize, line: u32) {
    let uuid = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], uuid, line);
    char_code_at(&mut chunks[current], line, uuid, 14);
    chunks[current].emit_i32_const('0' as i32, line);
    chunks[current].emit_op(Op::I32_SUB, line);
}

pub fn emit_variant(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(2, line);
}

pub fn emit_most_bits(chunks: &mut [Chunk], current: usize, line: u32) {
    let uuid = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], uuid, line);
    signed_bit_is_set(&mut chunks[current], line, uuid, 0);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i64_const(-1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i64_const(0, line);
    chunks[current].emit_end(line);
}

pub fn emit_least_bits(chunks: &mut [Chunk], current: usize, line: u32) {
    let uuid = chunks[current].alloc_scratch(1);
    let tail = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], uuid, line);

    signed_bit_is_set(&mut chunks[current], line, uuid, 19);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i64_const(-1, line);
    chunks[current].emit_else(line);
    signed_bit_is_set(&mut chunks[current], line, uuid, 24);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i64_const(-1, line);
    chunks[current].emit_else(line);
    substring(&mut chunks[current], line, uuid, 24, 36);
    set(&mut chunks[current], tail, line);
    get(&mut chunks[current], tail, line);
    chunks[current].emit_i32_const(16, line);
    host::emit(&mut chunks[current], "ecma:number", "parseInt", 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    host::emit(
        &mut chunks[current],
        "ecma:string",
        "localeCompare",
        2,
        line,
    );
}

pub fn emit_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(0, line);
}
