//! Shared binary string packing helpers.
//!
//! This module is for byte-level encode/decode mechanics shared by language
//! adapters such as Lua `string.pack`, Ruby `Array#pack`/`String#unpack`, PHP
//! `pack`/`unpack`, and Python `struct`-style surfaces.
//!
//! It is intentionally not a multi-value/list-context layer. Multi-value
//! controls how many values flow through calls/returns/assignments; packing
//! controls how values become bytes and how bytes become values.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Endian {
    Little,
    Big,
}

fn local_get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn local_set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_value_slot_to_masked_i32(chunk: &mut Chunk, value_slot: u16, line: u32) {
    local_get(chunk, value_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if(line);
    local_get(chunk, value_slot, line);
    chunk.emit_f64_const(256.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_else(line);
    local_get(chunk, value_slot, line);
    chunk.emit_end(line);
    chunk.emit_op(Op::I32_TRUNC_F64_U, line);
}

fn emit_i32_byte_to_string(chunk: &mut Chunk, from_char_code: u16, line: u32) {
    chunk.emit_i32_const(255, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_call(from_char_code, 1, line);
}

/// Pack a numeric local as one byte.
///
/// Stack: `[] -> [one_char_string]`.
pub fn emit_pack_byte_from_f64_slot(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    line: u32,
) {
    let from_char_code = chunks[current].add_import("ecma:string", "fromCharCode");
    emit_value_slot_to_masked_i32(&mut chunks[current], value_slot, line);
    emit_i32_byte_to_string(&mut chunks[current], from_char_code, line);
}

/// Pack a numeric local as an unsigned 16-bit value.
///
/// Stack: `[] -> [two_char_string]`.
pub fn emit_pack_u16_from_f64_slot(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    endian: Endian,
    line: u32,
) {
    let from_char_code = chunks[current].add_import("ecma:string", "fromCharCode");
    let lo_slot = chunks[current].alloc_scratch(1);
    let hi_slot = chunks[current].alloc_scratch(1);

    local_get(&mut chunks[current], value_slot, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    chunks[current].emit_i32_const(255, line);
    chunks[current].emit_op(Op::I32_AND, line);
    local_set(&mut chunks[current], lo_slot, line);

    local_get(&mut chunks[current], value_slot, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    chunks[current].emit_i32_const(8, line);
    chunks[current].emit_op(Op::I32_SHR_U, line);
    chunks[current].emit_i32_const(255, line);
    chunks[current].emit_op(Op::I32_AND, line);
    local_set(&mut chunks[current], hi_slot, line);

    let (first, second) = match endian {
        Endian::Little => (lo_slot, hi_slot),
        Endian::Big => (hi_slot, lo_slot),
    };
    local_get(&mut chunks[current], first, line);
    chunks[current].emit_call(from_char_code, 1, line);
    local_get(&mut chunks[current], second, line);
    chunks[current].emit_call(from_char_code, 1, line);
    crate::strings::emit_str_concat(&mut chunks[current], line);
}

/// Pack a numeric local as an unsigned 32-bit value.
///
/// Stack: `[] -> [four_char_string]`.
pub fn emit_pack_u32_from_f64_slot(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    endian: Endian,
    line: u32,
) {
    let from_char_code = chunks[current].add_import("ecma:string", "fromCharCode");
    let b0 = chunks[current].alloc_scratch(1);
    let b1 = chunks[current].alloc_scratch(1);
    let b2 = chunks[current].alloc_scratch(1);
    let b3 = chunks[current].alloc_scratch(1);

    for (slot, shift) in [(b0, 0), (b1, 8), (b2, 16), (b3, 24)] {
        local_get(&mut chunks[current], value_slot, line);
        chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
        if shift > 0 {
            chunks[current].emit_i32_const(shift, line);
            chunks[current].emit_op(Op::I32_SHR_U, line);
        }
        chunks[current].emit_i32_const(255, line);
        chunks[current].emit_op(Op::I32_AND, line);
        local_set(&mut chunks[current], slot, line);
    }

    let order = match endian {
        Endian::Little => [b0, b1, b2, b3],
        Endian::Big => [b3, b2, b1, b0],
    };
    local_get(&mut chunks[current], order[0], line);
    chunks[current].emit_call(from_char_code, 1, line);
    for slot in order.iter().copied().skip(1) {
        local_get(&mut chunks[current], slot, line);
        chunks[current].emit_call(from_char_code, 1, line);
        crate::strings::emit_str_concat(&mut chunks[current], line);
    }
}

/// Read one byte from a string local at a zero-based numeric index.
///
/// Stack: `[] -> [code_as_f64]`.
pub fn emit_char_code_at_zero_f64(
    chunks: &mut [Chunk],
    current: usize,
    string_slot: u16,
    index: f64,
    line: u32,
) {
    local_get(&mut chunks[current], string_slot, line);
    chunks[current].emit_f64_const(index, line);
    let char_code_at = chunks[current].add_import("wasm:js-string", "charCodeAt");
    chunks[current].emit_call(char_code_at, 2, line);
}

/// Read one byte from a string local at a zero-based i32 index local.
///
/// Stack: `[] -> [code_as_i32]`.
pub fn emit_char_code_at_i32_slot(
    chunks: &mut [Chunk],
    current: usize,
    string_slot: u16,
    index_slot: u16,
    line: u32,
) {
    local_get(&mut chunks[current], string_slot, line);
    local_get(&mut chunks[current], index_slot, line);
    let char_code_at = chunks[current].add_import("ecma:string", "charCodeAt");
    chunks[current].emit_call(char_code_at, 2, line);
}

/// Read one byte from a string local at a zero-based i32 constant index.
///
/// Stack: `[] -> [code_as_i32]`.
pub fn emit_char_code_at_i32_const(
    chunks: &mut [Chunk],
    current: usize,
    string_slot: u16,
    index: i32,
    line: u32,
) {
    local_get(&mut chunks[current], string_slot, line);
    chunks[current].emit_i32_const(index, line);
    let char_code_at = chunks[current].add_import("ecma:string", "charCodeAt");
    chunks[current].emit_call(char_code_at, 2, line);
}

/// Read one byte from a string local at a one-based numeric index local.
///
/// Stack: `[] -> [code_as_f64]`.
pub fn emit_char_code_at_one_based_pos_f64(
    chunks: &mut [Chunk],
    current: usize,
    string_slot: u16,
    pos_slot: u16,
    line: u32,
) {
    local_get(&mut chunks[current], string_slot, line);
    local_get(&mut chunks[current], pos_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    let char_code_at = chunks[current].add_import("wasm:js-string", "charCodeAt");
    chunks[current].emit_call(char_code_at, 2, line);
}

/// Unpack two bytes from a string local as an unsigned 16-bit numeric value.
///
/// Stack: `[] -> [value_as_f64]`.
pub fn emit_unpack_u16_from_string_slot_f64(
    chunks: &mut [Chunk],
    current: usize,
    string_slot: u16,
    endian: Endian,
    line: u32,
) {
    match endian {
        Endian::Little => {
            emit_char_code_at_zero_f64(chunks, current, string_slot, 0.0, line);
            emit_char_code_at_zero_f64(chunks, current, string_slot, 1.0, line);
            chunks[current].emit_f64_const(256.0, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
        Endian::Big => {
            emit_char_code_at_zero_f64(chunks, current, string_slot, 0.0, line);
            chunks[current].emit_f64_const(256.0, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            emit_char_code_at_zero_f64(chunks, current, string_slot, 1.0, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
    }
}

/// Unpack two bytes from a string local as an unsigned 16-bit i32 value.
///
/// Stack: `[] -> [value_as_i32]`.
pub fn emit_unpack_u16_from_string_slot_i32(
    chunks: &mut [Chunk],
    current: usize,
    string_slot: u16,
    endian: Endian,
    line: u32,
) {
    match endian {
        Endian::Little => {
            emit_char_code_at_i32_const(chunks, current, string_slot, 1, line);
            chunks[current].emit_i32_const(256, line);
            chunks[current].emit_op(Op::I32_MUL, line);
            emit_char_code_at_i32_const(chunks, current, string_slot, 0, line);
            chunks[current].emit_op(Op::I32_ADD, line);
        }
        Endian::Big => {
            emit_char_code_at_i32_const(chunks, current, string_slot, 0, line);
            chunks[current].emit_i32_const(256, line);
            chunks[current].emit_op(Op::I32_MUL, line);
            emit_char_code_at_i32_const(chunks, current, string_slot, 1, line);
            chunks[current].emit_op(Op::I32_ADD, line);
        }
    }
}

/// Unpack four bytes from a string local as an unsigned 32-bit numeric value.
///
/// Stack: `[] -> [value_as_f64]`.
pub fn emit_unpack_u32_from_string_slot_f64(
    chunks: &mut [Chunk],
    current: usize,
    string_slot: u16,
    endian: Endian,
    line: u32,
) {
    let order = match endian {
        Endian::Little => [0.0, 1.0, 2.0, 3.0],
        Endian::Big => [3.0, 2.0, 1.0, 0.0],
    };
    emit_char_code_at_zero_f64(chunks, current, string_slot, order[0], line);
    emit_char_code_at_zero_f64(chunks, current, string_slot, order[1], line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_char_code_at_zero_f64(chunks, current, string_slot, order[2], line);
    chunks[current].emit_f64_const(65536.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_char_code_at_zero_f64(chunks, current, string_slot, order[3], line);
    chunks[current].emit_f64_const(16777216.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}
