//! Shared binary string packing helpers.
//!
//! This module is for byte-level encode/decode mechanics shared by language
//! adapters such as Lua `string.pack`, Ruby `Array#pack`/`String#unpack`, PHP
//! `pack`/`unpack`, and Python `struct`-style surfaces.
//!
//! It is intentionally not a multi-value/list-context layer. Multi-value
//! controls how many values flow through calls/returns/assignments; packing
//! controls how values become bytes and how bytes become values.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

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
    // Truncate before extracting the low byte. Adding 256 to a negative
    // fraction before truncation gave the wrong byte (e.g. -1.5 -> 254).
    // Reduction by a power of two also handles larger finite magnitudes.
    let integer = chunk.alloc_scratch(1);
    local_get(chunk, value_slot, line);
    chunk.emit_op(Op::F64_TRUNC, line);
    local_set(chunk, integer, line);
    local_get(chunk, integer, line);
    local_get(chunk, integer, line);
    chunk.emit_f64_const(256.0, line);
    chunk.emit_op(Op::F64_DIV, line);
    chunk.emit_op(Op::F64_FLOOR, line);
    chunk.emit_f64_const(256.0, line);
    chunk.emit_op(Op::F64_MUL, line);
    chunk.emit_op(Op::F64_SUB, line);
    // Non-finite inputs retain unsigned Wasm truncation's trap contract.
    chunk.emit_op(Op::I32_TRUNC_F64_U, line);
}

fn emit_i32_byte_to_string(chunk: &mut Chunk, from_char_code: u16, line: u32) {
    chunk.emit_i32_const(255, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_call(from_char_code, 1, line);
}

/// Pack a proven f64 local as one byte, truncating and retaining its low byte.
/// Non-finite inputs trap in unsigned Wasm truncation.
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

/// Pack the low 16 bits of a proven f64 local after unsigned i32 truncation.
/// Language adapters own wrapping/coercion before this checked numeric entry point.
///
/// Stack: `[] -> [two_char_string]`.
pub fn emit_pack_u16_from_f64_slot(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    endian: Endian,
    line: u32,
) {
    emit_pack_unsigned_string(chunks, current, value_slot, 2, endian, line);
}

/// Pack a proven f64 local after unsigned i32 truncation.
/// Language adapters own wrapping/coercion before this checked numeric entry point.
///
/// Stack: `[] -> [four_char_string]`.
pub fn emit_pack_u32_from_f64_slot(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    endian: Endian,
    line: u32,
) {
    emit_pack_unsigned_string(chunks, current, value_slot, 4, endian, line);
}

fn emit_pack_unsigned_string(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    bytes: usize,
    endian: Endian,
    line: u32,
) {
    // The _f64_slot entry points require a numeric producer. Preserve their
    // unsigned Wasm truncation/trap contract, doing it once before any output.
    let bits = chunks[current].alloc_scratch(1);
    local_get(&mut chunks[current], value_slot, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    local_set(&mut chunks[current], bits, line);
    let from_char_code = chunks[current].add_import("wasm:js-string", "fromCharCode");
    for index in 0..bytes {
        let byte_index = match endian {
            Endian::Little => index,
            Endian::Big => bytes - index - 1,
        };
        local_get(&mut chunks[current], bits, line);
        if byte_index != 0 {
            chunks[current].emit_i32_const((byte_index * 8) as i32, line);
            chunks[current].emit_op(Op::I32_SHR_U, line);
        }
        emit_i32_byte_to_string(&mut chunks[current], from_char_code, line);
        if index != 0 {
            crate::primitives::strings::emit_str_concat(&mut chunks[current], line);
        }
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

fn emit_slot_to_u32_i32(chunk: &mut Chunk, value_slot: u16, line: u32) {
    local_get(chunk, value_slot, line);
    let to_number = chunk.add_import("ecma:value", "toNumber");
    chunk.emit_call(to_number, 1, line);
    chunk.emit_op(Op::I32_TRUNC_F64_U, line);
}

fn emit_store_array_byte(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    index: i32,
    byte: u16,
    line: u32,
) {
    local_get(&mut chunks[current], array_slot, line);
    chunks[current].emit_i32_const(index, line);
    local_get(&mut chunks[current], byte, line);
    crate::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_load_array_byte_f64(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    index: i32,
    line: u32,
) {
    local_get(&mut chunks[current], array_slot, line);
    chunks[current].emit_i32_const(index, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    let to_number = chunks[current].add_import("ecma:value", "toNumber");
    chunks[current].emit_call(to_number, 1, line);
}

/// Store the low 16 bits of a numeric local into a byte array.
///
/// Stack: `[] -> []`. The array is mutated through the shared collection setter.
pub fn emit_store_u16_to_array_from_number_slot(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    value_slot: u16,
    endian: Endian,
    line: u32,
) {
    let lo = chunks[current].alloc_scratch(1);
    let hi = chunks[current].alloc_scratch(1);

    emit_slot_to_u32_i32(&mut chunks[current], value_slot, line);
    chunks[current].emit_i32_const(255, line);
    chunks[current].emit_op(Op::I32_AND, line);
    local_set(&mut chunks[current], lo, line);

    emit_slot_to_u32_i32(&mut chunks[current], value_slot, line);
    chunks[current].emit_i32_const(8, line);
    chunks[current].emit_op(Op::I32_SHR_U, line);
    chunks[current].emit_i32_const(255, line);
    chunks[current].emit_op(Op::I32_AND, line);
    local_set(&mut chunks[current], hi, line);

    let order = match endian {
        Endian::Little => [lo, hi],
        Endian::Big => [hi, lo],
    };
    emit_store_array_byte(chunks, current, array_slot, 0, order[0], line);
    emit_store_array_byte(chunks, current, array_slot, 1, order[1], line);
}

/// Store the low 32 bits of a numeric local into a byte array.
///
/// Stack: `[] -> []`. The array is mutated through the shared collection setter.
pub fn emit_store_u32_to_array_from_number_slot(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    value_slot: u16,
    endian: Endian,
    line: u32,
) {
    let b0 = chunks[current].alloc_scratch(1);
    let b1 = chunks[current].alloc_scratch(1);
    let b2 = chunks[current].alloc_scratch(1);
    let b3 = chunks[current].alloc_scratch(1);

    for (slot, shift) in [(b0, 0), (b1, 8), (b2, 16), (b3, 24)] {
        emit_slot_to_u32_i32(&mut chunks[current], value_slot, line);
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
    for (index, slot) in order.iter().copied().enumerate() {
        emit_store_array_byte(chunks, current, array_slot, index as i32, slot, line);
    }
}

/// Store a split 64-bit value (hi/lo u32) into a byte array.
///
/// Stack: `[] -> []`. This avoids routing 64-bit integer literals through f64.
pub fn emit_store_u64_parts_to_array_from_number_slots(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    hi_slot: u16,
    lo_slot: u16,
    endian: Endian,
    line: u32,
) {
    let b0 = chunks[current].alloc_scratch(1);
    let b1 = chunks[current].alloc_scratch(1);
    let b2 = chunks[current].alloc_scratch(1);
    let b3 = chunks[current].alloc_scratch(1);
    let b4 = chunks[current].alloc_scratch(1);
    let b5 = chunks[current].alloc_scratch(1);
    let b6 = chunks[current].alloc_scratch(1);
    let b7 = chunks[current].alloc_scratch(1);

    for (slot, source, shift) in [
        (b0, hi_slot, 24),
        (b1, hi_slot, 16),
        (b2, hi_slot, 8),
        (b3, hi_slot, 0),
        (b4, lo_slot, 24),
        (b5, lo_slot, 16),
        (b6, lo_slot, 8),
        (b7, lo_slot, 0),
    ] {
        emit_slot_to_u32_i32(&mut chunks[current], source, line);
        if shift > 0 {
            chunks[current].emit_i32_const(shift, line);
            chunks[current].emit_op(Op::I32_SHR_U, line);
        }
        chunks[current].emit_i32_const(255, line);
        chunks[current].emit_op(Op::I32_AND, line);
        local_set(&mut chunks[current], slot, line);
    }

    let order = match endian {
        Endian::Big => [b0, b1, b2, b3, b4, b5, b6, b7],
        Endian::Little => [b7, b6, b5, b4, b3, b2, b1, b0],
    };
    for (index, slot) in order.iter().copied().enumerate() {
        emit_store_array_byte(chunks, current, array_slot, index as i32, slot, line);
    }
}

/// Read a u16 from a byte array as an f64.
pub fn emit_load_u16_from_array_f64(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    endian: Endian,
    line: u32,
) {
    match endian {
        Endian::Little => {
            emit_load_array_byte_f64(chunks, current, array_slot, 0, line);
            emit_load_array_byte_f64(chunks, current, array_slot, 1, line);
            chunks[current].emit_f64_const(256.0, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
        Endian::Big => {
            emit_load_array_byte_f64(chunks, current, array_slot, 0, line);
            chunks[current].emit_f64_const(256.0, line);
            chunks[current].emit_op(Op::F64_MUL, line);
            emit_load_array_byte_f64(chunks, current, array_slot, 1, line);
            chunks[current].emit_op(Op::F64_ADD, line);
        }
    }
}

/// Read a u32 from a byte array as an f64.
pub fn emit_load_u32_from_array_f64(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    endian: Endian,
    line: u32,
) {
    let order = match endian {
        Endian::Little => [0, 1, 2, 3],
        Endian::Big => [3, 2, 1, 0],
    };
    emit_load_array_byte_f64(chunks, current, array_slot, order[0], line);
    emit_load_array_byte_f64(chunks, current, array_slot, order[1], line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_load_array_byte_f64(chunks, current, array_slot, order[2], line);
    chunks[current].emit_f64_const(65536.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_load_array_byte_f64(chunks, current, array_slot, order[3], line);
    chunks[current].emit_f64_const(16777216.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

/// Store four bytes from a checked numeric producer; preserve unsigned truncation.
/// Caller owns numeric proof. No user coercion occurs after that proof.
pub fn emit_store_u32_to_array_from_f64_slot(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    value_slot: u16,
    endian: Endian,
    line: u32,
) {
    let bits = chunks[current].alloc_scratch(1);
    local_get(&mut chunks[current], value_slot, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    local_set(&mut chunks[current], bits, line);
    let order = match endian {
        Endian::Little => [0, 8, 16, 24],
        Endian::Big => [24, 16, 8, 0],
    };
    let byte = chunks[current].alloc_scratch(1);
    for (index, shift) in order.iter().copied().enumerate() {
        local_get(&mut chunks[current], bits, line);
        if shift != 0 {
            chunks[current].emit_i32_const(shift, line);
            chunks[current].emit_op(Op::I32_SHR_U, line);
        }
        chunks[current].emit_i32_const(255, line);
        chunks[current].emit_op(Op::I32_AND, line);
        local_set(&mut chunks[current], byte, line);
        emit_store_array_byte(chunks, current, array_slot, index as i32, byte, line);
    }
}
