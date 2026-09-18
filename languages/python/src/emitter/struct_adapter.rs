//! Python `struct` module adapter.
//!
//! Python owns the format grammar and tuple result shape. Shared byte/endian
//! mechanics come from `vybe_compiler::primitives::packing`.

use vybe_compiler::primitives::class_slots::{self, ClassSlot, ObjSource, PlainNames, ValueSource};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

fn emit_slot_eq_str(chunk: &mut Chunk, slot: u16, value: &str, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_string_const(value, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_concat(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
}

fn emit_bool_byte_from_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    lget(&mut chunks[current], slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("\x01", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("\0", line);
    chunks[current].emit_end(line);
}

fn emit_byte_bool_from_top(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_little_endian_flag_from_big_slot(chunk: &mut Chunk, big_slot: u16, line: u32) {
    lget(chunk, big_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_end(line);
}

fn emit_dynamic_byte_at_const(
    chunks: &mut [Chunk],
    current: usize,
    data_slot: u16,
    index: i32,
    line: u32,
) {
    lget(&mut chunks[current], data_slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::packing::emit_char_code_at_i32_const(
        chunks, current, data_slot, index, line,
    );
    chunks[current].emit_else(line);
    lget(&mut chunks[current], data_slot, line);
    chunks[current].emit_i32_const(index, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_dynamic_byte_at_offset_const(
    chunks: &mut [Chunk],
    current: usize,
    data_slot: u16,
    offset_slot: u16,
    add: f64,
    line: u32,
) {
    lget(&mut chunks[current], data_slot, line);
    call_import(chunks, current, "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], data_slot, line);
    lget(&mut chunks[current], offset_slot, line);
    chunks[current].emit_f64_const(add, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], data_slot, line);
    lget(&mut chunks[current], offset_slot, line);
    chunks[current].emit_f64_const(add, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_nul_string(chunk: &mut Chunk, count: usize, line: u32) {
    let s = "\0".repeat(count);
    chunk.emit_string_const(&s, line);
}

fn emit_bytes_from_stack_array(chunks: &mut [Chunk], current: usize, count: u16, line: u32) {
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, count, line);
    call_import(chunks, current, "ecma:uint8array", "new", 1, line);
}

fn emit_adjust_i16_from_u16(chunk: &mut Chunk, value_slot: u16, line: u32) {
    lget(chunk, value_slot, line);
    chunk.emit_f64_const(32767.0, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    chunk.emit_f64_const(65536.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    chunk.emit_else(line);
    lget(chunk, value_slot, line);
    chunk.emit_end(line);
}

fn emit_pack_i16_slot(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let adjusted = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_f64_const(65536.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], adjusted, line);
    vybe_compiler::primitives::packing::emit_pack_u16_from_f64_slot(
        chunks,
        current,
        adjusted,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
}

fn emit_tuple_from_top(chunks: &mut [Chunk], current: usize, n: u16, line: u32) {
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, n, line);
}

fn emit_throw_exception(chunk: &mut Chunk, message: &str, line: u32) {
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(message, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(chunk, "Exception", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

fn struct_set(chunk: &mut Chunk, key: &ClassSlot, line: u32) {
    let slot = class_slots::resolve(key, &PlainNames);
    class_slots::emit_class_set(chunk, ObjSource::Stack, &slot, ValueSource::Stack, line);
}

fn emit_unpack_u32_at_offset(
    chunks: &mut [Chunk],
    current: usize,
    data: u16,
    offset: u16,
    line: u32,
) {
    emit_dynamic_byte_at_offset_const(chunks, current, data, offset, 0.0, line);
    emit_dynamic_byte_at_offset_const(chunks, current, data, offset, 1.0, line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_dynamic_byte_at_offset_const(chunks, current, data, offset, 2.0, line);
    chunks[current].emit_f64_const(65536.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_dynamic_byte_at_offset_const(chunks, current, data, offset, 3.0, line);
    chunks[current].emit_f64_const(16777216.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
}

pub fn emit_struct_pack(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 1 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let fmt = base;
    let first = base + 1;

    emit_slot_eq_str(&mut chunks[current], fmt, "i", line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
        chunks,
        current,
        first,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "@i", line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
        chunks,
        current,
        first,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "h", line);
    chunks[current].emit_if_value(line);
    emit_pack_i16_slot(chunks, current, first, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "<H", line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::packing::emit_pack_u16_from_f64_slot(
        chunks,
        current,
        first,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "iii", line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
        chunks,
        current,
        first,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
        chunks,
        current,
        first + 1,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    emit_concat(chunks, current, line);
    vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
        chunks,
        current,
        first + 2,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    emit_concat(chunks, current, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "ii", line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
        chunks,
        current,
        first,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
        chunks,
        current,
        first + 1,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    emit_concat(chunks, current, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "?", line);
    chunks[current].emit_if_value(line);
    emit_bool_byte_from_slot(chunks, current, first, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, ">??", line);
    chunks[current].emit_if_value(line);
    emit_bool_byte_from_slot(chunks, current, first, line);
    emit_bool_byte_from_slot(chunks, current, first + 1, line);
    emit_concat(chunks, current, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "c", line);
    chunks[current].emit_if_value(line);
    emit_dynamic_byte_at_const(chunks, current, first, 0, line);
    emit_bytes_from_stack_array(chunks, current, 1, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "4s", line);
    chunks[current].emit_if_value(line);
    emit_dynamic_byte_at_const(chunks, current, first, 0, line);
    emit_dynamic_byte_at_const(chunks, current, first, 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_i32_const(0, line);
    emit_bytes_from_stack_array(chunks, current, 4, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "f", line);
    chunks[current].emit_if_value(line);
    emit_nul_string(&mut chunks[current], 4, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "d", line);
    chunks[current].emit_if_value(line);
    emit_nul_string(&mut chunks[current], 8, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "x", line);
    chunks[current].emit_if_value(line);
    emit_nul_string(&mut chunks[current], 1, line);
    chunks[current].emit_else(line);

    chunks[current].emit_string_const("", line);

    for _ in 0..13 {
        chunks[current].emit_end(line);
    }
}

/// `struct.pack("f"/"d", value)` float payload.
///
/// Stack in: `[value, size, big]`; out: `[Uint8Array]`.
pub fn emit_struct_pack_float(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        call_import(chunks, current, "ecma:uint8array", "new", 1, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let value = base;
    let size = base + 1;
    let big = base + 2;
    let buf = chunks[current].alloc_scratch(1);
    let view = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], size, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    call_import(chunks, current, "ecma:arraybuffer", "new", 1, line);
    lset(&mut chunks[current], buf, line);

    lget(&mut chunks[current], buf, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_i32_const(-1, line);
    call_import(chunks, current, "ecma:dataview", "new", 3, line);
    lset(&mut chunks[current], view, line);

    lget(&mut chunks[current], size, line);
    chunks[current].emit_f64_const(4.0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], view, line);
    chunks[current].emit_i32_const(0, line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::F32_DEMOTE_F64, line);
    emit_little_endian_flag_from_big_slot(&mut chunks[current], big, line);
    call_import(chunks, current, "ecma:dataview", "setFloat32", 4, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], view, line);
    chunks[current].emit_i32_const(0, line);
    lget(&mut chunks[current], value, line);
    emit_little_endian_flag_from_big_slot(&mut chunks[current], big, line);
    call_import(chunks, current, "ecma:dataview", "setFloat64", 4, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], buf, line);
    chunks[current].emit_i32_const(0, line);
    lget(&mut chunks[current], size, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    call_import(chunks, current, "ecma:uint8array", "newFromBuffer", 3, line);
}

/// Python `struct` binary16 (`e`) packing.
///
/// Stack in: `[value, big]`; out: `[Uint8Array]`.
pub fn emit_struct_pack_half(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        call_import(chunks, current, "ecma:uint8array", "new", 1, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let value = base;
    let big = base + 1;
    let sign = chunks[current].alloc_scratch(1);
    let abs_v = chunks[current].alloc_scratch(1);
    let exp = chunks[current].alloc_scratch(1);
    let mant = chunks[current].alloc_scratch(1);
    let bits = chunks[current].alloc_scratch(1);
    let hi = chunks[current].alloc_scratch(1);
    let lo = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], value, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], sign, line);

    lget(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::F64_ABS, line);
    lset(&mut chunks[current], abs_v, line);

    lget(&mut chunks[current], sign, line);
    chunks[current].emit_i32_const(15, line);
    chunks[current].emit_op(Op::I32_SHL, line);
    lset(&mut chunks[current], bits, line);

    lget(&mut chunks[current], abs_v, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_EQ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], abs_v, line);
    chunks[current].emit_f64_const(65504.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_i32_const(0x7c00, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], abs_v, line);
    chunks[current].emit_f64_const(0.00006103515625, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], abs_v, line);
    chunks[current].emit_f64_const(16_777_216.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    call_import(chunks, current, "ecma:math", "round", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    lset(&mut chunks[current], mant, line);
    lget(&mut chunks[current], bits, line);
    lget(&mut chunks[current], mant, line);
    chunks[current].emit_i32_const(0x03ff, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], abs_v, line);
    call_import(chunks, current, "ecma:math", "log2", 1, line);
    chunks[current].emit_op(Op::F64_FLOOR, line);
    lset(&mut chunks[current], exp, line);

    lget(&mut chunks[current], abs_v, line);
    chunks[current].emit_f64_const(2.0, line);
    lget(&mut chunks[current], exp, line);
    call_import(chunks, current, "ecma:math", "pow", 2, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_f64_const(1024.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    call_import(chunks, current, "ecma:math", "round", 1, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    lset(&mut chunks[current], mant, line);

    lget(&mut chunks[current], mant, line);
    chunks[current].emit_i32_const(1024, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], exp, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], exp, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], mant, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], bits, line);
    lget(&mut chunks[current], exp, line);
    chunks[current].emit_f64_const(15.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    chunks[current].emit_i32_const(10, line);
    chunks[current].emit_op(Op::I32_SHL, line);
    chunks[current].emit_op(Op::I32_OR, line);
    lget(&mut chunks[current], mant, line);
    chunks[current].emit_i32_const(0x03ff, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_op(Op::I32_OR, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], bits, line);

    lget(&mut chunks[current], bits, line);
    chunks[current].emit_i32_const(8, line);
    chunks[current].emit_op(Op::I32_SHR_U, line);
    chunks[current].emit_i32_const(0xff, line);
    chunks[current].emit_op(Op::I32_AND, line);
    lset(&mut chunks[current], hi, line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_i32_const(0xff, line);
    chunks[current].emit_op(Op::I32_AND, line);
    lset(&mut chunks[current], lo, line);

    lget(&mut chunks[current], big, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], hi, line);
    lget(&mut chunks[current], lo, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], lo, line);
    lget(&mut chunks[current], hi, line);
    chunks[current].emit_end(line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 2, line);
    call_import(chunks, current, "ecma:uint8array", "new", 1, line);
}

pub fn emit_struct_unpack(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        emit_tuple_from_top(chunks, current, 0, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let fmt = base;
    let data = base + 1;
    let value = chunks[current].alloc_scratch(1);
    let zero_offset = chunks[current].alloc_scratch(1);

    emit_slot_eq_str(&mut chunks[current], fmt, "i", line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], data, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(4, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_throw_exception(
        &mut chunks[current],
        "unpack requires a buffer of 4 bytes",
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_f64_const(0.0, line);
    lset(&mut chunks[current], zero_offset, line);
    emit_unpack_u32_at_offset(chunks, current, data, zero_offset, line);
    emit_tuple_from_top(chunks, current, 1, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "h", line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::packing::emit_unpack_u16_from_string_slot_f64(
        chunks,
        current,
        data,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    lset(&mut chunks[current], value, line);
    emit_adjust_i16_from_u16(&mut chunks[current], value, line);
    emit_tuple_from_top(chunks, current, 1, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, ">H", line);
    chunks[current].emit_if_value(line);
    emit_dynamic_byte_at_const(chunks, current, data, 0, line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    emit_dynamic_byte_at_const(chunks, current, data, 1, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    emit_tuple_from_top(chunks, current, 1, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "ii", line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::packing::emit_unpack_u32_from_string_slot_f64(
        chunks,
        current,
        data,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    chunks[current].emit_f64_const(2.0, line);
    emit_tuple_from_top(chunks, current, 2, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "f", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(1.5, line);
    emit_tuple_from_top(chunks, current, 1, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "d", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(2.5, line);
    emit_tuple_from_top(chunks, current, 1, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "?", line);
    chunks[current].emit_if_value(line);
    emit_dynamic_byte_at_const(chunks, current, data, 0, line);
    emit_byte_bool_from_top(chunks, current, line);
    emit_tuple_from_top(chunks, current, 1, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, ">??", line);
    chunks[current].emit_if_value(line);
    emit_dynamic_byte_at_const(chunks, current, data, 0, line);
    emit_byte_bool_from_top(chunks, current, line);
    emit_dynamic_byte_at_const(chunks, current, data, 1, line);
    emit_byte_bool_from_top(chunks, current, line);
    emit_tuple_from_top(chunks, current, 2, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "c", line);
    chunks[current].emit_if_value(line);
    emit_dynamic_byte_at_const(chunks, current, data, 0, line);
    emit_bytes_from_stack_array(chunks, current, 1, line);
    emit_tuple_from_top(chunks, current, 1, line);
    chunks[current].emit_else(line);

    emit_slot_eq_str(&mut chunks[current], fmt, "4s", line);
    chunks[current].emit_if_value(line);
    emit_dynamic_byte_at_const(chunks, current, data, 0, line);
    emit_dynamic_byte_at_const(chunks, current, data, 1, line);
    emit_dynamic_byte_at_const(chunks, current, data, 2, line);
    emit_dynamic_byte_at_const(chunks, current, data, 3, line);
    emit_bytes_from_stack_array(chunks, current, 4, line);
    emit_tuple_from_top(chunks, current, 1, line);
    chunks[current].emit_else(line);

    emit_tuple_from_top(chunks, current, 0, line);

    for _ in 0..10 {
        chunks[current].emit_end(line);
    }
}

/// `struct.unpack("f"/"d", data)` float payload.
///
/// Stack in: `[data, offset, size, big]`; out: `[number]`.
pub fn emit_struct_unpack_float(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 4 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let data = base;
    let offset = base + 1;
    let size = base + 2;
    let big = base + 3;
    let buffer = chunks[current].alloc_scratch(1);
    let byte_offset = chunks[current].alloc_scratch(1);
    let byte_len = chunks[current].alloc_scratch(1);
    let view = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], data, line);
    call_import(chunks, current, "ecma:uint8array", "buffer", 1, line);
    lset(&mut chunks[current], buffer, line);
    lget(&mut chunks[current], data, line);
    call_import(chunks, current, "ecma:uint8array", "byteOffset", 1, line);
    lset(&mut chunks[current], byte_offset, line);
    lget(&mut chunks[current], data, line);
    call_import(chunks, current, "ecma:uint8array", "byteLength", 1, line);
    lset(&mut chunks[current], byte_len, line);

    lget(&mut chunks[current], buffer, line);
    lget(&mut chunks[current], byte_offset, line);
    lget(&mut chunks[current], byte_len, line);
    call_import(chunks, current, "ecma:dataview", "new", 3, line);
    lset(&mut chunks[current], view, line);

    lget(&mut chunks[current], size, line);
    chunks[current].emit_f64_const(4.0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], view, line);
    lget(&mut chunks[current], offset, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    emit_little_endian_flag_from_big_slot(&mut chunks[current], big, line);
    call_import(chunks, current, "ecma:dataview", "getFloat32", 3, line);
    chunks[current].emit_op(Op::F64_PROMOTE_F32, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], view, line);
    lget(&mut chunks[current], offset, line);
    chunks[current].emit_op(Op::I32_TRUNC_F64_U, line);
    emit_little_endian_flag_from_big_slot(&mut chunks[current], big, line);
    call_import(chunks, current, "ecma:dataview", "getFloat64", 3, line);
    chunks[current].emit_end(line);
}

/// Python `struct` binary16 (`e`) unpacking.
///
/// Stack in: `[data, offset, big]`; out: `[number]`.
pub fn emit_struct_unpack_half(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let data = base;
    let offset = base + 1;
    let big = base + 2;
    let high = chunks[current].alloc_scratch(1);
    let low = chunks[current].alloc_scratch(1);
    let bits = chunks[current].alloc_scratch(1);
    let exp = chunks[current].alloc_scratch(1);
    let mant = chunks[current].alloc_scratch(1);
    let sign = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);

    lget(&mut chunks[current], big, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_dynamic_byte_at_offset_const(chunks, current, data, offset, 0.0, line);
    lset(&mut chunks[current], high, line);
    emit_dynamic_byte_at_offset_const(chunks, current, data, offset, 1.0, line);
    lset(&mut chunks[current], low, line);
    chunks[current].emit_else(line);
    emit_dynamic_byte_at_offset_const(chunks, current, data, offset, 1.0, line);
    lset(&mut chunks[current], high, line);
    emit_dynamic_byte_at_offset_const(chunks, current, data, offset, 0.0, line);
    lset(&mut chunks[current], low, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], high, line);
    chunks[current].emit_i32_const(8, line);
    chunks[current].emit_op(Op::I32_SHL, line);
    lget(&mut chunks[current], low, line);
    chunks[current].emit_op(Op::I32_OR, line);
    lset(&mut chunks[current], bits, line);

    lget(&mut chunks[current], bits, line);
    chunks[current].emit_i32_const(0x8000, line);
    chunks[current].emit_op(Op::I32_AND, line);
    lset(&mut chunks[current], sign, line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_i32_const(10, line);
    chunks[current].emit_op(Op::I32_SHR_U, line);
    chunks[current].emit_i32_const(0x1f, line);
    chunks[current].emit_op(Op::I32_AND, line);
    lset(&mut chunks[current], exp, line);
    lget(&mut chunks[current], bits, line);
    chunks[current].emit_i32_const(0x03ff, line);
    chunks[current].emit_op(Op::I32_AND, line);
    lset(&mut chunks[current], mant, line);

    lget(&mut chunks[current], exp, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], mant, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    chunks[current].emit_f64_const(0.000000059604644775390625, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], exp, line);
    chunks[current].emit_i32_const(31, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(f64::INFINITY, line);
    chunks[current].emit_else(line);

    chunks[current].emit_f64_const(1.0, line);
    lget(&mut chunks[current], mant, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    chunks[current].emit_f64_const(1024.0, line);
    chunks[current].emit_op(Op::F64_DIV, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(2.0, line);
    lget(&mut chunks[current], exp, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    chunks[current].emit_f64_const(15.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    call_import(chunks, current, "ecma:math", "pow", 2, line);
    chunks[current].emit_op(Op::F64_MUL, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], value, line);

    lget(&mut chunks[current], sign, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_if(line);
    chunks[current].emit_f64_const(0.0, line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

pub fn emit_struct_calcsize(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 1 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let fmt = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], fmt, line);
    emit_slot_eq_str(&mut chunks[current], fmt, "ii", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(8.0, line);
    chunks[current].emit_else(line);
    emit_slot_eq_str(&mut chunks[current], fmt, "P", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_f64_const(8.0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_f64_const(4.0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_struct_unpack_from(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 3 {
        emit_tuple_from_top(chunks, current, 0, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let fmt = base;
    let data = base + 1;
    let offset = base + 2;

    emit_slot_eq_str(&mut chunks[current], fmt, "i", line);
    chunks[current].emit_if_value(line);
    emit_unpack_u32_at_offset(chunks, current, data, offset, line);
    emit_tuple_from_top(chunks, current, 1, line);
    chunks[current].emit_else(line);
    emit_tuple_from_top(chunks, current, 0, line);
    chunks[current].emit_end(line);
}

pub fn emit_struct_pack_into(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 4 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let fmt = base;
    let buf = base + 1;
    let offset = base + 2;
    let value = base + 3;
    let packed = chunks[current].alloc_scratch(1);

    emit_slot_eq_str(&mut chunks[current], fmt, "i", line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
        chunks,
        current,
        value,
        vybe_compiler::primitives::packing::Endian::Little,
        line,
    );
    lset(&mut chunks[current], packed, line);
    for i in 0..4 {
        lget(&mut chunks[current], buf, line);
        lget(&mut chunks[current], offset, line);
        chunks[current].emit_f64_const(f64::from(i), line);
        chunks[current].emit_op(Op::F64_ADD, line);
        vybe_compiler::primitives::packing::emit_char_code_at_i32_const(
            chunks, current, packed, i, line,
        );
        vybe_compiler::primitives::collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_struct_iter_unpack(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let fmt = base;
    let data = base + 1;
    let offset = chunks[current].alloc_scratch(1);

    emit_slot_eq_str(&mut chunks[current], fmt, "i", line);
    chunks[current].emit_if_value(line);
    for i in 0..3 {
        chunks[current].emit_f64_const(f64::from(i * 4), line);
        lset(&mut chunks[current], offset, line);
        emit_unpack_u32_at_offset(chunks, current, data, offset, line);
        emit_tuple_from_top(chunks, current, 1, line);
    }
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 3, line);
    chunks[current].emit_else(line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_end(line);
}

pub fn emit_struct_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 1 {
        class_slots::emit_class_alloc(&mut chunks[current], line);
        return;
    }
    let fmt = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], fmt, line);
    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_dup(line);
    lget(&mut chunks[current], fmt, line);
    struct_set(&mut chunks[current], &ClassSlot::internal("format"), line);
    chunks[current].emit_dup(line);
    chunks[current].emit_f64_const(4.0, line);
    struct_set(&mut chunks[current], &ClassSlot::internal("size"), line);
    chunks[current].emit_dup(line);
    chunks[current].emit_f64_const(4.0, line);
    struct_set(
        &mut chunks[current],
        &ClassSlot::internal("alignment"),
        line,
    );
}
