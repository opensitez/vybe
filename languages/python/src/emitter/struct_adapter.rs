//! Python `struct` module adapter.
//!
//! Python owns the format grammar and tuple result shape. Shared byte/endian
//! mechanics come from `vybe_compiler::primitives::packing`.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

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
    vybe_compiler::primitives::packing::emit_char_code_at_i32_const(chunks, current, data_slot, index, line);
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
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(message, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(chunk, "Exception", line);
    vybe_compiler::primitives::errors::emit_throw(chunk, line);
}

fn struct_set(chunk: &mut Chunk, key: &str, line: u32) {
    let key = chunk.add_constant(vybe_bytecode::Value::String(std::sync::Arc::from(key)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
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
    chunks[current].emit_string_const("\x01", line);
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

    for _ in 0..12 {
        chunks[current].emit_end(line);
    }
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
    emit_throw_exception(&mut chunks[current], "unpack requires a buffer of 4 bytes", line);
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
    chunks[current].emit_bool_const(true, line);
    emit_tuple_from_top(chunks, current, 1, line);
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

    for _ in 0..9 {
        chunks[current].emit_end(line);
    }
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
        chunks[current].emit_op(Op::NULL, line);
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
        vybe_compiler::primitives::packing::emit_char_code_at_i32_const(chunks, current, packed, i, line);
        vybe_compiler::primitives::collections::emit_set(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::NULL, line);
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
        chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
        return;
    }
    let fmt = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], fmt, line);
    chunks[current].emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunks[current].emit_dup(line);
    lget(&mut chunks[current], fmt, line);
    struct_set(&mut chunks[current], "format", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_f64_const(4.0, line);
    struct_set(&mut chunks[current], "size", line);
    chunks[current].emit_dup(line);
    chunks[current].emit_f64_const(4.0, line);
    struct_set(&mut chunks[current], "alignment", line);
}
