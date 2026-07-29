//! Go `encoding/binary` fixed-width byte-order helpers.
//!
//! Go owns the `encoding/binary` API surface; shared byte/endian mechanics
//! live in `vybe_compiler::primitives::packing`.

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
        "go.binary_put_uint16" => emit_put_uint16(chunks, current, argc, line),
        "go.binary_put_int16" => emit_put_int16(chunks, current, argc, line),
        "go.binary_uint16" => emit_uint16(chunks, current, argc, line),
        "go.binary_put_uint32" => emit_put_uint32(chunks, current, argc, line),
        "go.binary_uint32" => emit_uint32(chunks, current, argc, line),
        "go.binary_int32" => emit_int32(chunks, current, argc, line),
        "go.binary_put_uint64_parts" => emit_put_uint64_parts(chunks, current, argc, line),
        "go.binary_append_uint16" => emit_append_uint16(chunks, current, argc, line),
        "go.binary_append_uint32" => emit_append_uint32(chunks, current, argc, line),
        _ => return false,
    }
    true
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn to_number(chunk: &mut Chunk, line: u32) {
    let to_number = chunk.add_import("ecma:value", "toNumber");
    chunk.emit_call(to_number, 1, line);
}

fn emit_order_branch<F, G>(
    chunks: &mut [Chunk],
    current: usize,
    order_slot: u16,
    line: u32,
    mut little: F,
    mut big: G,
) where
    F: FnMut(&mut [Chunk], usize, u32),
    G: FnMut(&mut [Chunk], usize, u32),
{
    lget(&mut chunks[current], order_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    little(chunks, current, line);
    chunks[current].emit_else(line);
    big(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_adjust_i16(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) -> u16 {
    let adjusted = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], value_slot, line);
    to_number(&mut chunks[current], line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], value_slot, line);
    to_number(&mut chunks[current], line);
    chunks[current].emit_f64_const(65536.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_end(line);
    lset(&mut chunks[current], adjusted, line);
    adjusted
}

fn emit_empty_array(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    let slot = chunks[current].alloc_scratch(1);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], slot, line);
    slot
}

pub fn emit_put_uint16(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let order = base;
    let buf = base + 1;
    let value = base + 2;
    emit_order_branch(
        chunks,
        current,
        order,
        line,
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u16_to_array_from_number_slot(
                chunks,
                current,
                buf,
                value,
                vybe_compiler::primitives::packing::Endian::Little,
                line,
            );
        },
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u16_to_array_from_number_slot(
                chunks,
                current,
                buf,
                value,
                vybe_compiler::primitives::packing::Endian::Big,
                line,
            );
        },
    );
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_put_int16(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let adjusted = emit_adjust_i16(chunks, current, base + 2, line);
    let order = base;
    let buf = base + 1;
    emit_order_branch(
        chunks,
        current,
        order,
        line,
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u16_to_array_from_number_slot(
                chunks,
                current,
                buf,
                adjusted,
                vybe_compiler::primitives::packing::Endian::Little,
                line,
            );
        },
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u16_to_array_from_number_slot(
                chunks,
                current,
                buf,
                adjusted,
                vybe_compiler::primitives::packing::Endian::Big,
                line,
            );
        },
    );
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_uint16(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    emit_order_branch(
        chunks,
        current,
        base,
        line,
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_load_u16_from_array_f64(
                chunks,
                current,
                base + 1,
                vybe_compiler::primitives::packing::Endian::Little,
                line,
            );
        },
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_load_u16_from_array_f64(
                chunks,
                current,
                base + 1,
                vybe_compiler::primitives::packing::Endian::Big,
                line,
            );
        },
    );
}

pub fn emit_put_uint32(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let order = base;
    let buf = base + 1;
    let value = base + 2;
    emit_order_branch(
        chunks,
        current,
        order,
        line,
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u32_to_array_from_number_slot(
                chunks,
                current,
                buf,
                value,
                vybe_compiler::primitives::packing::Endian::Little,
                line,
            );
        },
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u32_to_array_from_number_slot(
                chunks,
                current,
                buf,
                value,
                vybe_compiler::primitives::packing::Endian::Big,
                line,
            );
        },
    );
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_uint32(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    emit_order_branch(
        chunks,
        current,
        base,
        line,
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_load_u32_from_array_f64(
                chunks,
                current,
                base + 1,
                vybe_compiler::primitives::packing::Endian::Little,
                line,
            );
        },
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_load_u32_from_array_f64(
                chunks,
                current,
                base + 1,
                vybe_compiler::primitives::packing::Endian::Big,
                line,
            );
        },
    );
}

pub fn emit_int32(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_uint32(chunks, current, argc, line);
    let value = chunks[current].alloc_scratch(1);
    lset(&mut chunks[current], value, line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_f64_const(2147483648.0, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_f64_const(4294967296.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

pub fn emit_put_uint64_parts(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 4 {
        chunks[current].emit_op(Op::NULL, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let order = base;
    let buf = base + 1;
    let hi = base + 2;
    let lo = base + 3;
    emit_order_branch(
        chunks,
        current,
        order,
        line,
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u64_parts_to_array_from_number_slots(
                chunks,
                current,
                buf,
                hi,
                lo,
                vybe_compiler::primitives::packing::Endian::Little,
                line,
            );
        },
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u64_parts_to_array_from_number_slots(
                chunks,
                current,
                buf,
                hi,
                lo,
                vybe_compiler::primitives::packing::Endian::Big,
                line,
            );
        },
    );
    chunks[current].emit_op(Op::NULL, line);
}

pub fn emit_append_uint16(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let tmp = emit_empty_array(chunks, current, line);
    let order = base;
    let src = base + 1;
    let value = base + 2;
    emit_order_branch(
        chunks,
        current,
        order,
        line,
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u16_to_array_from_number_slot(
                chunks,
                current,
                tmp,
                value,
                vybe_compiler::primitives::packing::Endian::Little,
                line,
            );
        },
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u16_to_array_from_number_slot(
                chunks,
                current,
                tmp,
                value,
                vybe_compiler::primitives::packing::Endian::Big,
                line,
            );
        },
    );
    lget(&mut chunks[current], src, line);
    lget(&mut chunks[current], tmp, line);
    let concat = chunks[current].add_import("ecma:array", "concat");
    chunks[current].emit_call(concat, 2, line);
}

pub fn emit_append_uint32(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let tmp = emit_empty_array(chunks, current, line);
    let order = base;
    let src = base + 1;
    let value = base + 2;
    emit_order_branch(
        chunks,
        current,
        order,
        line,
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u32_to_array_from_number_slot(
                chunks,
                current,
                tmp,
                value,
                vybe_compiler::primitives::packing::Endian::Little,
                line,
            );
        },
        |chunks, current, line| {
            vybe_compiler::primitives::packing::emit_store_u32_to_array_from_number_slot(
                chunks,
                current,
                tmp,
                value,
                vybe_compiler::primitives::packing::Endian::Big,
                line,
            );
        },
    );
    lget(&mut chunks[current], src, line);
    lget(&mut chunks[current], tmp, line);
    let concat = chunks[current].add_import("ecma:array", "concat");
    chunks[current].emit_call(concat, 2, line);
}
