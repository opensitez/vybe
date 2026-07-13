//! Java `Optional` adapters.
//!
//! The shared host has no `ecma:optional` module, so Java lowers Optional to a
//! tiny pair array: `[present: bool, value]`.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_emitter::collections;

pub fn emit_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op(Op::NULL, line);
    collections::emit_array_new(chunks, current, 2, line);
}

pub fn emit_of(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_array_new(chunks, current, 2, line);
}

pub fn emit_of_nullable(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::NULL, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_empty(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_of(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_is_present(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
}

pub fn emit_or_else(chunks: &mut [Chunk], current: usize, call_supplier: bool, line: u32) {
    let fallback_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    let result_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fallback_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fallback_slot, line);
    if call_supplier {
        chunks[current].emit_op_u8(Op::CALL_REF, 0, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_if_present(chunks: &mut [Chunk], current: usize, line: u32) {
    let consumer_slot = chunks[current].alloc_scratch(1);
    let optional_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, consumer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, optional_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    emit_is_present(chunks, current, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, consumer_slot, line);
    emit_value_from_slot(chunks, current, optional_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_value_from_slot(chunks: &mut [Chunk], current: usize, optional_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, optional_slot, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
}
