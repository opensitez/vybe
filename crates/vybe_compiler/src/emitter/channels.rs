use std::sync::Arc;

use vybe_bytecode::Chunk;
use vybe_bytecode::Value;
use vybe_bytecode::opcode::Op;

use crate::emitter::collections;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let slot = chunk.local_count;
    chunk.local_count += 1;
    slot
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}

fn struct_key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(name)))
}

fn emit_autoderef_cell(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], obj_slot, line);

    let done_block = chunks[current].emit_block(line);
    let fallback_block = chunks[current].emit_block(line);

    lget(&mut chunks[current], obj_slot, line);
    chunks[current].emit_op(Op::REF_IS_OBJECT, line);
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(0, line);

    lget(&mut chunks[current], obj_slot, line);
    let kind_key = struct_key(&mut chunks[current], "__ref_kind");
    chunks[current].emit_op_u16(Op::STRUCT_GET, kind_key, line);
    let cell_value = chunks[current].add_constant(Value::String(Arc::from("cell")));
    chunks[current].emit_op_u16(Op::CONST, cell_value, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(0, line);

    lget(&mut chunks[current], obj_slot, line);
    let value_key = struct_key(&mut chunks[current], "__value");
    chunks[current].emit_op_u16(Op::STRUCT_GET, value_key, line);
    chunks[current].emit_br(1, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(fallback_block);
    lget(&mut chunks[current], obj_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done_block);
}

pub fn emit_send(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = alloc_local(&mut chunks[current]);
    let queue_cell_slot = alloc_local(&mut chunks[current]);
    let queue_slot = alloc_local(&mut chunks[current]);
    let next_queue_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], value_slot, line);
    lset(&mut chunks[current], queue_cell_slot, line);

    lget(&mut chunks[current], queue_cell_slot, line);
    emit_autoderef_cell(chunks, current, line);
    lset(&mut chunks[current], queue_slot, line);

    lget(&mut chunks[current], queue_slot, line);
    lget(&mut chunks[current], value_slot, line);
    collections::emit_array_new(chunks, current, 1, line);
    collections::emit_concat(chunks, current, line);
    lset(&mut chunks[current], next_queue_slot, line);

    lget(&mut chunks[current], queue_cell_slot, line);
    lget(&mut chunks[current], next_queue_slot, line);
    let value_key = struct_key(&mut chunks[current], "__value");
    chunks[current].emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], next_queue_slot, line);
}

pub fn emit_receive(chunks: &mut [Chunk], current: usize, line: u32) {
    let channel_slot = alloc_local(&mut chunks[current]);
    let queue_slot = alloc_local(&mut chunks[current]);
    let value_slot = alloc_local(&mut chunks[current]);
    let next_queue_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], channel_slot, line);

    lget(&mut chunks[current], channel_slot, line);
    emit_autoderef_cell(chunks, current, line);
    let queue_key = struct_key(&mut chunks[current], "queue");
    chunks[current].emit_op_u16(Op::STRUCT_GET, queue_key, line);
    emit_autoderef_cell(chunks, current, line);
    lset(&mut chunks[current], queue_slot, line);

    lget(&mut chunks[current], queue_slot, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], value_slot, line);

    lget(&mut chunks[current], queue_slot, line);
    chunks[current].emit_op(Op::I32_CONST_1, line);
    let max_index = chunks[current].add_constant(Value::I32(i32::MAX));
    chunks[current].emit_op_u16(Op::CONST, max_index, line);
    collections::emit_slice(chunks, current, line);
    lset(&mut chunks[current], next_queue_slot, line);

    lget(&mut chunks[current], channel_slot, line);
    emit_autoderef_cell(chunks, current, line);
    chunks[current].emit_op_u16(Op::STRUCT_GET, queue_key, line);
    lget(&mut chunks[current], next_queue_slot, line);
    let value_key = struct_key(&mut chunks[current], "__value");
    chunks[current].emit_op_u16(Op::STRUCT_SET, value_key, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], value_slot, line);
}

pub fn emit_len(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_autoderef_cell(chunks, current, line);
    let queue_key = chunks[current].add_constant(Value::String(Arc::from("queue")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, queue_key, line);
    emit_autoderef_cell(chunks, current, line);
    collections::emit_len(chunks, current, line);
}

pub fn emit_cap(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_autoderef_cell(chunks, current, line);
    let cap_key = chunks[current].add_constant(Value::String(Arc::from("capacity")));
    chunks[current].emit_op_u16(Op::STRUCT_GET, cap_key, line);
}

pub fn emit_close(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_autoderef_cell(chunks, current, line);
    let closed_key = chunks[current].add_constant(Value::String(Arc::from("closed")));
    chunks[current].emit_op(Op::TRUE, line);
    chunks[current].emit_op_u16(Op::STRUCT_SET, closed_key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}
