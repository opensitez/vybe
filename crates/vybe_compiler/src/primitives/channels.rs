use std::sync::Arc;

use vybe_runtime::Chunk;
use vybe_runtime::Value;
use vybe_runtime::opcode::Op;

use crate::primitives::collections;

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn struct_key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(name)))
}

fn emit_ref_is_object_like(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);

    for (module, func) in [
        ("wasm:js-undefined", "test"),
        ("wasm:js-number", "test"),
        ("wasm:js-string", "test"),
        ("wasm:js-boolean", "test"),
        ("wasm:js-bigint", "test"),
    ] {
        lget(chunk, slot, line);
        let idx = chunk.add_import(module, func);
        chunk.emit_call(idx, 1, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_op(Op::I32_AND, line);
    }
}

fn emit_autoderef_cell(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj_slot = alloc_local(&mut chunks[current]);
    let result_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], obj_slot, line);

    lget(&mut chunks[current], obj_slot, line);
    lset(&mut chunks[current], result_slot, line);

    lget(&mut chunks[current], obj_slot, line);
    emit_ref_is_object_like(&mut chunks[current], obj_slot, line);
    chunks[current].emit_if(line);

    lget(&mut chunks[current], obj_slot, line);
    let kind_key = struct_key(&mut chunks[current], "__ref_kind");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, kind_key, line);
    chunks[current].emit_string_const("cell", line);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);

    lget(&mut chunks[current], obj_slot, line);
    let value_key = struct_key(&mut chunks[current], "__value");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, value_key, line);
    lset(&mut chunks[current], result_slot, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], result_slot, line);
}

pub fn emit_send(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = alloc_local(&mut chunks[current]);
    let channel_slot = alloc_local(&mut chunks[current]);
    let queue_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], value_slot, line);
    lset(&mut chunks[current], channel_slot, line);

    lget(&mut chunks[current], channel_slot, line);
    emit_autoderef_cell(chunks, current, line);
    let queue_key = struct_key(&mut chunks[current], "queue");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, queue_key, line);
    emit_autoderef_cell(chunks, current, line);
    lset(&mut chunks[current], queue_slot, line);

    lget(&mut chunks[current], queue_slot, line);
    lget(&mut chunks[current], value_slot, line);
    collections::emit_push(chunks, current, line);
}

pub fn emit_receive(chunks: &mut [Chunk], current: usize, line: u32) {
    let channel_slot = alloc_local(&mut chunks[current]);
    let queue_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], channel_slot, line);

    lget(&mut chunks[current], channel_slot, line);
    emit_autoderef_cell(chunks, current, line);
    let queue_key = struct_key(&mut chunks[current], "queue");
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, queue_key, line);
    emit_autoderef_cell(chunks, current, line);
    lset(&mut chunks[current], queue_slot, line);

    lget(&mut chunks[current], queue_slot, line);
    collections::emit_shift(chunks, current, line);
}

pub fn emit_len(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_autoderef_cell(chunks, current, line);
    let queue_key = chunks[current].add_constant(Value::String(Arc::from("queue")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, queue_key, line);
    emit_autoderef_cell(chunks, current, line);
    collections::emit_len(chunks, current, line);
}

pub fn emit_cap(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_autoderef_cell(chunks, current, line);
    let cap_key = chunks[current].add_constant(Value::String(Arc::from("capacity")));
    chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, cap_key, line);
}

pub fn emit_close(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_autoderef_cell(chunks, current, line);
    let closed_key = chunks[current].add_constant(Value::String(Arc::from("closed")));
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_struct_field_op(Op::STRUCT_SET, 0, closed_key, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

// ── AST lowering ────────────────────────────────────────────────────────────
//
// Language-agnostic helpers that build the canonical channel AST shape a
// walker splices in; the `emit_*` fns above lower that shape to bytecode.
// Kept here (not in a per-language folder) so any language with channels can
// reuse both halves.
