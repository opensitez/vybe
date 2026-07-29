//! Dart reflection/runtimeType adapters backed by the shared reflection shape.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_compiler::primitives::instructions::{core_wasm, host};
use vybe_compiler::primitives::{loops, reflection};

const SET_MARKER_KEY: &str = "__dart_set_marker";
const MAP_ORDER_KEY: &str = "__dart_map_order";

fn reserve_slot(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn get_field(chunk: &mut Chunk, name: &str, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
}

fn set_field(chunk: &mut Chunk, name: &str, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(name)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn set_string_field(chunk: &mut Chunk, name: &str, value: &str, line: u32) {
    core_wasm::dup(chunk, line);
    chunk.emit_string_const(value, line);
    set_field(chunk, name, line);
}

fn emit_type_descriptor(chunk: &mut Chunk, name: &str, kind: reflection::ReflectKind, line: u32) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    set_string_field(chunk, reflection::FIELD_TYPE, name, line);
    set_string_field(chunk, reflection::FIELD_TYPE_NAME, name, line);
    set_string_field(chunk, reflection::FIELD_KIND, kind.as_str(), line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const(name, line);
    chunk.emit_op_u16(Op::ARRAY_NEW_FIXED, 1, line);
    set_field(chunk, reflection::FIELD_TYPES, line);
}

fn emit_type_from_slot(
    chunk: &mut Chunk,
    value_slot: u16,
    name: &str,
    kind: reflection::ReflectKind,
    line: u32,
) {
    let type_name_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    get_field(chunk, reflection::FIELD_TYPE, line);
    chunk.emit_op_u16(Op::LOCAL_SET, type_name_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, type_name_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    get_field(chunk, reflection::FIELD_TYPE_NAME, line);
    chunk.emit_op_u16(Op::LOCAL_SET, type_name_slot, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, type_name_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    emit_type_descriptor(chunk, name, kind, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, type_name_slot, line);
    emit_type_descriptor_from_name_on_stack(chunk, kind, line);
    chunk.emit_end(line);
}

fn emit_type_descriptor_from_name_on_stack(
    chunk: &mut Chunk,
    kind: reflection::ReflectKind,
    line: u32,
) {
    let name_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    let descriptor_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, descriptor_slot, line);
    reflection::emit_set_slot_field_from_local(
        chunk,
        descriptor_slot,
        reflection::FIELD_TYPE,
        name_slot,
        line,
    );
    reflection::emit_set_slot_field_from_local(
        chunk,
        descriptor_slot,
        reflection::FIELD_TYPE_NAME,
        name_slot,
        line,
    );
    reflection::emit_stamp_kind(chunk, descriptor_slot, kind, line);
    chunk.emit_op_u16(Op::LOCAL_GET, descriptor_slot, line);
}

fn emit_slot_truthy_field(chunk: &mut Chunk, value_slot: u16, name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    get_field(chunk, name, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

/// Dart `value.runtimeType`.
///
/// Stack: `[value] -> [Type-like shared reflection descriptor]`.
pub fn emit_dart_runtime_type(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    emit_type_descriptor(chunk, "Null", reflection::ReflectKind::Null, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    reflection::emit_typeof_in_chunk(chunk, line);
    let tag_slot = reserve_slot(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, tag_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, tag_slot, line);
    chunk.emit_string_const("number", line);
    chunk.emit_op(Op::EQ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "ecma:number", "isInteger", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    emit_type_descriptor(chunk, "int", reflection::ReflectKind::Number, line);
    chunk.emit_else(line);
    emit_type_descriptor(chunk, "double", reflection::ReflectKind::Number, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    for (tag, name, kind) in [
        ("string", "String", reflection::ReflectKind::String),
        ("boolean", "bool", reflection::ReflectKind::Bool),
        ("function", "Function", reflection::ReflectKind::Function),
    ] {
        chunk.emit_op_u16(Op::LOCAL_GET, tag_slot, line);
        chunk.emit_string_const(tag, line);
        chunk.emit_op(Op::EQ, line);
        chunk.emit_if(line);
        emit_type_descriptor(chunk, name, kind, line);
        chunk.emit_else(line);
    }

    emit_slot_truthy_field(
        chunk,
        value_slot,
        vybe_compiler::primitives::tuples::TUPLE_TAG,
        line,
    );
    chunk.emit_if(line);
    emit_type_descriptor(chunk, "Record", reflection::ReflectKind::Struct, line);
    chunk.emit_else(line);

    emit_slot_truthy_field(chunk, value_slot, SET_MARKER_KEY, line);
    chunk.emit_if(line);
    emit_type_descriptor(chunk, "Set", reflection::ReflectKind::Set, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(chunk, "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    emit_type_descriptor(chunk, "List", reflection::ReflectKind::Array, line);
    chunk.emit_else(line);

    emit_slot_truthy_field(chunk, value_slot, MAP_ORDER_KEY, line);
    chunk.emit_if(line);
    emit_type_descriptor(chunk, "Map", reflection::ReflectKind::Map, line);
    chunk.emit_else(line);

    emit_type_from_slot(chunk, value_slot, "Map", reflection::ReflectKind::Map, line);

    for _ in 0..9 {
        chunk.emit_end(line);
    }
}

/// Dart `Type.toString()` on the descriptor produced by `runtimeType`.
///
/// Stack: `[type_descriptor] -> [type_name_string]`.
pub fn emit_dart_type_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    get_field(chunk, reflection::FIELD_TYPE_NAME, line);
}

/// Dart `value is List<int>`.
///
/// Stack: `[value] -> [bool]`.
pub fn emit_dart_is_list_of_int(chunks: &mut [Chunk], current: usize, line: u32) {
    let value_slot = reserve_slot(&mut chunks[current]);
    let result_slot = reserve_slot(&mut chunks[current]);
    let elem_slot = reserve_slot(&mut chunks[current]);
    let idx_slot = reserve_slot(&mut chunks[current]);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    host::emit(&mut chunks[current], "ecma:array", "isArray", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    let state = loops::emit_for_in_start(chunks, current, value_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    reflection::emit_typeof_in_chunk(&mut chunks[current], line);
    chunks[current].emit_string_const("number", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    host::emit(&mut chunks[current], "ecma:number", "isInteger", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}
