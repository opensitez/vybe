//! Go `reflect` facade routed through the shared reflection substrate.

use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};
use vybe_emitter::{collections, reflection};

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.reflect_typeof" => emit_typeof(chunks, current, argc, line),
        "go.reflect_valueof" => emit_valueof(chunks, current, argc, line),
        "go.reflect_kind" => emit_field(&mut chunks[current], reflection::FIELD_KIND, line),
        "go.reflect_name" => emit_field(&mut chunks[current], reflection::FIELD_TYPE_NAME, line),
        "go.reflect_interface" => emit_field(&mut chunks[current], reflection::FIELD_VALUE, line),
        "go.reflect_int" | "go.reflect_uint" | "go.reflect_float" | "go.reflect_bool"
        | "go.reflect_string" => emit_field(&mut chunks[current], reflection::FIELD_VALUE, line),
        "go.reflect_num_field" => {
            emit_field(&mut chunks[current], reflection::FIELD_FIELDS, line);
            collections::emit_len(chunks, current, line);
        }
        "go.reflect_field" => {
            let index = chunks[current].alloc_scratch(1);
            let recv = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
            emit_field(&mut chunks[current], reflection::FIELD_FIELDS, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
            collections::emit_get(chunks, current, line);
        }
        "go.reflect_field_by_name" => {
            let name = chunks[current].alloc_scratch(1);
            let recv = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
            chunks[current].emit_op(Op::NULL, line);
        }
        "go.reflect_len" => {
            emit_field(&mut chunks[current], reflection::FIELD_VALUE, line);
            collections::emit_len(chunks, current, line);
        }
        "go.reflect_index" => {
            let index = chunks[current].alloc_scratch(1);
            let recv = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
            emit_field(&mut chunks[current], reflection::FIELD_VALUE, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
            collections::emit_get(chunks, current, line);
            emit_wrap_existing_value(chunks, current, line);
        }
        "go.reflect_map_index" => {
            let key = chunks[current].alloc_scratch(1);
            let recv = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
            emit_field(&mut chunks[current], reflection::FIELD_VALUE, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
            emit_field(&mut chunks[current], reflection::FIELD_VALUE, line);
            collections::emit_get(chunks, current, line);
            emit_wrap_existing_value(chunks, current, line);
        }
        "go.reflect_is_valid" => emit_bool(&mut chunks[current], true, line),
        "go.reflect_is_nil" => {
            emit_field(&mut chunks[current], reflection::FIELD_VALUE, line);
            chunks[current].emit_op(Op::REF_IS_NULL, line);
            vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "go.reflect_can_set" => {
            emit_field(&mut chunks[current], reflection::FIELD_REF, line);
            chunks[current].emit_op(Op::REF_IS_NULL, line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
        }
        "go.reflect_is_zero" => emit_is_zero(chunks, current, line),
        "go.reflect_elem" => emit_elem(chunks, current, line),
        "go.reflect_set" => emit_set_value(chunks, current, line),
        "go.reflect_set_int"
        | "go.reflect_set_uint"
        | "go.reflect_set_string"
        | "go.reflect_set_bool" => emit_set_primitive(chunks, current, line),
        _ => return false,
    }
    true
}

fn emit_typeof(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let fields_slot = chunks[current].alloc_scratch(1);
    let kind_slot = chunks[current].alloc_scratch(1);
    let type_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, fields_slot, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, fields_slot, line);
    }
    if argc >= 2 {
        if argc >= 4 {
            chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        } else {
            chunks[current].emit_op(Op::NULL, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    let methods_slot = chunks[current].alloc_scratch(1);
    let attrs_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, methods_slot, line);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, attrs_slot, line);
    let out_slot = chunks[current].alloc_scratch(1);
    reflection::emit_type_descriptor(
        &mut chunks[current],
        out_slot,
        type_slot,
        reflection::ReflectKind::Object,
        fields_slot,
        methods_slot,
        attrs_slot,
        line,
    );
    stamp_kind_from_type_name(&mut chunks[current], out_slot, kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn emit_valueof(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let kind_slot = chunks[current].alloc_scratch(1);
    let type_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    if argc >= 2 {
        if argc >= 3 {
            chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        } else {
            chunks[current].emit_op(Op::NULL, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    } else {
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, kind_slot, line);
        chunks[current].emit_op(Op::NULL, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    let ref_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op(Op::NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ref_slot, line);
    let out_slot = chunks[current].alloc_scratch(1);
    reflection::emit_value_descriptor(
        &mut chunks[current],
        out_slot,
        value_slot,
        type_slot,
        reflection::ReflectKind::Object,
        ref_slot,
        line,
    );
    stamp_kind_from_type_name(&mut chunks[current], out_slot, kind_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn emit_wrap_existing_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    chunks[current].emit_string_const("any", line);
    chunks[current].emit_string_const("any", line);
    emit_valueof(chunks, current, 3, line);
}

fn emit_elem(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_field(&mut chunks[current], "__elem", line);
    let elem = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem, line);
    chunks[current].emit_end(line);
}

fn emit_set_value(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_field(&mut chunks[current], reflection::FIELD_VALUE, line);
    set_field_from_stack(&mut chunks[current], reflection::FIELD_VALUE, line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_set_primitive(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    set_field_from_stack(&mut chunks[current], reflection::FIELD_VALUE, line);
    chunks[current].emit_op(Op::NULL, line);
}

fn emit_is_zero(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_field(&mut chunks[current], reflection::FIELD_KIND, line);
    chunks[current].emit_string_const("string", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_field(&mut chunks[current], reflection::FIELD_VALUE, line);
    chunks[current].emit_string_const("", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    emit_field(&mut chunks[current], reflection::FIELD_VALUE, line);
    chunks[current].emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn stamp_kind_from_type_name(chunk: &mut Chunk, object_slot: u16, type_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, type_slot, line);
    let key = sconst(chunk, reflection::FIELD_KIND);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_field(chunk: &mut Chunk, field: &str, line: u32) {
    chunk.emit_string_const(field, line);
    let idx = chunk.add_import("ecma:reflect", "get");
    chunk.emit_call(idx, 2, line);
}

fn set_field_from_stack(chunk: &mut Chunk, field: &str, line: u32) {
    let key = sconst(chunk, field);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_bool(chunk: &mut Chunk, value: bool, line: u32) {
    chunk.emit_bool_const(value, line);
}

fn sconst(chunk: &mut Chunk, s: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(s)))
}
