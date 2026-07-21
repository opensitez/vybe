//! Java class-literal helpers for ECMA-backed values.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_emitter::{
    collections,
    instructions::{core_wasm, host},
    reflection,
};

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn class_eq(chunks: &mut [Chunk], current: usize, class_slot: u16, name: &str, line: u32) {
    get(&mut chunks[current], class_slot, line);
    chunks[current].emit_string_const(name, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn typeof_eq(chunks: &mut [Chunk], current: usize, value_slot: u16, name: &str, line: u32) {
    get(&mut chunks[current], value_slot, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const(name, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

fn emit_is_array(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    get(&mut chunks[current], value_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_is_map(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    get(&mut chunks[current], value_slot, line);
    chunks[current].emit_string_const("size", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("number", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_is_string_builder(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    get(&mut chunks[current], value_slot, line);
    chunks[current].emit_string_const("__buffer", line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

fn emit_is_string_array(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    get(&mut chunks[current], value_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], value_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    host::emit(&mut chunks[current], "ecma:array", "get", 2, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn emit_class_branch(
    chunks: &mut [Chunk],
    current: usize,
    class_slot: u16,
    class_name: &str,
    line: u32,
) {
    class_eq(chunks, current, class_slot, class_name, line);
    chunks[current].emit_if(line);
}

fn normalize_class_token(chunks: &mut [Chunk], current: usize, class_slot: u16, line: u32) {
    get(&mut chunks[current], class_slot, line);
    reflection::emit_typeof_in_chunk(&mut chunks[current], line);
    chunks[current].emit_string_const("string", line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], class_slot, line);
    reflection::emit_descriptor_field(&mut chunks[current], reflection::FIELD_TYPE_NAME, line);
    set(&mut chunks[current], class_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_is_instance(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let class_name = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], class_name, line);
    normalize_class_token(chunks, current, class_name, line);

    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::NULL, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_class_branch(chunks, current, class_name, "Object", line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_class_branch(chunks, current, class_name, "Class", line);
    typeof_eq(chunks, current, value, "string", line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_class_branch(chunks, current, class_name, "String", line);
    typeof_eq(chunks, current, value, "string", line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_class_branch(chunks, current, class_name, "Comparable", line);
    typeof_eq(chunks, current, value, "string", line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_class_branch(chunks, current, class_name, "Serializable", line);
    typeof_eq(chunks, current, value, "string", line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    for name in [
        "Integer", "Long", "Double", "Float", "Number", "Short", "Byte",
    ] {
        emit_class_branch(chunks, current, class_name, name, line);
        typeof_eq(chunks, current, value, "number", line);
        vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
        chunks[current].emit_else(line);
    }

    emit_class_branch(chunks, current, class_name, "Boolean", line);
    typeof_eq(chunks, current, value, "boolean", line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_class_branch(chunks, current, class_name, "Character", line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    for name in [
        "int[]",
        "byte[]",
        "List",
        "Collection",
        "Vector",
        "Cloneable",
    ] {
        emit_class_branch(chunks, current, class_name, name, line);
        emit_is_array(chunks, current, value, line);
        chunks[current].emit_else(line);
    }

    emit_class_branch(chunks, current, class_name, "String[]", line);
    emit_is_string_array(chunks, current, value, line);
    chunks[current].emit_else(line);

    emit_class_branch(chunks, current, class_name, "StringBuilder", line);
    emit_is_string_builder(chunks, current, value, line);
    chunks[current].emit_else(line);

    for name in ["Set", "HashSet"] {
        emit_class_branch(chunks, current, class_name, name, line);
        emit_is_array(chunks, current, value, line);
        chunks[current].emit_else(line);
    }

    for name in ["Map", "HashMap"] {
        emit_class_branch(chunks, current, class_name, name, line);
        emit_is_map(chunks, current, value, line);
        chunks[current].emit_else(line);
    }

    emit_class_branch(chunks, current, class_name, "Throwable", line);
    get(&mut chunks[current], value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);

    // Close every branch above plus the initial null check.
    for _ in 0..28 {
        chunks[current].emit_end(line);
    }
}
