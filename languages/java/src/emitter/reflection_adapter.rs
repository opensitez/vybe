//! Java reflection surface backed by the shared reflection descriptor shape.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::{reflection, strings};

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_is_string(chunk: &mut Chunk, slot: u16, line: u32) {
    get(chunk, slot, line);
    reflection::emit_typeof_in_chunk(chunk, line);
    chunk.emit_string_const("string", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

/// Java `Class.getName()`.
///
/// Stack: `[class_token] -> [class_name]`. The token may be the historical
/// Java string token or a shared reflection descriptor carrying `__typename`.
pub fn emit_class_name(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let token = chunk.alloc_scratch(1);
    set(chunk, token, line);

    emit_is_string(chunk, token, line);
    chunk.emit_if_value(line);
    get(chunk, token, line);
    chunk.emit_else(line);
    get(chunk, token, line);
    reflection::emit_descriptor_field(chunk, reflection::FIELD_TYPE_NAME, line);
    let name = chunk.alloc_scratch(1);
    set(chunk, name, line);
    get(chunk, name, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    get(chunk, token, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_else(line);
    get(chunk, name, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Java `Class.getSimpleName()`.
///
/// Stack: `[class_token] -> [simple_name]`.
pub fn emit_class_simple_name(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_class_name(chunks, current, line);

    let chunk = &mut chunks[current];
    let name = chunk.alloc_scratch(1);
    let dot = chunk.alloc_scratch(1);
    let dollar = chunk.alloc_scratch(1);
    let start = chunk.alloc_scratch(1);

    set(chunk, name, line);

    get(chunk, name, line);
    chunk.emit_string_const(".", line);
    host::emit(chunk, "ecma:string", "lastIndexOf", 2, line);
    set(chunk, dot, line);

    get(chunk, name, line);
    chunk.emit_string_const("$", line);
    host::emit(chunk, "ecma:string", "lastIndexOf", 2, line);
    set(chunk, dollar, line);

    get(chunk, dollar, line);
    get(chunk, dot, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    get(chunk, dollar, line);
    chunk.emit_else(line);
    get(chunk, dot, line);
    chunk.emit_end(line);
    chunk.emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    set(chunk, start, line);

    get(chunk, name, line);
    get(chunk, start, line);
    get(chunk, name, line);
    strings::emit_length(chunk, line);
    host::emit(chunk, "ecma:string", "substring", 3, line);
}

/// Java `Object.getClass()`.
///
/// Stack: `[value] -> [class_token]`. Object instances use the shared
/// reflection stamps (`__typename`, then `__type`); primitive wrapper cases
/// expose Java's boxed class names.
pub fn emit_object_get_class(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let tag = chunk.alloc_scratch(1);
    let name = chunk.alloc_scratch(1);
    set(chunk, value, line);

    get(chunk, value, line);
    reflection::emit_typeof_in_chunk(chunk, line);
    set(chunk, tag, line);

    get(chunk, tag, line);
    chunk.emit_string_const("string", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("String", line);
    chunk.emit_else(line);

    get(chunk, tag, line);
    chunk.emit_string_const("number", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("Integer", line);
    chunk.emit_else(line);

    get(chunk, tag, line);
    chunk.emit_string_const("boolean", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("Boolean", line);
    chunk.emit_else(line);

    get(chunk, value, line);
    reflection::emit_descriptor_field(chunk, reflection::FIELD_TYPE, line);
    set(chunk, name, line);
    get(chunk, name, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    get(chunk, value, line);
    reflection::emit_descriptor_field(chunk, reflection::FIELD_TYPE_NAME, line);
    set(chunk, name, line);
    get(chunk, name, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("Object", line);
    chunk.emit_else(line);
    get(chunk, name, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    get(chunk, name, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
