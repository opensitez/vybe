//! JVM `java.lang.Class` / `Object.getClass()` reflection, backed by the
//! shared reflection descriptor shape.
//!
//! Moved here from `languages/java` unchanged: `Class` is a JDK type, so any
//! JVM language that reaches it should reach the same emitter. The language
//! crate keeps only the profile rows that NAME these — `getName`,
//! `getSimpleName` — because those spellings are how Java asks for it.

use vybe_compiler::primitives::instructions::host;
use vybe_compiler::primitives::{reflection, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `Field.set(obj, name, value)` — reflective field write, answering the
/// written value. Stack: `[obj, name, value] -> [value]`.
pub fn emit_field_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let field = chunks[current].alloc_scratch(1);
    let object = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, field, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, object, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
}

/// Reflective numeric field increment: `obj.<name> += delta`, answering the
/// new value. Stack: `[obj, name, delta] -> [new value]`.
pub fn emit_field_inc(chunks: &mut [Chunk], current: usize, line: u32) {
    let delta = chunks[current].alloc_scratch(1);
    let field = chunks[current].alloc_scratch(1);
    let object = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, delta, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, field, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, object, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
    host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delta, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, object, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    host::emit(&mut chunks[current], "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
}

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
