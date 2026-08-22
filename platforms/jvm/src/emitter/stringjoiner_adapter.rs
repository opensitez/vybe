//! JVM `java.util.StringJoiner` adapter.
//!
//! A joiner is a struct holding its three fixed strings and the joined
//! buffer so far, plus a "has any element" flag — `toString` is `prefix +
//! buffer + suffix` unless nothing was added and an empty-value override
//! exists, in which case it is that override verbatim (JDK semantics).

use vybe_compiler::primitives::{instructions::host, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

const BUFFER: &str = "__sj_buffer";
const DELIM: &str = "__sj_delimiter";
const PREFIX: &str = "__sj_prefix";
const SUFFIX: &str = "__sj_suffix";
const HAS_ANY: &str = "__sj_has_any";
const EMPTY: &str = "__sj_empty_value";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn field_get(chunk: &mut Chunk, obj: u16, name: &str, line: u32) {
    get(chunk, obj, line);
    let k = chunk.add_constant(vybe_runtime::Value::String(name.into()));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

fn field_set_from_stack(chunk: &mut Chunk, obj: u16, name: &str, line: u32) {
    let value = chunk.alloc_scratch(1);
    set(chunk, value, line);
    get(chunk, obj, line);
    get(chunk, value, line);
    let k = chunk.add_constant(vybe_runtime::Value::String(name.into()));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

/// `new StringJoiner(delim)` / `new StringJoiner(delim, prefix, suffix)`.
pub fn emit_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let suffix = chunks[current].alloc_scratch(1);
    let prefix = chunks[current].alloc_scratch(1);
    let delim = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        set(&mut chunks[current], suffix, line);
        set(&mut chunks[current], prefix, line);
        for _ in 3..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        set(&mut chunks[current], delim, line);
    } else {
        if argc >= 1 {
            set(&mut chunks[current], delim, line);
            for _ in 1..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
        } else {
            chunks[current].emit_string_const("", line);
            set(&mut chunks[current], delim, line);
        }
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], prefix, line);
        chunks[current].emit_string_const("", line);
        set(&mut chunks[current], suffix, line);
    }
    let obj = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    set(&mut chunks[current], obj, line);
    get(&mut chunks[current], delim, line);
    field_set_from_stack(&mut chunks[current], obj, DELIM, line);
    get(&mut chunks[current], prefix, line);
    field_set_from_stack(&mut chunks[current], obj, PREFIX, line);
    get(&mut chunks[current], suffix, line);
    field_set_from_stack(&mut chunks[current], obj, SUFFIX, line);
    chunks[current].emit_string_const("", line);
    field_set_from_stack(&mut chunks[current], obj, BUFFER, line);
    chunks[current].emit_bool_const(false, line);
    field_set_from_stack(&mut chunks[current], obj, HAS_ANY, line);
    get(&mut chunks[current], obj, line);
}

/// `sj.add(text)` → the joiner (chainable).
pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], obj, HAS_ANY, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    field_get(&mut chunks[current], obj, BUFFER, line);
    field_get(&mut chunks[current], obj, DELIM, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    chunks[current].emit_else(line);
    field_get(&mut chunks[current], obj, BUFFER, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], value, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    field_set_from_stack(&mut chunks[current], obj, BUFFER, line);
    chunks[current].emit_bool_const(true, line);
    field_set_from_stack(&mut chunks[current], obj, HAS_ANY, line);
    get(&mut chunks[current], obj, line);
}

/// `sj.merge(other)` — other's BUFFER (its prefix/suffix do not travel),
/// added as ONE element; empty other is a no-op. → the joiner.
pub fn emit_merge(chunks: &mut [Chunk], current: usize, line: u32) {
    let other = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], other, line);
    set(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], other, HAS_ANY, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(&mut chunks[current], obj, line);
    field_get(&mut chunks[current], other, BUFFER, line);
    emit_add(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    get(&mut chunks[current], obj, line);
}

/// `sj.setEmptyValue(s)` → the joiner.
pub fn emit_set_empty_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], value, line);
    set(&mut chunks[current], obj, line);
    get(&mut chunks[current], value, line);
    field_set_from_stack(&mut chunks[current], obj, EMPTY, line);
    get(&mut chunks[current], obj, line);
}

/// The rendered string: empty-value override when nothing was added,
/// otherwise `prefix + buffer + suffix`. `[joiner] -> [string]`.
///
/// A value-producing if/else tree, NOT an early `RETURN`: these emitters
/// are inlined into whatever chunk is being compiled, so a return here
/// would return from the CALLER's function.
pub fn emit_to_string(chunks: &mut [Chunk], current: usize, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], obj, line);
    let rendered = chunks[current].alloc_scratch(1);
    field_get(&mut chunks[current], obj, PREFIX, line);
    field_get(&mut chunks[current], obj, BUFFER, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    field_get(&mut chunks[current], obj, SUFFIX, line);
    strings::emit_str_concat_coercing(&mut chunks[current], line);
    set(&mut chunks[current], rendered, line);

    field_get(&mut chunks[current], obj, HAS_ANY, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], rendered, line);
    chunks[current].emit_else(line);
    field_get(&mut chunks[current], obj, EMPTY, line);
    let empty = chunks[current].alloc_scratch(1);
    set(&mut chunks[current], empty, line);
    get(&mut chunks[current], empty, line);
    host::emit(&mut chunks[current], "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("string", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], empty, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], rendered, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `sj.length()` — the length of what `toString` would render.
pub fn emit_length(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_to_string(chunks, current, line);
    strings::emit_length(&mut chunks[current], line);
}
