//! Shared multi-value packet emission.
//!
//! This is the language-neutral runtime shape for call/return values that can
//! expand or truncate depending on context: Lua multiple returns, Go
//! multi-return, future Wasm multi-value bridges, and similar surfaces.
//! It is intentionally separate from `packing.rs`, which is binary/data layout.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

pub const MULTI_VALUE_TAG: &str = "__vybe_multi_value";

fn load(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn save(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Stamp the array-like value on top of stack as a multi-value packet.
/// Stack: `[array] -> [array]`.
pub fn emit_tag(chunks: &mut [Chunk], current: usize, line: u32) {
    let tag = chunks[current].add_constant(Value::String(Arc::from(MULTI_VALUE_TAG)));
    chunks[current].emit_dup(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::STRUCT_SET, tag, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// Pack the top `n` stack values into a shared multi-value packet.
/// Stack: `[v0, ..., vn] -> [packet]`.
pub fn emit_from_stack(chunks: &mut [Chunk], current: usize, n: u16, line: u32) {
    let base = chunks[current].alloc_scratch(n.max(1));
    crate::primitives::collections::emit_pack_n(chunks, current, n, base, line);
    emit_tag(chunks, current, line);
}

/// Push i32 truthiness for whether `slot` is a shared multi-value packet.
pub fn emit_is_multi_value_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    load(chunk, slot, line);
    let tag = chunk.add_constant(Value::String(Arc::from(MULTI_VALUE_TAG)));
    chunk.emit_op_u16(Op::STRUCT_GET, tag, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
}

/// Coerce top-of-stack to a multi-value packet: packets pass through, scalars
/// become a one-element packet.
pub fn emit_as_multi_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let row = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], value, line);
    emit_is_multi_value_slot(&mut chunks[current], value, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_else(line);
    crate::primitives::collections::emit_array_new(chunks, current, 0, line);
    save(&mut chunks[current], row, line);
    load(&mut chunks[current], row, line);
    load(&mut chunks[current], value, line);
    crate::primitives::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    load(&mut chunks[current], row, line);
    emit_tag(chunks, current, line);
    chunks[current].emit_end(line);
}

/// Read the first value from a packet, or pass through a scalar.
/// Stack: `[value_or_packet] -> [first_or_value]`.
pub fn emit_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], value, line);
    emit_is_multi_value_slot(&mut chunks[current], value, line);
    chunks[current].emit_if(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_i32_const(0, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    load(&mut chunks[current], value, line);
    chunks[current].emit_end(line);
}

/// Read a zero-based index from a packet/array-like value.
/// Stack: `[source, index] -> [value]`.
pub fn emit_index0(chunks: &mut [Chunk], current: usize, line: u32) {
    let index = chunks[current].alloc_scratch(1);
    let source = chunks[current].alloc_scratch(1);
    save(&mut chunks[current], index, line);
    save(&mut chunks[current], source, line);
    load(&mut chunks[current], source, line);
    load(&mut chunks[current], index, line);
    crate::primitives::collections::emit_get(chunks, current, line);
}
