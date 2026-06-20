//! PHP relational comparison — Rust inline opcode emitter.
//!
//! PHP `<` / `>` / `<=` / `>=` (and the `<=>` spaceship) compare two
//! strings lexicographically (`STR_COMPARE`) but fall back to the
//! numeric/dynamic comparison otherwise — unlike JS, which coerces to
//! primitive. DateTime objects are unboxed to their `__time` field
//! first so chronological comparison works.
//!
//! Mirrors the inline-emit shape of the other `languages/php/emitter`
//! adapters: writes WASM opcodes straight into the chunk, composing only
//! core ops + `crate::emitter::ops` dynamic helpers. The shared compiler
//! routes here via the `string_aware_relational` profile flag — no
//! `profile.name == "php"` branch.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::ops::{emit_dyn_eq, emit_dyn_to_bool};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    let s = chunk.local_count;
    chunk.local_count = s + 1;
    s
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    let idx = chunk.add_constant(val);
    chunk.emit_op_u16(Op::CONST, idx, line);
}
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

pub fn emit_php_loose_eq(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    negate: bool,
    line: u32,
) {
    let parse_float = chunks[0].add_import("ecma:number", "parseFloat");
    let abstract_eq = chunks[0].add_import("ecma:value", "abstractEq");
    let chunk = &mut chunks[current];
    let b_slot = alloc_local(chunk);
    let a_slot = alloc_local(chunk);
    let a_num_slot = alloc_local(chunk);
    let b_num_slot = alloc_local(chunk);

    lset(chunk, b_slot, line);
    lset(chunk, a_slot, line);

    lget(chunk, a_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "string", line);
    chunk.emit_op(Op::STR_EQUALS, line);
    chunk.emit_if_value(line);

    lget(chunk, b_slot, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "string", line);
    chunk.emit_op(Op::STR_EQUALS, line);
    chunk.emit_if_value(line);

    lget(chunk, a_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, parse_float, line);
    chunk.emit(1, line);
    lset(chunk, a_num_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, parse_float, line);
    chunk.emit(1, line);
    lset(chunk, b_num_slot, line);

    lget(chunk, a_num_slot, line);
    lget(chunk, a_num_slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    lget(chunk, b_num_slot, line);
    lget(chunk, b_num_slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    lget(chunk, a_num_slot, line);
    lget(chunk, b_num_slot, line);
    chunk.emit_op(Op::F64_EQ, line);
    crate::emitter::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    emit_dyn_eq(chunk, line);
    crate::emitter::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, abstract_eq, line);
    chunk.emit(2, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_op_u16(Op::CALL_IMPORT, abstract_eq, line);
    chunk.emit(2, line);
    chunk.emit_end(line);

    if negate {
        emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        crate::emitter::ops::emit_i32_to_bool(chunk, line);
    }
}

/// Consume the top two stack values (`[a, b]`) and push `a <op> b` using
/// PHP comparison semantics, where `cmp_fn` emits the numeric/dynamic
/// fallback op (e.g. `emit_dyn_lt`).
pub fn emit_relational_compare(chunk: &mut Chunk, cmp_fn: fn(&mut Chunk, u32), line: u32) {
    let t_b = alloc_local(chunk);
    let t_a = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, t_b, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, t_a, line);
    chunk.emit_op(Op::DROP, line);

    maybe_unbox_datetime(chunk, t_a, line);
    maybe_unbox_datetime(chunk, t_b, line);

    chunk.emit_op_u16(Op::LOCAL_GET, t_a, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "string", line);
    emit_dyn_eq(chunk, line);
    emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, t_b, line);
    chunk.emit_op(Op::REF_TYPEOF, line);
    push_str(chunk, "string", line);
    emit_dyn_eq(chunk, line);
    emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    // Both strings → lexicographic compare.
    chunk.emit_op_u16(Op::LOCAL_GET, t_a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t_b, line);
    chunk.emit_op(Op::STR_COMPARE, line);
    push_const(chunk, Value::I32(0), line);
    cmp_fn(chunk, line);

    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, t_a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t_b, line);
    cmp_fn(chunk, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, t_a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t_b, line);
    cmp_fn(chunk, line);
    chunk.emit_end(line);
}

/// If the value in `slot` is a boxed DateTime-like object, replace it
/// with its `__time` field so comparisons operate on the timestamp.
fn maybe_unbox_datetime(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_OBJECT, line);
    emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let time_key = chunk.add_constant(Value::String(Arc::from("__time")));
    chunk.emit_op_u16(Op::STRUCT_GET, time_key, line);
    let time_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, time_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, time_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, time_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
