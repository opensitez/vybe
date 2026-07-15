//! PHP relational comparison — Rust inline opcode emitter.
//!
//! PHP `<` / `>` / `<=` / `>=` (and the `<=>` spaceship) compare two
//! strings lexicographically (`wasm:js-string.compare`) but fall back to the
//! numeric/dynamic comparison otherwise — unlike JS, which coerces to
//! primitive. DateTime objects are unboxed to their `__time` field
//! first so chronological comparison works.
//!
//! Mirrors the inline-emit shape of the other `languages/php/emitter`
//! adapters: writes WASM opcodes straight into the chunk, composing only
//! core ops + `vybe_emitter::ops` dynamic helpers. The shared compiler
//! routes here via the `string_aware_relational` profile flag — no
//! `profile.name == "php"` branch.

use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use vybe_emitter::ops::{emit_dyn_eq, emit_dyn_to_bool};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}
fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_op(Op::NULL, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),

        _ => {
            unreachable!("push_const: unexpected value type");
        }
    }
}
#[allow(dead_code)]
fn push_str(chunk: &mut Chunk, v: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(v)), line);
}
fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn emit_numeric_fallback(
    chunk: &mut Chunk,
    left_slot: u16,
    right_slot: u16,
    cmp_fn: fn(&mut Chunk, u32),
    line: u32,
) {
    let to_number = chunk.add_import("ecma:value", "toNumber");
    chunk.emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunk.emit_call(to_number, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunk.emit_call(to_number, 1, line);
    cmp_fn(chunk, line);
}

pub fn emit_php_loose_eq(chunks: &mut [Chunk], current: usize, _argc: u8, negate: bool, line: u32) {
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
    let test_str_a = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_a, 1, line);
    chunk.emit_if_value(line);

    lget(chunk, b_slot, line);
    let test_str_b = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_b, 1, line);
    chunk.emit_if_value(line);

    lget(chunk, a_slot, line);
    chunk.emit_call(parse_float, 1, line);
    lset(chunk, a_num_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(parse_float, 1, line);
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
    vybe_emitter::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    emit_dyn_eq(chunk, line);
    vybe_emitter::ops::emit_i32_to_bool(chunk, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(abstract_eq, 2, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    lget(chunk, a_slot, line);
    lget(chunk, b_slot, line);
    chunk.emit_call(abstract_eq, 2, line);
    chunk.emit_end(line);

    if negate {
        emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        vybe_emitter::ops::emit_i32_to_bool(chunk, line);
    }
}

/// Consume the top two stack values (`[a, b]`) and push `a <op> b` using
/// PHP comparison semantics, where `cmp_fn` emits the numeric/dynamic
/// fallback op (e.g. `emit_dyn_lt`).
pub fn emit_relational_compare(chunk: &mut Chunk, cmp_fn: fn(&mut Chunk, u32), line: u32) {
    let t_b = alloc_local(chunk);
    let t_a = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, t_b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, t_a, line);

    maybe_unbox_datetime(chunk, t_a, line);
    maybe_unbox_datetime(chunk, t_b, line);

    chunk.emit_op_u16(Op::LOCAL_GET, t_a, line);
    let test_str_ta = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_ta, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, t_b, line);
    let test_str_tb = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_tb, 1, line);
    chunk.emit_if_value(line);

    // Both strings → lexicographic compare.
    chunk.emit_op_u16(Op::LOCAL_GET, t_a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, t_b, line);
    {
        let idx = chunk.add_import("wasm:js-string", "compare");
        chunk.emit_call(idx, 2, line);
    }
    push_const(chunk, Value::I32(0), line);
    cmp_fn(chunk, line);

    chunk.emit_else(line);
    emit_numeric_fallback(chunk, t_a, t_b, cmp_fn, line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    emit_numeric_fallback(chunk, t_a, t_b, cmp_fn, line);
    chunk.emit_end(line);
}

/// If the value in `slot` is a boxed DateTime-like object, replace it
/// with its `__time` field so comparisons operate on the timestamp.
fn maybe_unbox_datetime(chunk: &mut Chunk, slot: u16, line: u32) {
    // object test: not null AND not number AND not string AND not boolean
    let obj_dt_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_dt_slot, line);
    // not null
    chunk.emit_op_u16(Op::LOCAL_GET, obj_dt_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    // AND not number
    chunk.emit_op_u16(Op::LOCAL_GET, obj_dt_slot, line);
    let test_num_dt = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(test_num_dt, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    // AND not string
    chunk.emit_op_u16(Op::LOCAL_GET, obj_dt_slot, line);
    let test_str_dt = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_dt, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    // AND not boolean
    chunk.emit_op_u16(Op::LOCAL_GET, obj_dt_slot, line);
    let test_bool_dt = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_dt, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let time_key = chunk.add_constant(Value::String(Arc::from("__time")));
    chunk.emit_op_u16(Op::STRUCT_GET, time_key, line);
    let time_slot = alloc_local(chunk);
    chunk.emit_op_u16(Op::LOCAL_SET, time_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, time_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, time_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
