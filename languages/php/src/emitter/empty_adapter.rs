//! PHP `empty()` adapter.
//!
//! Kept as a focused `php.*` resolver leaf rather than living in the
//! miscellaneous adapter bucket.

use std::sync::Arc;

use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames,
};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::F64(v) => chunk.emit_f64_const(*v, line),
        Value::I32(v) => chunk.emit_i32_const(*v, line),
        Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::BigInt(v) => chunk.emit_i64_const(v.to_i64_wrapping(), line),
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::Bool(b) => chunk.emit_bool_const(*b, line),
        _ => unreachable!("push_const: unexpected value type"),
    }
}

fn push_str(chunk: &mut Chunk, value: &str, line: u32) {
    push_const(chunk, Value::String(Arc::from(value)), line);
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module.to_string(), name.to_string());
    chunks[current].emit_call(idx, argc, line);
}

pub fn emit_php_empty(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let value_slot = alloc_local(chunk);
    let _type_slot = alloc_local(chunk);

    lset(chunk, value_slot, line);

    lget(chunk, value_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    let test_bool_empty = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(test_bool_empty, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    let test_num_empty = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(test_num_empty, 1, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    let test_str_empty = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(test_str_empty, 1, line);
    chunk.emit_if_value(line);

    lget(chunk, value_slot, line);
    push_str(chunk, "", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    push_str(chunk, "0", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    core_wasm::i32_const(chunk, line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    let cs_id = class_slots::resolve(&ClassSlot::TypeIdentity, &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &cs_id, Dest::Stack, line);
    let type_slot = alloc_local(chunk);
    lset(chunk, type_slot, line);
    lget(chunk, type_slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    lget(chunk, type_slot, line);
    {
        let undef_idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(undef_idx, 1, line);
    }
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);
    push_const(chunk, Value::Bool(true), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_if_value(line);
    push_const(chunk, Value::Bool(false), line);
    chunk.emit_else(line);

    lget(chunk, value_slot, line);
    let _ = chunk;
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    lget(chunk, value_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    let base_len_slot = alloc_local(chunk);
    let extra_len_slot = alloc_local(chunk);
    lset(chunk, base_len_slot, line);

    lget(chunk, value_slot, line);
    let cs_slot = class_slots::resolve(&ClassSlot::Internal("vybe$assoc_keys_csv".to_string()), &PlainNames);
    class_slots::emit_class_get(chunk, ObjSource::Stack, &cs_slot, Dest::Stack, line);
    chunk.emit_dup(line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, base_len_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);
    push_str(chunk, "\x1F", line);
    let _ = chunk;
    call_import(chunks, current, "ecma:string", "split", 2, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    lset(chunk, extra_len_slot, line);
    lget(chunk, base_len_slot, line);
    lget(chunk, extra_len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    core_wasm::i32_const(chunk, line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
    chunk.emit_else(line);

    push_const(chunk, Value::Bool(false), line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}
