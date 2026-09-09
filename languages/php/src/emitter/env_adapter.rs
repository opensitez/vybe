//! PHP environment-function adapters.
//!
//! `getenv`/`putenv` are PHP semantics over the WASI environment surface:
//! the real environment is read from `wasi:cli/environment.get-environment`,
//! while `putenv` mutates a PHP process-local overlay stored in shared globals.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const PHP_ENV_OVERLAY_GLOBAL: &str = "__php_env_overlay";

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

fn emit_is_null_or_undefined(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    lget(chunk, slot, line);
    let undef = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn ensure_overlay(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    vybe_compiler::primitives::globals::emit_read(
        &mut chunks[current],
        PHP_ENV_OVERLAY_GLOBAL,
        line,
    );
    let overlay_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], overlay_slot, line);

    emit_is_null_or_undefined(&mut chunks[current], overlay_slot, line);
    chunks[current].emit_if(line);
    call_import(chunks, current, "ecma:map", "new", 0, line);
    lset(&mut chunks[current], overlay_slot, line);
    lget(&mut chunks[current], overlay_slot, line);
    vybe_compiler::primitives::globals::emit_write(
        &mut chunks[current],
        PHP_ENV_OVERLAY_GLOBAL,
        line,
    );
    chunks[current].emit_end(line);

    overlay_slot
}

fn emit_pair_to_map(
    chunks: &mut [Chunk],
    current: usize,
    out_slot: u16,
    pair_slot: u16,
    line: u32,
) {
    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_copy_pairs_to_map(
    chunks: &mut [Chunk],
    current: usize,
    pairs_slot: u16,
    out_slot: u16,
    line: u32,
) {
    let i_slot = alloc_local(&mut chunks[current]);
    chunks[current].emit_f64_const(0.0, line);
    lset(&mut chunks[current], i_slot, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], pairs_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], pairs_slot, line);
    lget(&mut chunks[current], i_slot, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    let pair_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], pair_slot, line);
    emit_pair_to_map(chunks, current, out_slot, pair_slot, line);

    lget(&mut chunks[current], i_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
}

fn emit_scan_pairs_for_name(
    chunks: &mut [Chunk],
    current: usize,
    pairs_slot: u16,
    name_slot: u16,
    line: u32,
) {
    let result_slot = alloc_local(&mut chunks[current]);
    chunks[current].emit_bool_const(false, line);
    lset(&mut chunks[current], result_slot, line);

    let i_slot = alloc_local(&mut chunks[current]);
    chunks[current].emit_f64_const(0.0, line);
    lset(&mut chunks[current], i_slot, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], pairs_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], pairs_slot, line);
    lget(&mut chunks[current], i_slot, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    let pair_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], pair_slot, line);

    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lget(&mut chunks[current], name_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lset(&mut chunks[current], result_slot, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], i_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], result_slot, line);
}

pub fn emit_php_getenv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }

    call_import(
        chunks,
        current,
        "wasi:cli/environment",
        "get-environment",
        0,
        line,
    );
    let pairs_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], pairs_slot, line);
    let overlay_slot = ensure_overlay(chunks, current, line);

    if argc == 0 {
        call_import(chunks, current, "ecma:map", "new", 0, line);
        let out_slot = alloc_local(&mut chunks[current]);
        lset(&mut chunks[current], out_slot, line);
        emit_copy_pairs_to_map(chunks, current, pairs_slot, out_slot, line);
        lget(&mut chunks[current], overlay_slot, line);
        call_import(chunks, current, "ecma:map", "entries", 1, line);
        let overlay_pairs_slot = alloc_local(&mut chunks[current]);
        lset(&mut chunks[current], overlay_pairs_slot, line);
        emit_copy_pairs_to_map(chunks, current, overlay_pairs_slot, out_slot, line);
        lget(&mut chunks[current], out_slot, line);
        return;
    }

    call_import(chunks, current, "ecma:string", "String", 1, line);
    let name_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], name_slot, line);

    lget(&mut chunks[current], overlay_slot, line);
    lget(&mut chunks[current], name_slot, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], overlay_slot, line);
    lget(&mut chunks[current], name_slot, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    chunks[current].emit_else(line);
    emit_scan_pairs_for_name(chunks, current, pairs_slot, name_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_php_putenv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_bool_const(false, line);
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }

    call_import(chunks, current, "ecma:string", "String", 1, line);
    let assignment_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], assignment_slot, line);
    let overlay_slot = ensure_overlay(chunks, current, line);

    lget(&mut chunks[current], assignment_slot, line);
    push_str(&mut chunks[current], "=", line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    let pos_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], pos_slot, line);

    lget(&mut chunks[current], pos_slot, line);
    chunks[current].emit_f64_const(-1.0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], overlay_slot, line);
    lget(&mut chunks[current], assignment_slot, line);
    call_import(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], overlay_slot, line);
    lget(&mut chunks[current], assignment_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    lget(&mut chunks[current], pos_slot, line);
    call_import(chunks, current, "wasm:js-string", "substring", 3, line);
    lget(&mut chunks[current], assignment_slot, line);
    lget(&mut chunks[current], pos_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lget(&mut chunks[current], assignment_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    call_import(chunks, current, "wasm:js-string", "substring", 3, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
}
