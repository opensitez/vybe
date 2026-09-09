//! PHP superglobal shape adapters.
//!
//! The raw request/environment data is shared (`http_request`, `http_form`,
//! `http_cookie`, WASI env). This file owns only PHP's public array shape:
//! `$_SERVER['PHP_SELF']`, `$_FILES[*]`, and name-keyed `$_ENV`.

use std::sync::Arc;

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

fn emit_is_null_or_undefined(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    lget(chunk, slot, line);
    let undef = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn map_get_const(chunks: &mut [Chunk], current: usize, map_slot: u16, key: &str, line: u32) {
    lget(&mut chunks[current], map_slot, line);
    push_str(&mut chunks[current], key, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
}

fn map_set_const_value(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    key: &str,
    value_slot: u16,
    line: u32,
) {
    lget(&mut chunks[current], map_slot, line);
    push_str(&mut chunks[current], key, line);
    lget(&mut chunks[current], value_slot, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn map_set_const_literal(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    key: &str,
    value: Value,
    line: u32,
) {
    lget(&mut chunks[current], map_slot, line);
    push_str(&mut chunks[current], key, line);
    push_const(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn bump_index(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_f64_const(1.0, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, slot, line);
}

pub fn emit_php_superglobal_server(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    vybe_compiler::primitives::http_request_env::emit_environ(chunks, current, line);
    let server_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], server_slot, line);

    map_get_const(chunks, current, server_slot, "SCRIPT_NAME", line);
    let self_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], self_slot, line);
    emit_is_null_or_undefined(&mut chunks[current], self_slot, line);
    chunks[current].emit_if(line);
    push_str(&mut chunks[current], "", line);
    lset(&mut chunks[current], self_slot, line);
    chunks[current].emit_end(line);

    map_set_const_value(chunks, current, server_slot, "PHP_SELF", self_slot, line);
    lget(&mut chunks[current], server_slot, line);
}

pub fn emit_php_superglobal_files(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    vybe_compiler::primitives::http_form::emit_body_files(chunks, current, line);
    vybe_compiler::primitives::collections::emit_iter_entries(chunks, current, line);
    let entries_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], entries_slot, line);

    call_import(chunks, current, "ecma:map", "new", 0, line);
    let out_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], out_slot, line);

    lget(&mut chunks[current], entries_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    let len_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], len_slot, line);

    let i_slot = alloc_local(&mut chunks[current]);
    chunks[current].emit_f64_const(0.0, line);
    lset(&mut chunks[current], i_slot, line);

    let pair_slot = alloc_local(&mut chunks[current]);
    let field_slot = alloc_local(&mut chunks[current]);
    let upload_slot = alloc_local(&mut chunks[current]);
    let file_slot = alloc_local(&mut chunks[current]);
    let value_slot = alloc_local(&mut chunks[current]);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], entries_slot, line);
    lget(&mut chunks[current], i_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], pair_slot, line);

    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], field_slot, line);

    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], upload_slot, line);

    call_import(chunks, current, "ecma:map", "new", 0, line);
    lset(&mut chunks[current], file_slot, line);

    map_get_const(chunks, current, upload_slot, "filename", line);
    lset(&mut chunks[current], value_slot, line);
    map_set_const_value(chunks, current, file_slot, "name", value_slot, line);
    map_get_const(chunks, current, upload_slot, "type", line);
    lset(&mut chunks[current], value_slot, line);
    map_set_const_value(chunks, current, file_slot, "type", value_slot, line);
    map_set_const_literal(
        chunks,
        current,
        file_slot,
        "tmp_name",
        Value::String(Arc::from("")),
        line,
    );
    map_get_const(chunks, current, upload_slot, "size", line);
    lset(&mut chunks[current], value_slot, line);
    map_set_const_value(chunks, current, file_slot, "size", value_slot, line);
    map_set_const_literal(chunks, current, file_slot, "error", Value::I32(0), line);

    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], field_slot, line);
    lget(&mut chunks[current], file_slot, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    bump_index(&mut chunks[current], i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], out_slot, line);
}

pub fn emit_php_superglobal_env(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
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

    call_import(chunks, current, "ecma:map", "new", 0, line);
    let out_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], out_slot, line);

    lget(&mut chunks[current], pairs_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    let len_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], len_slot, line);

    let i_slot = alloc_local(&mut chunks[current]);
    chunks[current].emit_f64_const(0.0, line);
    lset(&mut chunks[current], i_slot, line);

    let pair_slot = alloc_local(&mut chunks[current]);
    let key_slot = alloc_local(&mut chunks[current]);
    let value_slot = alloc_local(&mut chunks[current]);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], pairs_slot, line);
    lget(&mut chunks[current], i_slot, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lset(&mut chunks[current], pair_slot, line);
    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lset(&mut chunks[current], key_slot, line);
    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    call_import(chunks, current, "ecma:array", "get", 2, line);
    lset(&mut chunks[current], value_slot, line);

    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], key_slot, line);
    lget(&mut chunks[current], value_slot, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    bump_index(&mut chunks[current], i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], out_slot, line);
}
