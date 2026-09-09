//! PHP HTTP/SAPI adapters.
//!
//! PHP owns the stdlib call shape; HTTP state itself is handled by `node:http`
//! and shared request/cookie primitives.

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

fn emit_status_or_default(chunk: &mut Chunk, status: u16, line: u32) {
    let current = alloc_local(chunk);
    chunk.emit_call(status, 0, line);
    lset(chunk, current, line);
    lget(chunk, current, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, current, line);
    chunk.emit_else(line);
    push_const(chunk, Value::F64(200.0), line);
    chunk.emit_end(line);
}

fn emit_apply_header(
    chunks: &mut [Chunk],
    current: usize,
    header_slot: u16,
    replace_slot: Option<u16>,
    line: u32,
) {
    let set_status = chunks[current].add_import("node:http", "set_status");
    let set_header = chunks[current].add_import("node:http", "set_header");
    let add_header = chunks[current].add_import("node:http", "add_header");
    let index_of = chunks[current].add_import("ecma:string", "indexOf");
    let starts_with = chunks[current].add_import("ecma:string", "startsWith");
    let substring = chunks[current].add_import("wasm:js-string", "substring");
    let length = chunks[current].add_import("wasm:js-string", "length");
    let trim = chunks[current].add_import("ecma:string", "trim");
    let to_number = chunks[current].add_import("ecma:value", "toNumber");

    let at = alloc_local(&mut chunks[current]);
    let rest = alloc_local(&mut chunks[current]);
    let name_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];

    lget(chunk, header_slot, line);
    push_str(chunk, "HTTP/", line);
    chunk.emit_call(starts_with, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    lget(chunk, header_slot, line);
    push_str(chunk, " ", line);
    chunk.emit_call(index_of, 2, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    lset(chunk, at, line);

    lget(chunk, header_slot, line);
    lget(chunk, at, line);
    lget(chunk, header_slot, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_call(substring, 3, line);
    lset(chunk, rest, line);

    lget(chunk, rest, line);
    push_str(chunk, " ", line);
    chunk.emit_call(index_of, 2, line);
    lset(chunk, at, line);

    lget(chunk, at, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);
    lget(chunk, rest, line);
    chunk.emit_i32_const(0, line);
    lget(chunk, at, line);
    chunk.emit_call(substring, 3, line);
    chunk.emit_else(line);
    lget(chunk, rest, line);
    chunk.emit_end(line);
    chunk.emit_call(to_number, 1, line);
    chunk.emit_call(set_status, 1, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_else(line);

    lget(chunk, header_slot, line);
    push_str(chunk, ":", line);
    chunk.emit_call(index_of, 2, line);
    lset(chunk, at, line);

    lget(chunk, at, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    chunk.emit_if(line);

    lget(chunk, header_slot, line);
    chunk.emit_i32_const(0, line);
    lget(chunk, at, line);
    chunk.emit_call(substring, 3, line);
    chunk.emit_call(trim, 1, line);
    lset(chunk, name_slot, line);

    lget(chunk, name_slot, line);
    lget(chunk, header_slot, line);
    lget(chunk, at, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    lget(chunk, header_slot, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_call(substring, 3, line);
    chunk.emit_call(trim, 1, line);

    match replace_slot {
        Some(slot) => {
            lget(chunk, slot, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_if(line);
            chunk.emit_call(set_header, 2, line);
            chunk.emit_else(line);
            chunk.emit_call(add_header, 2, line);
            chunk.emit_end(line);
        }
        None => chunk.emit_call(set_header, 2, line),
    }
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_send_cookie(chunks: &mut [Chunk], current: usize, line: u32) {
    let add_header = chunks[current].add_import("node:http", "add_header");
    let cookie_slot = alloc_local(&mut chunks[current]);
    let chunk = &mut chunks[current];
    lset(chunk, cookie_slot, line);
    push_str(chunk, "Set-Cookie", line);
    lget(chunk, cookie_slot, line);
    chunk.emit_call(add_header, 2, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_php_sapi_name(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::globals::emit_read(
        chunk,
        vybe_compiler::primitives::http_request_env::REQUEST_GLOBAL,
        line,
    );
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    push_str(chunk, "cli", line);
    chunk.emit_else(line);
    push_str(chunk, "vybex-server", line);
    chunk.emit_end(line);
}

pub fn emit_php_http_response_code(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let status = chunks[current].add_import("node:http", "status");
    let set_status = chunks[current].add_import("node:http", "set_status");
    let chunk = &mut chunks[current];

    if argc == 0 {
        emit_status_or_default(chunk, status, line);
        return;
    }

    let code = alloc_local(chunk);
    lset(chunk, code, line);
    lget(chunk, code, line);
    push_const(chunk, Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    chunk.emit_if(line);
    lget(chunk, code, line);
    chunk.emit_call(set_status, 1, line);
    chunk.emit_op(Op::DROP, line);
    lget(chunk, code, line);
    chunk.emit_else(line);
    emit_status_or_default(chunk, status, line);
    chunk.emit_end(line);
}

pub fn emit_php_header(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let set_status_import = chunks[current].add_import("node:http", "set_status");
    let status_import = chunks[current].add_import("node:http", "status");
    let string_import = chunks[current].add_import("ecma:string", "String");
    let lower_import = chunks[current].add_import("ecma:string", "toLowerCase");
    let starts_with_import = chunks[current].add_import("ecma:string", "startsWith");
    let chunk = &mut chunks[current];

    let header_slot = alloc_local(chunk);
    let replace_slot = if argc >= 2 {
        Some(alloc_local(chunk))
    } else {
        None
    };
    let response_code_slot = if argc >= 3 {
        Some(alloc_local(chunk))
    } else {
        None
    };

    if let Some(slot) = response_code_slot {
        lset(chunk, slot, line);
    }
    if let Some(slot) = replace_slot {
        lset(chunk, slot, line);
    }
    lset(chunk, header_slot, line);

    emit_apply_header(chunks, current, header_slot, replace_slot, line);

    let chunk = &mut chunks[current];
    if let Some(slot) = response_code_slot {
        lget(chunk, slot, line);
        push_const(chunk, Value::F64(0.0), line);
        vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
        chunk.emit_if(line);
        lget(chunk, slot, line);
        chunk.emit_call(set_status_import, 1, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_end(line);
    }

    if let Some(slot) = response_code_slot {
        lget(chunk, slot, line);
        push_const(chunk, Value::F64(0.0), line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        chunk.emit_if(line);
    }

    lget(chunk, header_slot, line);
    chunk.emit_call(string_import, 1, line);
    chunk.emit_call(lower_import, 1, line);
    push_str(chunk, "location:", line);
    chunk.emit_call(starts_with_import, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);

    chunk.emit_call(status_import, 0, line);
    push_const(chunk, Value::F64(200.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if(line);

    push_const(chunk, Value::F64(302.0), line);
    chunk.emit_call(set_status_import, 1, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    if response_code_slot.is_some() {
        chunk.emit_end(line);
    }

    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_php_setcookie(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_setcookie_inner(chunks, current, argc, true, line);
}

pub fn emit_php_setrawcookie(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_setcookie_inner(chunks, current, argc, false, line);
}

fn emit_setcookie_inner(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    encode_value: bool,
    line: u32,
) {
    let attrs_slot = if argc >= 3 {
        Some(alloc_local(&mut chunks[current]))
    } else {
        None
    };
    let value_slot = alloc_local(&mut chunks[current]);
    let name_slot = alloc_local(&mut chunks[current]);

    if let Some(slot) = attrs_slot {
        lset(&mut chunks[current], slot, line);
    }
    if argc >= 2 {
        lset(&mut chunks[current], value_slot, line);
    } else {
        push_str(&mut chunks[current], "", line);
        lset(&mut chunks[current], value_slot, line);
    }
    lset(&mut chunks[current], name_slot, line);

    lget(&mut chunks[current], name_slot, line);
    lget(&mut chunks[current], value_slot, line);
    if encode_value {
        crate::emitter::string_adapter::emit_urlencode(chunks, current, 1, line);
    }
    if let Some(slot) = attrs_slot {
        lget(&mut chunks[current], slot, line);
    }
    vybe_compiler::primitives::http_cookie::emit_serialize(
        chunks,
        current,
        if attrs_slot.is_some() { 3 } else { 2 },
        line,
    );
    emit_send_cookie(chunks, current, line);
    push_const(&mut chunks[current], Value::Bool(true), line);
}
