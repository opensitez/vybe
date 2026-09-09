//! PHP session adapters.
//!
//! Session lifecycle delegates to the shared HTTP session primitive. PHP owns
//! only the stdlib spelling and its `$_SESSION`/legacy-global mirror behavior.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const PHP_SESSION_COOKIE_NAME: &str = "PHPSESSID";
const PHP_SESSION_ID_GLOBAL: &str = "__php_session_id";
const PHP_SESSION_STARTED_GLOBAL: &str = "__php_session_started";
const PHP_SESSION_DESTROYED_GLOBAL: &str = "__php_session_destroyed";
const PHP_SESSION_NEEDS_COOKIE_GLOBAL: &str = "__php_session_needs_cookie";

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

fn global_set(chunk: &mut Chunk, key: &str, line: u32) {
    vybe_compiler::primitives::globals::emit_write(chunk, key, line);
}

fn global_get(chunk: &mut Chunk, key: &str, line: u32) {
    vybe_compiler::primitives::globals::emit_read(chunk, key, line);
}

fn emit_is_null_or_undefined(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    lget(chunk, slot, line);
    let undef = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_global_or_default(chunk: &mut Chunk, key: &str, default: Value, line: u32) {
    global_get(chunk, key, line);
    let slot = alloc_local(chunk);
    lset(chunk, slot, line);
    emit_is_null_or_undefined(chunk, slot, line);
    chunk.emit_if_value(line);
    push_const(chunk, default, line);
    chunk.emit_else(line);
    lget(chunk, slot, line);
    chunk.emit_end(line);
}

fn map_set_from_global_or_default(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    field: &str,
    global: &str,
    default: Value,
    line: u32,
) {
    lget(&mut chunks[current], map_slot, line);
    push_str(&mut chunks[current], field, line);
    emit_global_or_default(&mut chunks[current], global, default, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_global_from_option_if_present(
    chunks: &mut [Chunk],
    current: usize,
    options_slot: u16,
    field: &str,
    global: &str,
    line: u32,
) {
    lget(&mut chunks[current], options_slot, line);
    push_str(&mut chunks[current], field, line);
    call_import(chunks, current, "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], options_slot, line);
    push_str(&mut chunks[current], field, line);
    call_import(chunks, current, "ecma:map", "get", 2, line);
    global_set(&mut chunks[current], global, line);
    chunks[current].emit_end(line);
}

pub fn emit_php_session_get_cookie_params(
    chunks: &mut [Chunk],
    current: usize,
    _argc: u8,
    line: u32,
) {
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let out_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], out_slot, line);
    map_set_from_global_or_default(
        chunks,
        current,
        out_slot,
        "lifetime",
        "__php_session_cookie_lifetime",
        Value::I32(0),
        line,
    );
    map_set_from_global_or_default(
        chunks,
        current,
        out_slot,
        "path",
        "__php_session_cookie_path",
        Value::String(Arc::from("/")),
        line,
    );
    map_set_from_global_or_default(
        chunks,
        current,
        out_slot,
        "domain",
        "__php_session_cookie_domain",
        Value::String(Arc::from("")),
        line,
    );
    map_set_from_global_or_default(
        chunks,
        current,
        out_slot,
        "secure",
        "__php_session_cookie_secure",
        Value::Bool(false),
        line,
    );
    map_set_from_global_or_default(
        chunks,
        current,
        out_slot,
        "httponly",
        "__php_session_cookie_httponly",
        Value::Bool(false),
        line,
    );
    map_set_from_global_or_default(
        chunks,
        current,
        out_slot,
        "samesite",
        "__php_session_cookie_samesite",
        Value::String(Arc::from("")),
        line,
    );
    lget(&mut chunks[current], out_slot, line);
}

pub fn emit_php_session_set_cookie_params(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let httponly_slot = alloc_local(&mut chunks[current]);
    let secure_slot = alloc_local(&mut chunks[current]);
    let domain_slot = alloc_local(&mut chunks[current]);
    let path_slot = alloc_local(&mut chunks[current]);
    let lifetime_slot = alloc_local(&mut chunks[current]);

    if argc >= 5 {
        lset(&mut chunks[current], httponly_slot, line);
    } else {
        push_const(&mut chunks[current], Value::Bool(false), line);
        lset(&mut chunks[current], httponly_slot, line);
    }
    if argc >= 4 {
        lset(&mut chunks[current], secure_slot, line);
    } else {
        push_const(&mut chunks[current], Value::Bool(false), line);
        lset(&mut chunks[current], secure_slot, line);
    }
    if argc >= 3 {
        lset(&mut chunks[current], domain_slot, line);
    } else {
        push_str(&mut chunks[current], "", line);
        lset(&mut chunks[current], domain_slot, line);
    }
    if argc >= 2 {
        lset(&mut chunks[current], path_slot, line);
    } else {
        push_str(&mut chunks[current], "/", line);
        lset(&mut chunks[current], path_slot, line);
    }
    if argc >= 1 {
        lset(&mut chunks[current], lifetime_slot, line);
    } else {
        push_const(&mut chunks[current], Value::I32(0), line);
        lset(&mut chunks[current], lifetime_slot, line);
    }

    lget(&mut chunks[current], lifetime_slot, line);
    super::array_adapter::emit_php_is_array(chunks, current, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    set_global_from_option_if_present(
        chunks,
        current,
        lifetime_slot,
        "lifetime",
        "__php_session_cookie_lifetime",
        line,
    );
    set_global_from_option_if_present(
        chunks,
        current,
        lifetime_slot,
        "path",
        "__php_session_cookie_path",
        line,
    );
    set_global_from_option_if_present(
        chunks,
        current,
        lifetime_slot,
        "domain",
        "__php_session_cookie_domain",
        line,
    );
    set_global_from_option_if_present(
        chunks,
        current,
        lifetime_slot,
        "secure",
        "__php_session_cookie_secure",
        line,
    );
    set_global_from_option_if_present(
        chunks,
        current,
        lifetime_slot,
        "httponly",
        "__php_session_cookie_httponly",
        line,
    );
    set_global_from_option_if_present(
        chunks,
        current,
        lifetime_slot,
        "samesite",
        "__php_session_cookie_samesite",
        line,
    );
    chunks[current].emit_else(line);
    lget(&mut chunks[current], lifetime_slot, line);
    global_set(
        &mut chunks[current],
        "__php_session_cookie_lifetime",
        line,
    );
    lget(&mut chunks[current], path_slot, line);
    global_set(&mut chunks[current], "__php_session_cookie_path", line);
    lget(&mut chunks[current], domain_slot, line);
    global_set(&mut chunks[current], "__php_session_cookie_domain", line);
    lget(&mut chunks[current], secure_slot, line);
    global_set(&mut chunks[current], "__php_session_cookie_secure", line);
    lget(&mut chunks[current], httponly_slot, line);
    global_set(&mut chunks[current], "__php_session_cookie_httponly", line);
    chunks[current].emit_end(line);

    push_const(&mut chunks[current], Value::Bool(true), line);
}

fn emit_get_set_global(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    global: &str,
    default: Value,
    line: u32,
) {
    let new_slot = if argc > 0 {
        let slot = alloc_local(&mut chunks[current]);
        lset(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    emit_global_or_default(&mut chunks[current], global, default, line);
    let old_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], old_slot, line);
    if let Some(slot) = new_slot {
        lget(&mut chunks[current], slot, line);
        global_set(&mut chunks[current], global, line);
    }
    lget(&mut chunks[current], old_slot, line);
}

fn emit_eq_str(chunk: &mut Chunk, slot: u16, value: &str, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, value, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_substring_slots(
    chunks: &mut [Chunk],
    current: usize,
    text_slot: u16,
    start_slot: u16,
    end_slot: u16,
    line: u32,
) {
    lget(&mut chunks[current], text_slot, line);
    lget(&mut chunks[current], start_slot, line);
    lget(&mut chunks[current], end_slot, line);
    call_import(chunks, current, "wasm:js-string", "substring", 3, line);
}

fn emit_index_of_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    text_slot: u16,
    needle: &str,
    line: u32,
) {
    lget(&mut chunks[current], text_slot, line);
    push_str(&mut chunks[current], needle, line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
}

fn bump_slot(chunk: &mut Chunk, slot: u16, by: f64, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_f64_const(by, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, slot, line);
}

pub fn emit_php_session_cache_limiter(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    emit_get_set_global(
        chunks,
        current,
        argc,
        "__php_session_cache_limiter",
        Value::String(Arc::from("nocache")),
        line,
    );
}

pub fn emit_php_session_cache_expire(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    emit_get_set_global(
        chunks,
        current,
        argc,
        "__php_session_cache_expire",
        Value::I32(180),
        line,
    );
}

pub fn emit_php_session_module_name(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    push_str(&mut chunks[current], "files", line);
}

pub fn emit_php_session_save_path(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    emit_get_set_global(
        chunks,
        current,
        argc,
        "__php_session_save_path",
        Value::String(Arc::from("")),
        line,
    );
}

pub fn emit_php_session_create_id(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        push_str(
            &mut chunks[current],
            "0123456789abcdef0123456789abcdef",
            line,
        );
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    push_str(
        &mut chunks[current],
        "0123456789abcdef0123456789abcdef",
        line,
    );
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 2, line);
}

pub fn emit_php_session_gc(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::I32(0), line);
}

pub fn emit_php_session_set_save_handler(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    push_const(&mut chunks[current], Value::Bool(true), line);
}

pub fn emit_php_session_encode(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    global_get(&mut chunks[current], "$_SESSION", line);
    vybe_compiler::primitives::collections::emit_iter_entries(chunks, current, line);
    let entries_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], entries_slot, line);

    lget(&mut chunks[current], entries_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    let len_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], len_slot, line);

    let out_slot = alloc_local(&mut chunks[current]);
    push_str(&mut chunks[current], "", line);
    lset(&mut chunks[current], out_slot, line);

    let i_slot = alloc_local(&mut chunks[current]);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    lset(&mut chunks[current], i_slot, line);

    let pair_slot = alloc_local(&mut chunks[current]);
    let key_slot = alloc_local(&mut chunks[current]);
    let value_slot = alloc_local(&mut chunks[current]);
    let blob_slot = alloc_local(&mut chunks[current]);

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
    push_const(&mut chunks[current], Value::F64(0.0), line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    lset(&mut chunks[current], key_slot, line);

    lget(&mut chunks[current], pair_slot, line);
    push_const(&mut chunks[current], Value::F64(1.0), line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], value_slot, line);

    lget(&mut chunks[current], value_slot, line);
    super::serialization_adapter::emit_php_serialize(chunks, current, 1, line);
    lset(&mut chunks[current], blob_slot, line);

    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], key_slot, line);
    push_str(&mut chunks[current], "|", line);
    lget(&mut chunks[current], blob_slot, line);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 4, line);
    lset(&mut chunks[current], out_slot, line);

    bump_slot(&mut chunks[current], i_slot, 1.0, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], out_slot, line);
}

pub fn emit_php_session_decode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        call_import(chunks, current, "ecma:map", "new", 0, line);
        global_set(&mut chunks[current], "$_SESSION", line);
        push_const(&mut chunks[current], Value::Bool(true), line);
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    let data_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], data_slot, line);

    call_import(chunks, current, "ecma:map", "new", 0, line);
    let session_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], session_slot, line);

    lget(&mut chunks[current], data_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    let len_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], len_slot, line);

    let i_slot = alloc_local(&mut chunks[current]);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    lset(&mut chunks[current], i_slot, line);

    let rest_slot = alloc_local(&mut chunks[current]);
    let rel_slot = alloc_local(&mut chunks[current]);
    let key_slot = alloc_local(&mut chunks[current]);
    let blob_start_slot = alloc_local(&mut chunks[current]);
    let semi_rel_slot = alloc_local(&mut chunks[current]);
    let semi_slot = alloc_local(&mut chunks[current]);
    let tag_slot = alloc_local(&mut chunks[current]);
    let value_slot = alloc_local(&mut chunks[current]);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    emit_substring_slots(chunks, current, data_slot, i_slot, len_slot, line);
    lset(&mut chunks[current], rest_slot, line);
    emit_index_of_from_slot(chunks, current, rest_slot, "|", line);
    lset(&mut chunks[current], rel_slot, line);

    lget(&mut chunks[current], rel_slot, line);
    push_const(&mut chunks[current], Value::F64(0.0), line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], len_slot, line);
    lset(&mut chunks[current], i_slot, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], rel_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    let key_end_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], key_end_slot, line);
    emit_substring_slots(chunks, current, data_slot, i_slot, key_end_slot, line);
    lset(&mut chunks[current], key_slot, line);

    lget(&mut chunks[current], key_end_slot, line);
    push_const(&mut chunks[current], Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    lset(&mut chunks[current], blob_start_slot, line);

    lget(&mut chunks[current], data_slot, line);
    lget(&mut chunks[current], blob_start_slot, line);
    lget(&mut chunks[current], blob_start_slot, line);
    push_const(&mut chunks[current], Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    call_import(chunks, current, "wasm:js-string", "substring", 3, line);
    lset(&mut chunks[current], tag_slot, line);

    emit_substring_slots(chunks, current, data_slot, blob_start_slot, len_slot, line);
    lset(&mut chunks[current], rest_slot, line);
    emit_index_of_from_slot(chunks, current, rest_slot, ";", line);
    lset(&mut chunks[current], semi_rel_slot, line);
    lget(&mut chunks[current], blob_start_slot, line);
    lget(&mut chunks[current], semi_rel_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    lset(&mut chunks[current], semi_slot, line);

    emit_eq_str(&mut chunks[current], tag_slot, "N", line);
    chunks[current].emit_if(line);
    push_const(&mut chunks[current], Value::Null, line);
    lset(&mut chunks[current], value_slot, line);
    chunks[current].emit_else(line);

    emit_eq_str(&mut chunks[current], tag_slot, "s", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], rest_slot, line);
    push_str(&mut chunks[current], "\"", line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    lget(&mut chunks[current], blob_start_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    let q1_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], q1_slot, line);
    lget(&mut chunks[current], q1_slot, line);
    push_const(&mut chunks[current], Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    let val_start_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], val_start_slot, line);
    emit_substring_slots(chunks, current, data_slot, val_start_slot, len_slot, line);
    lset(&mut chunks[current], rest_slot, line);
    lget(&mut chunks[current], rest_slot, line);
    push_str(&mut chunks[current], "\";", line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    lget(&mut chunks[current], val_start_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    let q2_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], q2_slot, line);
    emit_substring_slots(chunks, current, data_slot, val_start_slot, q2_slot, line);
    lset(&mut chunks[current], value_slot, line);
    lget(&mut chunks[current], q2_slot, line);
    push_const(&mut chunks[current], Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    lset(&mut chunks[current], semi_slot, line);
    chunks[current].emit_else(line);

    emit_eq_str(&mut chunks[current], tag_slot, "b", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], data_slot, line);
    lget(&mut chunks[current], blob_start_slot, line);
    push_const(&mut chunks[current], Value::F64(2.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    lget(&mut chunks[current], semi_slot, line);
    call_import(chunks, current, "wasm:js-string", "substring", 3, line);
    push_str(&mut chunks[current], "1", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    lset(&mut chunks[current], value_slot, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], data_slot, line);
    lget(&mut chunks[current], blob_start_slot, line);
    push_const(&mut chunks[current], Value::F64(2.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    lget(&mut chunks[current], semi_slot, line);
    call_import(chunks, current, "wasm:js-string", "substring", 3, line);
    call_import(chunks, current, "ecma:number", "Number", 1, line);
    lset(&mut chunks[current], value_slot, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], session_slot, line);
    lget(&mut chunks[current], key_slot, line);
    lget(&mut chunks[current], value_slot, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(&mut chunks[current], semi_slot, line);
    push_const(&mut chunks[current], Value::F64(1.0), line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    lset(&mut chunks[current], i_slot, line);

    chunks[current].emit_end(line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], session_slot, line);
    global_set(&mut chunks[current], "$_SESSION", line);
    push_const(&mut chunks[current], Value::Bool(true), line);
}

fn emit_php_session_sync_legacy_id(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::http_session::emit_id(
        chunks,
        current,
        PHP_SESSION_COOKIE_NAME,
        line,
    );
    vybe_compiler::primitives::globals::emit_write(
        &mut chunks[current],
        PHP_SESSION_ID_GLOBAL,
        line,
    );
}

fn emit_php_session_load_working_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::http_session::emit_data(chunks, current, line);
    vybe_compiler::primitives::collections::emit_map_clone(chunks, current, line);
    vybe_compiler::primitives::globals::emit_write(&mut chunks[current], "$_SESSION", line);
}

fn emit_php_session_commit_working_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::globals::emit_read(&mut chunks[current], "$_SESSION", line);
    vybe_compiler::primitives::collections::emit_map_clone(chunks, current, line);
    vybe_compiler::primitives::globals::emit_write(
        &mut chunks[current],
        vybe_compiler::primitives::http_session::SESSION_DATA_GLOBAL,
        line,
    );
}

fn emit_php_session_set_status(chunks: &mut [Chunk], current: usize, status: i32, line: u32) {
    push_const(&mut chunks[current], Value::I32(status), line);
    vybe_compiler::primitives::globals::emit_write(
        &mut chunks[current],
        vybe_compiler::primitives::http_session::SESSION_STATUS_GLOBAL,
        line,
    );
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

pub fn emit_php_session_start(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc > 0 {
        chunks[current].emit_op(Op::DROP, line);
    }
    vybe_compiler::primitives::http_session::emit_start(
        chunks,
        current,
        PHP_SESSION_COOKIE_NAME,
        line,
    );
    let result_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], result_slot, line);
    emit_php_session_sync_legacy_id(chunks, current, line);
    emit_php_session_load_working_copy(chunks, current, line);

    push_const(&mut chunks[current], Value::Bool(true), line);
    vybe_compiler::primitives::globals::emit_write(
        &mut chunks[current],
        PHP_SESSION_STARTED_GLOBAL,
        line,
    );
    push_const(&mut chunks[current], Value::Bool(false), line);
    vybe_compiler::primitives::globals::emit_write(
        &mut chunks[current],
        PHP_SESSION_DESTROYED_GLOBAL,
        line,
    );
    lget(&mut chunks[current], result_slot, line);
}

pub fn emit_php_session_id(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let new_slot = if argc > 0 {
        let slot = alloc_local(&mut chunks[current]);
        lset(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    vybe_compiler::primitives::http_session::emit_id(
        chunks,
        current,
        PHP_SESSION_COOKIE_NAME,
        line,
    );
    let old_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], old_slot, line);
    if let Some(slot) = new_slot {
        lget(&mut chunks[current], slot, line);
        vybe_compiler::primitives::globals::emit_write(
            &mut chunks[current],
            vybe_compiler::primitives::http_session::SESSION_ID_GLOBAL,
            line,
        );
        lget(&mut chunks[current], slot, line);
        vybe_compiler::primitives::globals::emit_write(
            &mut chunks[current],
            PHP_SESSION_ID_GLOBAL,
            line,
        );
    }
    lget(&mut chunks[current], old_slot, line);
}

pub fn emit_php_session_name(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let new_slot = if argc > 0 {
        let slot = alloc_local(&mut chunks[current]);
        lset(&mut chunks[current], slot, line);
        Some(slot)
    } else {
        None
    };
    vybe_compiler::primitives::http_session::emit_name(
        chunks,
        current,
        PHP_SESSION_COOKIE_NAME,
        line,
    );
    let old_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], old_slot, line);
    if let Some(slot) = new_slot {
        lget(&mut chunks[current], slot, line);
        vybe_compiler::primitives::http_session::emit_set_name(chunks, current, line);
    }
    lget(&mut chunks[current], old_slot, line);
}

pub fn emit_php_session_status(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    vybe_compiler::primitives::http_session::emit_status(chunks, current, line);
}

pub fn emit_php_session_regenerate_id(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    if argc > 0 {
        chunks[current].emit_op(Op::DROP, line);
    }
    vybe_compiler::primitives::http_session::emit_status(chunks, current, line);
    push_const(
        &mut chunks[current],
        Value::I32(vybe_compiler::primitives::http_session::STATUS_ACTIVE),
        line,
    );
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    vybe_compiler::primitives::http_session::emit_regenerate_id(chunks, current, line);
    let ok_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], ok_slot, line);
    emit_php_session_sync_legacy_id(chunks, current, line);
    lget(&mut chunks[current], ok_slot, line);
    chunks[current].emit_else(line);
    push_const(&mut chunks[current], Value::Bool(false), line);
    chunks[current].emit_end(line);
}

pub fn emit_php_session_write_close(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_php_session_commit_working_copy(chunks, current, line);
    emit_php_session_set_status(
        chunks,
        current,
        vybe_compiler::primitives::http_session::STATUS_NONE,
        line,
    );
    push_const(&mut chunks[current], Value::Bool(false), line);
    vybe_compiler::primitives::globals::emit_write(
        &mut chunks[current],
        PHP_SESSION_STARTED_GLOBAL,
        line,
    );
    push_const(&mut chunks[current], Value::Bool(true), line);
}

pub fn emit_php_session_abort(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_php_session_load_working_copy(chunks, current, line);
    emit_php_session_set_status(
        chunks,
        current,
        vybe_compiler::primitives::http_session::STATUS_NONE,
        line,
    );
    push_const(&mut chunks[current], Value::Bool(false), line);
    vybe_compiler::primitives::globals::emit_write(
        &mut chunks[current],
        PHP_SESSION_STARTED_GLOBAL,
        line,
    );
    push_const(&mut chunks[current], Value::Bool(true), line);
}

pub fn emit_php_session_reset(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_php_session_load_working_copy(chunks, current, line);
    emit_php_session_set_status(
        chunks,
        current,
        vybe_compiler::primitives::http_session::STATUS_ACTIVE,
        line,
    );
    push_const(&mut chunks[current], Value::Bool(true), line);
}

pub fn emit_php_session_unset(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    chunks[current].emit_array_new_fixed(0, 0, line);
    vybe_compiler::primitives::globals::emit_write(&mut chunks[current], "$_SESSION", line);
    call_import(chunks, current, "ecma:map", "new", 0, line);
    vybe_compiler::primitives::globals::emit_write(
        &mut chunks[current],
        vybe_compiler::primitives::http_session::SESSION_DATA_GLOBAL,
        line,
    );
    push_const(&mut chunks[current], Value::Bool(true), line);
}

pub fn emit_php_session_destroy(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_php_session_unset(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    push_const(
        chunk,
        Value::I32(vybe_compiler::primitives::http_session::STATUS_NONE),
        line,
    );
    vybe_compiler::primitives::globals::emit_write(
        chunk,
        vybe_compiler::primitives::http_session::SESSION_STATUS_GLOBAL,
        line,
    );

    chunk.emit_op(Op::DROP, line);
    push_const(chunk, Value::Bool(false), line);
    vybe_compiler::primitives::globals::emit_write(chunk, PHP_SESSION_STARTED_GLOBAL, line);
    push_const(chunk, Value::Bool(true), line);
    vybe_compiler::primitives::globals::emit_write(chunk, PHP_SESSION_DESTROYED_GLOBAL, line);
    push_const(chunk, Value::Bool(false), line);
    vybe_compiler::primitives::globals::emit_write(chunk, PHP_SESSION_NEEDS_COOKIE_GLOBAL, line);

    push_str(chunk, "PHPSESSID", line);
    push_str(chunk, "", line);
    let attrs = chunks[current].add_import("ecma:map", "new");
    chunks[current].emit_call(attrs, 0, line);
    let attrs_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], attrs_slot, line);
    lget(&mut chunks[current], attrs_slot, line);
    push_str(&mut chunks[current], "expires", line);
    push_const(&mut chunks[current], Value::F64(1.0), line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], attrs_slot, line);
    vybe_compiler::primitives::http_cookie::emit_serialize(chunks, current, 3, line);
    emit_send_cookie(chunks, current, line);

    push_const(&mut chunks[current], Value::Bool(true), line);
}
