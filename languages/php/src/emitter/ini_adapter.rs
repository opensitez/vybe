//! PHP INI/runtime-configuration adapters.
//!
//! This replaces the old PHP-source INI prelude. The callable surface is
//! registered as `php.*` namespace tree leaves and dispatched locally as
//! `common:php.*`; state lives in shared globals and values are composed with
//! the existing ECMA/common primitives.

use std::sync::Arc;

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const PHP_INI_ENTRIES: [(&str, &str, &str); 11] = [
    ("display_errors", "__php_ini_display_errors", "1"),
    ("precision", "__php_ini_precision", "14"),
    ("memory_limit", "__php_ini_memory_limit", "128M"),
    ("post_max_size", "__php_ini_post_max_size", "8M"),
    (
        "upload_max_filesize",
        "__php_ini_upload_max_filesize",
        "2M",
    ),
    ("default_charset", "__php_ini_default_charset", "UTF-8"),
    ("error_reporting", "__php_ini_error_reporting", "32767"),
    ("max_execution_time", "__php_ini_max_execution_time", "0"),
    ("include_path", "__php_ini_include_path", ".:/usr/share/php"),
    ("session.save_path", "__php_ini_session_save_path", ""),
    ("opcache.enable", "__php_ini_opcache_enable", "1"),
];

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

fn global_get(chunk: &mut Chunk, key: &str, line: u32) {
    vybe_compiler::primitives::globals::emit_read(chunk, key, line);
}

fn global_set(chunk: &mut Chunk, key: &str, line: u32) {
    vybe_compiler::primitives::globals::emit_write(chunk, key, line);
}

fn emit_is_null_or_undefined(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    lget(chunk, slot, line);
    let undef = chunk.add_import("wasm:js-undefined", "test");
    chunk.emit_call(undef, 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_global_or_default(chunk: &mut Chunk, global: &str, default: &str, line: u32) {
    global_get(chunk, global, line);
    let slot = alloc_local(chunk);
    lset(chunk, slot, line);
    emit_is_null_or_undefined(chunk, slot, line);
    chunk.emit_if_value(line);
    push_str(chunk, default, line);
    chunk.emit_else(line);
    lget(chunk, slot, line);
    chunk.emit_end(line);
}

fn emit_ini_key_chain<F>(chunk: &mut Chunk, name_slot: u16, line: u32, mut emit_known: F)
where
    F: FnMut(&mut Chunk, &str, &str, &str, u32),
{
    for (name, global, default) in PHP_INI_ENTRIES {
        lget(chunk, name_slot, line);
        push_str(chunk, name, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        emit_known(chunk, name, global, default, line);
        chunk.emit_else(line);
    }
    chunk.emit_bool_const(false, line);
    for _ in PHP_INI_ENTRIES {
        chunk.emit_end(line);
    }
}

fn map_set_literal(
    chunks: &mut [Chunk],
    current: usize,
    map_slot: u16,
    key: &str,
    value: &str,
    line: u32,
) {
    lget(&mut chunks[current], map_slot, line);
    push_str(&mut chunks[current], key, line);
    push_str(&mut chunks[current], value, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn map_set_slot_value(
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

fn emit_ini_entry_map(
    chunks: &mut [Chunk],
    current: usize,
    global: &str,
    default: &str,
    line: u32,
) -> u16 {
    call_import(chunks, current, "ecma:map", "new", 0, line);
    let entry_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], entry_slot, line);

    emit_global_or_default(&mut chunks[current], global, default, line);
    let value_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], value_slot, line);

    map_set_literal(chunks, current, entry_slot, "global_value", default, line);
    map_set_slot_value(chunks, current, entry_slot, "local_value", value_slot, line);
    lget(&mut chunks[current], entry_slot, line);
    push_str(&mut chunks[current], "access", line);
    chunks[current].emit_f64_const(7.0, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    entry_slot
}

pub fn emit_php_ini_get(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let name_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], name_slot, line);
    emit_ini_key_chain(
        &mut chunks[current],
        name_slot,
        line,
        |chunk, _name, global, default, line| {
            emit_global_or_default(chunk, global, default, line);
        },
    );
}

pub fn emit_php_ini_set(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let value_slot = alloc_local(&mut chunks[current]);
    let name_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], value_slot, line);
    lget(&mut chunks[current], value_slot, line);
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    lset(&mut chunks[current], value_slot, line);
    lset(&mut chunks[current], name_slot, line);

    emit_ini_key_chain(
        &mut chunks[current],
        name_slot,
        line,
        |chunk, _name, global, default, line| {
            emit_global_or_default(chunk, global, default, line);
            let old_slot = alloc_local(chunk);
            lset(chunk, old_slot, line);
            lget(chunk, value_slot, line);
            global_set(chunk, global, line);
            lget(chunk, old_slot, line);
        },
    );
}

pub fn emit_php_ini_restore(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let name_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], name_slot, line);
    emit_ini_key_chain(
        &mut chunks[current],
        name_slot,
        line,
        |chunk, _name, global, default, line| {
            push_str(chunk, default, line);
            global_set(chunk, global, line);
            chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        },
    );
}

pub fn emit_php_ini_get_all(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc >= 2 {
        chunks[current].emit_op(Op::DROP, line);
    }
    if argc >= 1 {
        chunks[current].emit_op(Op::DROP, line);
    }

    call_import(chunks, current, "ecma:map", "new", 0, line);
    let out_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], out_slot, line);

    for (name, global, default) in PHP_INI_ENTRIES {
        let entry_slot = emit_ini_entry_map(chunks, current, global, default, line);
        lget(&mut chunks[current], out_slot, line);
        push_str(&mut chunks[current], name, line);
        lget(&mut chunks[current], entry_slot, line);
        call_import(chunks, current, "ecma:map", "set", 3, line);
        chunks[current].emit_op(Op::DROP, line);
    }

    lget(&mut chunks[current], out_slot, line);
}

pub fn emit_php_get_cfg_var(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let name_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], name_slot, line);
    lget(&mut chunks[current], name_slot, line);
    push_str(&mut chunks[current], "PHP_VERSION", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    push_str(&mut chunks[current], "8.0.0", line);
    chunks[current].emit_else(line);
    emit_ini_key_chain(
        &mut chunks[current],
        name_slot,
        line,
        |chunk, _name, global, default, line| {
            emit_global_or_default(chunk, global, default, line);
        },
    );
    chunks[current].emit_end(line);
}

pub fn emit_php_get_include_path(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_global_or_default(
        &mut chunks[current],
        "__php_ini_include_path",
        ".:/usr/share/php",
        line,
    );
}

pub fn emit_php_set_include_path(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    let value_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], value_slot, line);
    emit_global_or_default(
        &mut chunks[current],
        "__php_ini_include_path",
        ".:/usr/share/php",
        line,
    );
    let old_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], old_slot, line);
    lget(&mut chunks[current], value_slot, line);
    global_set(&mut chunks[current], "__php_ini_include_path", line);
    lget(&mut chunks[current], old_slot, line);
}
