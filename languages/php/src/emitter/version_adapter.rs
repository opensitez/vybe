//! PHP version/runtime-introspection adapters.
//!
//! Replaces the old PHP-source `VERSION_PRELUDE` functions with PHP-local
//! `common:php.*` dispatch and `php.*` namespace-tree leaves.

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

fn string_arg(chunk: &mut Chunk, line: u32) {
    let string = chunk.add_import("ecma:string", "String");
    chunk.emit_call(string, 1, line);
}

fn emit_eq_str(chunk: &mut Chunk, slot: u16, value: &str, line: u32) {
    lget(chunk, slot, line);
    push_str(chunk, value, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
}

fn emit_cmp_op_result(chunk: &mut Chunk, cmp_slot: u16, op_slot: u16, op: &str, line: u32) {
    emit_eq_str(chunk, op_slot, op, line);
    chunk.emit_if_value(line);
    lget(chunk, cmp_slot, line);
    match op {
        "<" | "lt" => {
            chunk.emit_f64_const(0.0, line);
            vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        }
        "<=" | "le" => {
            chunk.emit_f64_const(0.0, line);
            vybe_compiler::primitives::ops::emit_dyn_le(chunk, line);
        }
        ">" | "gt" => {
            chunk.emit_f64_const(0.0, line);
            vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
        }
        ">=" | "ge" => {
            chunk.emit_f64_const(0.0, line);
            vybe_compiler::primitives::ops::emit_dyn_ge(chunk, line);
        }
        "==" | "=" | "eq" => {
            chunk.emit_f64_const(0.0, line);
            vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        }
        "!=" | "<>" | "ne" => {
            chunk.emit_f64_const(0.0, line);
            vybe_compiler::primitives::ops::emit_dyn_ne(chunk, line);
        }
        _ => chunk.emit_bool_const(false, line),
    }
    chunk.emit_else(line);
}

pub fn emit_php_phpversion(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "8.0.0", line);
}

pub fn emit_php_phpinfo(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let report_slot = alloc_local(chunk);
    push_str(
        chunk,
        "phpinfo()\n\
         PHP Version => 8.0.0\n\
         System => Darwin\n\
         Build Date => vybe\n\
         Server API => cli\n\
         PHP API => vybex\n\
         PHP Extension Build => vybe\n\
         Zend Extension Build => n/a\n\
         PHP Integer Size => 8\n",
        line,
    );
    lset(chunk, report_slot, line);
    vybe_compiler::primitives::io::emit_write_stdout_slot(chunk, report_slot, line);
    push_const(chunk, Value::Bool(true), line);
}

pub fn emit_php_get_loaded_extensions(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    for ext in [
        "Core", "standard", "date", "json", "pcre", "SPL", "PDO", "mysqli", "mysqlnd",
        "pdo_mysql",
    ] {
        push_str(chunk, ext, line);
    }
    chunk.emit_array_new_fixed(0, 10, line);
}

pub fn emit_php_extension_loaded(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc == 0 {
        chunk.emit_bool_const(false, line);
        return;
    }
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    string_arg(chunk, line);
    let lower = chunk.add_import("ecma:string", "toLowerCase");
    chunk.emit_call(lower, 1, line);
    let ext_slot = alloc_local(chunk);
    lset(chunk, ext_slot, line);

    for ext in [
        "core", "standard", "date", "json", "pcre", "spl", "pdo", "mysqli", "mysqlnd",
        "pdo_mysql",
    ] {
        emit_eq_str(chunk, ext_slot, ext, line);
        chunk.emit_if_value(line);
        chunk.emit_bool_const(true, line);
        chunk.emit_else(line);
    }
    chunk.emit_bool_const(false, line);
    for _ in 0..10 {
        chunk.emit_end(line);
    }
}

pub fn emit_php_php_uname(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    push_str(chunk, "Darwin vybe", line);
}

pub fn emit_php_version_compare(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let op_slot = alloc_local(chunk);
    if argc >= 3 {
        lset(chunk, op_slot, line);
    }
    let right_slot = alloc_local(chunk);
    let left_slot = alloc_local(chunk);
    lset(chunk, right_slot, line);
    lset(chunk, left_slot, line);
    lget(chunk, left_slot, line);
    string_arg(chunk, line);
    lset(chunk, left_slot, line);
    lget(chunk, right_slot, line);
    string_arg(chunk, line);
    lset(chunk, right_slot, line);

    let cmp_slot = alloc_local(chunk);
    lget(chunk, left_slot, line);
    lget(chunk, right_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_else(line);

    lget(chunk, left_slot, line);
    push_str(chunk, "beta", line);
    let includes = chunk.add_import("ecma:string", "includes");
    chunk.emit_call(includes, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_f64_const(-1.0, line);
    chunk.emit_else(line);

    lget(chunk, left_slot, line);
    lget(chunk, right_slot, line);
    let compare = chunk.add_import("wasm:js-string", "compare");
    chunk.emit_call(compare, 2, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    lset(chunk, cmp_slot, line);

    if argc < 3 {
        lget(chunk, cmp_slot, line);
        return;
    }

    for op in [
        "<", "lt", "<=", "le", ">", "gt", ">=", "ge", "==", "=", "eq", "!=", "<>", "ne",
    ] {
        emit_cmp_op_result(chunk, cmp_slot, op_slot, op, line);
    }
    chunk.emit_bool_const(false, line);
    for _ in 0..14 {
        chunk.emit_end(line);
    }
}
