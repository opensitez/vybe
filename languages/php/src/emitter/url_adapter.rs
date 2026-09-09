//! PHP URL/query adapters.
//!
//! Percent encode/decode itself is shared (`primitives::url`). This file owns
//! PHP's argument order and query-array shape for `http_build_query`.

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

fn bump_index(chunk: &mut Chunk, slot: u16, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_f64_const(1.0, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    lset(chunk, slot, line);
}

fn emit_encoded_slot(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    encoding_slot: u16,
    line: u32,
) {
    lget(&mut chunks[current], encoding_slot, line);
    push_const(&mut chunks[current], Value::F64(2.0), line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], slot, line);
    super::string_adapter::emit_rawurlencode(chunks, current, 1, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], slot, line);
    super::string_adapter::emit_urlencode(chunks, current, 1, line);
    chunks[current].emit_end(line);
}

fn emit_stringify_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) -> u16 {
    lget(&mut chunks[current], slot, line);
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    let out = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], out, line);
    out
}

fn emit_push_query_pair(
    chunks: &mut [Chunk],
    current: usize,
    pairs_slot: u16,
    key_slot: u16,
    value_slot: u16,
    encoding_slot: u16,
    line: u32,
) {
    let value_string_slot = emit_stringify_slot(chunks, current, value_slot, line);
    lget(&mut chunks[current], pairs_slot, line);
    emit_encoded_slot(chunks, current, key_slot, encoding_slot, line);
    push_str(&mut chunks[current], "=", line);
    emit_encoded_slot(chunks, current, value_string_slot, encoding_slot, line);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 3, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_top_key(
    chunks: &mut [Chunk],
    current: usize,
    raw_key_slot: u16,
    numeric_prefix_slot: u16,
    line: u32,
) -> u16 {
    lget(&mut chunks[current], raw_key_slot, line);
    call_import(chunks, current, "wasm:js-number", "test", 1, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], numeric_prefix_slot, line);
    lget(&mut chunks[current], raw_key_slot, line);
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 2, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], raw_key_slot, line);
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    chunks[current].emit_end(line);
    let key_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], key_slot, line);
    key_slot
}

fn emit_child_key(
    chunks: &mut [Chunk],
    current: usize,
    parent_key_slot: u16,
    raw_key_slot: u16,
    line: u32,
) -> u16 {
    lget(&mut chunks[current], parent_key_slot, line);
    push_str(&mut chunks[current], "[", line);
    lget(&mut chunks[current], raw_key_slot, line);
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    push_str(&mut chunks[current], "]", line);
    vybe_compiler::primitives::strings::emit_concat(&mut chunks[current], 4, line);
    let key_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], key_slot, line);
    key_slot
}

fn emit_one_level_children(
    chunks: &mut [Chunk],
    current: usize,
    pairs_slot: u16,
    parent_key_slot: u16,
    value_slot: u16,
    encoding_slot: u16,
    line: u32,
) {
    lget(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::collections::emit_iter_entries(chunks, current, line);
    let entries_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], entries_slot, line);

    lget(&mut chunks[current], entries_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    let len_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], len_slot, line);

    let i_slot = alloc_local(&mut chunks[current]);
    chunks[current].emit_f64_const(0.0, line);
    lset(&mut chunks[current], i_slot, line);

    let pair_slot = alloc_local(&mut chunks[current]);
    let raw_key_slot = alloc_local(&mut chunks[current]);
    let child_value_slot = alloc_local(&mut chunks[current]);

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
    lset(&mut chunks[current], raw_key_slot, line);
    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], child_value_slot, line);

    let child_key_slot = emit_child_key(chunks, current, parent_key_slot, raw_key_slot, line);
    emit_push_query_pair(
        chunks,
        current,
        pairs_slot,
        child_key_slot,
        child_value_slot,
        encoding_slot,
        line,
    );

    bump_index(&mut chunks[current], i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
}

pub fn emit_php_http_build_query(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let encoding_slot = alloc_local(&mut chunks[current]);
    let sep_slot = alloc_local(&mut chunks[current]);
    let numeric_prefix_slot = alloc_local(&mut chunks[current]);
    let data_slot = alloc_local(&mut chunks[current]);

    if argc >= 4 {
        lset(&mut chunks[current], encoding_slot, line);
    } else {
        push_const(&mut chunks[current], Value::F64(1.0), line);
        lset(&mut chunks[current], encoding_slot, line);
    }
    if argc >= 3 {
        lset(&mut chunks[current], sep_slot, line);
    } else {
        push_str(&mut chunks[current], "&", line);
        lset(&mut chunks[current], sep_slot, line);
    }
    if argc >= 2 {
        lset(&mut chunks[current], numeric_prefix_slot, line);
    } else {
        push_str(&mut chunks[current], "", line);
        lset(&mut chunks[current], numeric_prefix_slot, line);
    }
    if argc >= 1 {
        lset(&mut chunks[current], data_slot, line);
    } else {
        call_import(chunks, current, "ecma:map", "new", 0, line);
        lset(&mut chunks[current], data_slot, line);
    }

    chunks[current].emit_array_new_fixed(0, 0, line);
    let pairs_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], pairs_slot, line);

    lget(&mut chunks[current], data_slot, line);
    vybe_compiler::primitives::collections::emit_iter_entries(chunks, current, line);
    let entries_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], entries_slot, line);

    lget(&mut chunks[current], entries_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    let len_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], len_slot, line);

    let i_slot = alloc_local(&mut chunks[current]);
    chunks[current].emit_f64_const(0.0, line);
    lset(&mut chunks[current], i_slot, line);

    let pair_slot = alloc_local(&mut chunks[current]);
    let raw_key_slot = alloc_local(&mut chunks[current]);
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
    lset(&mut chunks[current], raw_key_slot, line);
    lget(&mut chunks[current], pair_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(&mut chunks[current], value_slot, line);

    let key_slot = emit_top_key(chunks, current, raw_key_slot, numeric_prefix_slot, line);

    lget(&mut chunks[current], value_slot, line);
    super::array_adapter::emit_php_is_array(chunks, current, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_one_level_children(
        chunks,
        current,
        pairs_slot,
        key_slot,
        value_slot,
        encoding_slot,
        line,
    );
    chunks[current].emit_else(line);
    emit_push_query_pair(
        chunks,
        current,
        pairs_slot,
        key_slot,
        value_slot,
        encoding_slot,
        line,
    );
    chunks[current].emit_end(line);

    bump_index(&mut chunks[current], i_slot, line);
    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], pairs_slot, line);
    lget(&mut chunks[current], sep_slot, line);
    call_import(chunks, current, "ecma:array", "join", 2, line);
}

pub fn emit_php_parse_str_result(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    if argc == 0 {
        call_import(chunks, current, "ecma:map", "new", 0, line);
        return;
    }
    super::string_adapter::emit_echo_stringify(chunks, current, 1, line);
    vybe_compiler::primitives::http_form::emit_parse_urlencoded(chunks, current, line);
}
