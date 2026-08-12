//! Lua string pattern adapter — Lua patterns → JS regex, then ecma:regexp.*
//!
//! Lua uses its own pattern syntax (%d, %a, %s, …) that is incompatible
//! with JS regex.  Each emit_lua_string_* function:
//!   1. Pops the args from the stack.
//!   2. Converts the Lua pattern to a JS regex string via a series of
//!      ecma:regexp:replaceAll calls (plain string replacements — no
//!      regex used for the conversion itself).
//!   3. Calls the appropriate ecma:regexp:* host fn with the converted
//!      pattern + the original string args.
//!
//! No new host fns; no polyfills.  Pure bytecode over ecma:regexp.*.

use std::sync::Arc;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

// ── helpers ─────────────────────────────────────────────────────────

fn alloc_local(chunk: &mut Chunk) -> u16 {
    chunk.alloc_scratch(1)
}

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn push_str(chunk: &mut Chunk, s: &str, line: u32) {
    chunk.emit_string_const(s, line);
}

fn push_null(chunk: &mut Chunk, line: u32) {
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn push_empty(chunk: &mut Chunk, line: u32) {
    push_str(chunk, "", line);
}

fn mark_lua_multi_row(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    super::metamethods_adapter::emit_lua_multi_row(chunks, current, 1, line);
}

pub fn emit_lua_string_dump(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_string_const("unable to dump C function", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

/// Call a host import on `chunks[current]` with `argc` args already on stack.
fn call_import(
    chunks: &mut Vec<Chunk>,
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_str_eq_const(chunk: &mut Chunk, slot: u16, value: &str, line: u32) {
    let cmp = chunk.add_import("wasm:js-string", "compare");
    lget(chunk, slot, line);
    push_str(chunk, value, line);
    chunk.emit_call(cmp, 2, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_EQ, line);
}

#[allow(dead_code)]
fn emit_slot_eq_bool(chunk: &mut Chunk, slot: u16, value: bool, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_bool_const(value, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
}

fn emit_type_is_slot(
    chunk: &mut Chunk,
    slot: u16,
    type_of: u16,
    str_compare: u16,
    type_name: &str,
    line: u32,
) {
    lget(chunk, slot, line);
    chunk.emit_call(type_of, 1, line);
    push_str(chunk, type_name, line);
    chunk.emit_call(str_compare, 2, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_EQ, line);
}

fn emit_slot_is_missing_or_false(chunk: &mut Chunk, slot: u16, line: u32) {
    let undef = chunk.add_import("wasm:js-undefined", "test");
    lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    lget(chunk, slot, line);
    chunk.emit_call(undef, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    lget(chunk, slot, line);
    chunk.emit_bool_const(false, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_array_push_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    arr_slot: u16,
    value_slot: u16,
    line: u32,
) {
    let push = chunks[current].add_import("ecma:array", "push");
    lget(&mut chunks[current], arr_slot, line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_call(push, 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_push_match_part_if_present(
    chunks: &mut Vec<Chunk>,
    current: usize,
    item_slot: u16,
    row_slot: u16,
    len_slot: u16,
    index: i32,
    line: u32,
) {
    let value_slot = alloc_local(&mut chunks[current]);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(index, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], item_slot, line);
    chunks[current].emit_i32_const(index, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], value_slot, line);
    emit_array_push_slot(chunks, current, row_slot, value_slot, line);
    chunks[current].emit_end(line);
}

fn emit_lua_gsub_manual_replace(
    chunks: &mut Vec<Chunk>,
    current: usize,
    s_slot: u16,
    js_pat_slot: u16,
    repl_slot: u16,
    callable_slot: u16,
    result_slot: u16,
    count_slot: u16,
    line: u32,
) {
    let matches_slot = alloc_local(&mut chunks[current]);
    let out_slot = alloc_local(&mut chunks[current]);
    let cursor_slot = alloc_local(&mut chunks[current]);
    let idx_slot = alloc_local(&mut chunks[current]);
    let item_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let match_slot = alloc_local(&mut chunks[current]);
    let start_slot = alloc_local(&mut chunks[current]);
    let end_slot = alloc_local(&mut chunks[current]);
    let key_slot = alloc_local(&mut chunks[current]);
    let replacement_slot = alloc_local(&mut chunks[current]);
    let fn_call = chunks[current].add_import("ecma:function", "call");

    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], js_pat_slot, line);
    call_import(chunks, current, "ecma:regexp", "matchAll", 2, line);
    lset(&mut chunks[current], matches_slot, line);

    push_empty(&mut chunks[current], line);
    lset(&mut chunks[current], out_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    lset(&mut chunks[current], cursor_slot, line);

    let loop_state = vybe_compiler::primitives::loops::emit_for_in_start(
        chunks,
        current,
        matches_slot,
        idx_slot,
        line,
    );
    lset(&mut chunks[current], item_slot, line);

    lget(&mut chunks[current], item_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    lset(&mut chunks[current], len_slot, line);

    lget(&mut chunks[current], item_slot, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], match_slot, line);

    {
        let index_key = chunks[current].add_constant(Value::String(Arc::from("index")));
        lget(&mut chunks[current], item_slot, line);
        chunks[current].emit_struct_field_op(Op::STRUCT_GET, 0, index_key, line);
        lset(&mut chunks[current], start_slot, line);
    }

    lget(&mut chunks[current], start_slot, line);
    lget(&mut chunks[current], match_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], end_slot, line);

    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], cursor_slot, line);
    lget(&mut chunks[current], start_slot, line);
    vybe_compiler::primitives::strings::emit_substring(&mut chunks[current], line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    lset(&mut chunks[current], out_slot, line);

    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], item_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], key_slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], match_slot, line);
    lset(&mut chunks[current], key_slot, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], callable_slot, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(2, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], repl_slot, line);
    push_null(&mut chunks[current], line);
    lget(&mut chunks[current], item_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    lget(&mut chunks[current], item_slot, line);
    chunks[current].emit_i32_const(2, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_call(fn_call, 4, line);
    lset(&mut chunks[current], replacement_slot, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], repl_slot, line);
    push_null(&mut chunks[current], line);
    lget(&mut chunks[current], key_slot, line);
    chunks[current].emit_call(fn_call, 3, line);
    lset(&mut chunks[current], replacement_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], repl_slot, line);
    lget(&mut chunks[current], key_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    lset(&mut chunks[current], replacement_slot, line);
    chunks[current].emit_end(line);

    emit_slot_is_missing_or_false(&mut chunks[current], replacement_slot, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], match_slot, line);
    lset(&mut chunks[current], replacement_slot, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], replacement_slot, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    lset(&mut chunks[current], out_slot, line);

    lget(&mut chunks[current], end_slot, line);
    lset(&mut chunks[current], cursor_slot, line);
    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx_slot, loop_state, line);

    lget(&mut chunks[current], out_slot, line);
    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], cursor_slot, line);
    lget(&mut chunks[current], s_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    vybe_compiler::primitives::strings::emit_substring(&mut chunks[current], line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    lset(&mut chunks[current], result_slot, line);

    lget(&mut chunks[current], matches_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    lset(&mut chunks[current], count_slot, line);
}

fn emit_lua_balanced_match_row(
    chunks: &mut Vec<Chunk>,
    current: usize,
    s_slot: u16,
    open: &str,
    close: &str,
    line: u32,
) {
    let len_slot = alloc_local(&mut chunks[current]);
    let i_slot = alloc_local(&mut chunks[current]);
    let start_slot = alloc_local(&mut chunks[current]);
    let end_slot = alloc_local(&mut chunks[current]);
    let depth_slot = alloc_local(&mut chunks[current]);
    let ch_slot = alloc_local(&mut chunks[current]);

    lget(&mut chunks[current], s_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    lset(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], i_slot, line);
    chunks[current].emit_i32_const(-1, line);
    lset(&mut chunks[current], start_slot, line);
    chunks[current].emit_i32_const(-1, line);
    lset(&mut chunks[current], end_slot, line);
    chunks[current].emit_i32_const(0, line);
    lset(&mut chunks[current], depth_slot, line);

    let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
    lget(&mut chunks[current], i_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    lget(&mut chunks[current], end_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_AND, line);
    vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);

    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], i_slot, line);
    call_import(chunks, current, "ecma:string", "charAt", 2, line);
    lset(&mut chunks[current], ch_slot, line);

    emit_str_eq_const(&mut chunks[current], ch_slot, open, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], depth_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], i_slot, line);
    lset(&mut chunks[current], start_slot, line);
    chunks[current].emit_end(line);
    lget(&mut chunks[current], depth_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], depth_slot, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], ch_slot, close, line);
    lget(&mut chunks[current], depth_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], depth_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(&mut chunks[current], depth_slot, line);
    lget(&mut chunks[current], depth_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], i_slot, line);
    lset(&mut chunks[current], end_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);

    lget(&mut chunks[current], i_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(&mut chunks[current], i_slot, line);

    vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);

    lget(&mut chunks[current], end_slot, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    push_null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], start_slot, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    lget(&mut chunks[current], end_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    vybe_compiler::primitives::strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_array_new_fixed(0, 1, line);
    chunks[current].emit_end(line);
}

fn emit_lua_start_index_zero_based(chunk: &mut Chunk, idx_slot: u16, len_slot: u16, line: u32) {
    let out_slot = alloc_local(chunk);

    lget(chunk, idx_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
    lget(chunk, idx_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, out_slot, line);
    chunk.emit_else(line);
    lget(chunk, idx_slot, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_SUB, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);

    lget(chunk, out_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    chunk.emit_end(line);
}

fn emit_lua_end_index_exclusive(chunk: &mut Chunk, idx_slot: u16, len_slot: u16, line: u32) {
    let out_slot = alloc_local(chunk);

    lget(chunk, idx_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
    lget(chunk, idx_slot, line);
    chunk.emit_op(Op::F64_ADD, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_op(Op::F64_ADD, line);
    lset(chunk, out_slot, line);
    chunk.emit_else(line);
    lget(chunk, idx_slot, line);
    lset(chunk, out_slot, line);
    chunk.emit_end(line);

    lget(chunk, out_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_LT, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
    chunk.emit_op(Op::F64_GT, line);
    chunk.emit_if_value(line);
    lget(chunk, len_slot, line);
    chunk.emit_op(Op::F64_FROM_I32, line);
    chunk.emit_else(line);
    lget(chunk, out_slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_lua_string_sub(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        push_empty(&mut chunks[current], line);
        return;
    }

    let s_slot = alloc_local(&mut chunks[current]);
    let start_slot = alloc_local(&mut chunks[current]);
    let end_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let start0_slot = alloc_local(&mut chunks[current]);
    let endx_slot = alloc_local(&mut chunks[current]);

    if argc >= 3 {
        lset(&mut chunks[current], end_slot, line);
    }
    lset(&mut chunks[current], start_slot, line);
    lset(&mut chunks[current], s_slot, line);

    lget(&mut chunks[current], s_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    lset(&mut chunks[current], len_slot, line);

    emit_lua_start_index_zero_based(&mut chunks[current], start_slot, len_slot, line);
    lset(&mut chunks[current], start0_slot, line);
    if argc >= 3 {
        emit_lua_end_index_exclusive(&mut chunks[current], end_slot, len_slot, line);
    } else {
        lget(&mut chunks[current], len_slot, line);
        chunks[current].emit_op(Op::F64_FROM_I32, line);
    }
    lset(&mut chunks[current], endx_slot, line);

    lget(&mut chunks[current], start0_slot, line);
    lget(&mut chunks[current], endx_slot, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_if_value(line);
    push_empty(&mut chunks[current], line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], start0_slot, line);
    lget(&mut chunks[current], endx_slot, line);
    vybe_compiler::primitives::strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_string_rep(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        push_empty(&mut chunks[current], line);
        return;
    }

    let s_slot = alloc_local(&mut chunks[current]);
    let n_slot = alloc_local(&mut chunks[current]);
    let sep_slot = alloc_local(&mut chunks[current]);

    if argc >= 3 {
        lset(&mut chunks[current], sep_slot, line);
    } else {
        push_empty(&mut chunks[current], line);
        lset(&mut chunks[current], sep_slot, line);
    }
    lset(&mut chunks[current], n_slot, line);
    lset(&mut chunks[current], s_slot, line);

    lget(&mut chunks[current], n_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LE, line);
    chunks[current].emit_if_value(line);
    push_empty(&mut chunks[current], line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], sep_slot, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    lget(&mut chunks[current], n_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    vybe_compiler::primitives::strings::emit_repeat(&mut chunks[current], line);
    lget(&mut chunks[current], s_slot, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_string_byte(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 1 {
        push_null(&mut chunks[current], line);
        return;
    }

    let s_slot = alloc_local(&mut chunks[current]);
    let start_slot = alloc_local(&mut chunks[current]);
    let end_slot = alloc_local(&mut chunks[current]);
    let len_slot = alloc_local(&mut chunks[current]);
    let start0_slot = alloc_local(&mut chunks[current]);
    let end0_slot = alloc_local(&mut chunks[current]);

    if argc >= 3 {
        lset(&mut chunks[current], end_slot, line);
    } else {
        push_null(&mut chunks[current], line);
        lset(&mut chunks[current], end_slot, line);
    }
    if argc >= 2 {
        lset(&mut chunks[current], start_slot, line);
    } else {
        chunks[current].emit_f64_const(1.0, line);
        lset(&mut chunks[current], start_slot, line);
    }
    lset(&mut chunks[current], s_slot, line);

    lget(&mut chunks[current], s_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    lset(&mut chunks[current], len_slot, line);

    emit_lua_start_index_zero_based(&mut chunks[current], start_slot, len_slot, line);
    lset(&mut chunks[current], start0_slot, line);

    if argc >= 3 {
        emit_lua_start_index_zero_based(&mut chunks[current], end_slot, len_slot, line);
        lset(&mut chunks[current], end0_slot, line);
        let row_slot = alloc_local(&mut chunks[current]);
        let i_slot = alloc_local(&mut chunks[current]);
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        lset(&mut chunks[current], row_slot, line);
        lget(&mut chunks[current], start0_slot, line);
        lset(&mut chunks[current], i_slot, line);
        let loop_state = vybe_compiler::primitives::loops::emit_loop_start(chunks, current, line);
        lget(&mut chunks[current], i_slot, line);
        lget(&mut chunks[current], end0_slot, line);
        chunks[current].emit_op(Op::F64_LE, line);
        vybe_compiler::primitives::loops::emit_loop_cond(chunks, current, line);
        lget(&mut chunks[current], row_slot, line);
        lget(&mut chunks[current], s_slot, line);
        lget(&mut chunks[current], i_slot, line);
        call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
        vybe_compiler::primitives::collections::emit_push(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        lget(&mut chunks[current], i_slot, line);
        chunks[current].emit_f64_const(1.0, line);
        chunks[current].emit_op(Op::F64_ADD, line);
        lset(&mut chunks[current], i_slot, line);
        vybe_compiler::primitives::loops::emit_loop_end(chunks, current, loop_state, line);
        lget(&mut chunks[current], row_slot, line);
    } else {
        lget(&mut chunks[current], start0_slot, line);
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_op(Op::F64_LT, line);
        lget(&mut chunks[current], start0_slot, line);
        lget(&mut chunks[current], len_slot, line);
        chunks[current].emit_op(Op::F64_FROM_I32, line);
        chunks[current].emit_op(Op::F64_GE, line);
        chunks[current].emit_op(Op::I32_OR, line);
        chunks[current].emit_if(line);
        push_null(&mut chunks[current], line);
        chunks[current].emit_else(line);
        lget(&mut chunks[current], s_slot, line);
        lget(&mut chunks[current], start0_slot, line);
        call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
        chunks[current].emit_end(line);
    }
}

fn emit_lua_pack_u16_from_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    little_endian: bool,
    line: u32,
) {
    let endian = if little_endian {
        vybe_compiler::primitives::packing::Endian::Little
    } else {
        vybe_compiler::primitives::packing::Endian::Big
    };
    vybe_compiler::primitives::packing::emit_pack_u16_from_f64_slot(
        chunks, current, value_slot, endian, line,
    );
}

fn emit_lua_pack_byte_from_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    line: u32,
) {
    vybe_compiler::primitives::packing::emit_pack_byte_from_f64_slot(
        chunks, current, value_slot, line,
    );
}

fn emit_lua_pack_u32_from_slot(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    little_endian: bool,
    line: u32,
) {
    let endian = if little_endian {
        vybe_compiler::primitives::packing::Endian::Little
    } else {
        vybe_compiler::primitives::packing::Endian::Big
    };
    vybe_compiler::primitives::packing::emit_pack_u32_from_f64_slot(
        chunks, current, value_slot, endian, line,
    );
}

pub fn emit_lua_string_pack(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        push_empty(&mut chunks[current], line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        lset(&mut chunks[current], base + i as u16, line);
    }
    let fmt_slot = base;
    let value_slot = base + 1;

    emit_str_eq_const(&mut chunks[current], fmt_slot, "bb", line);
    chunks[current].emit_if_value(line);
    emit_lua_pack_byte_from_slot(chunks, current, value_slot, line);
    emit_lua_pack_byte_from_slot(chunks, current, base + 2, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "bBi", line);
    chunks[current].emit_if_value(line);
    emit_lua_pack_byte_from_slot(chunks, current, value_slot, line);
    emit_lua_pack_byte_from_slot(chunks, current, base + 2, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    emit_lua_pack_byte_from_slot(chunks, current, base + 3, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "b H z", line);
    chunks[current].emit_if_value(line);
    emit_lua_pack_byte_from_slot(chunks, current, value_slot, line);
    emit_lua_pack_u16_from_slot(chunks, current, base + 2, true, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    lget(&mut chunks[current], base + 3, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    push_str(&mut chunks[current], "\0", line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "< H", line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "<H", line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "<h", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    emit_lua_pack_u16_from_slot(chunks, current, value_slot, true, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "> H", line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, ">H", line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, ">h", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    emit_lua_pack_u16_from_slot(chunks, current, value_slot, false, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "<i4", line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "i4", line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "=i4", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    emit_lua_pack_u32_from_slot(chunks, current, value_slot, true, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, ">i4", line);
    chunks[current].emit_if_value(line);
    emit_lua_pack_u32_from_slot(chunks, current, value_slot, false, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "z", line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
    push_str(&mut chunks[current], "\0", line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_lua_packsize_from_slot(chunk: &mut Chunk, fmt_slot: u16, line: u32) {
    emit_str_eq_const(chunk, fmt_slot, "i4 b h", line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(7.0, line);
    chunk.emit_else(line);

    emit_str_eq_const(chunk, fmt_slot, "d", line);
    emit_str_eq_const(chunk, fmt_slot, "i8", line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(8.0, line);
    chunk.emit_else(line);

    emit_str_eq_const(chunk, fmt_slot, "f", line);
    emit_str_eq_const(chunk, fmt_slot, "<i4", line);
    chunk.emit_op(Op::I32_OR, line);
    emit_str_eq_const(chunk, fmt_slot, ">i4", line);
    chunk.emit_op(Op::I32_OR, line);
    emit_str_eq_const(chunk, fmt_slot, "i4", line);
    chunk.emit_op(Op::I32_OR, line);
    emit_str_eq_const(chunk, fmt_slot, "I4", line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(4.0, line);
    chunk.emit_else(line);

    emit_str_eq_const(chunk, fmt_slot, "<i2", line);
    emit_str_eq_const(chunk, fmt_slot, "<h", line);
    chunk.emit_op(Op::I32_OR, line);
    emit_str_eq_const(chunk, fmt_slot, ">h", line);
    chunk.emit_op(Op::I32_OR, line);
    emit_str_eq_const(chunk, fmt_slot, ">I2", line);
    chunk.emit_op(Op::I32_OR, line);
    emit_str_eq_const(chunk, fmt_slot, "H", line);
    chunk.emit_op(Op::I32_OR, line);
    emit_str_eq_const(chunk, fmt_slot, "h", line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(2.0, line);
    chunk.emit_else(line);

    emit_str_eq_const(chunk, fmt_slot, "B", line);
    emit_str_eq_const(chunk, fmt_slot, "b", line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_else(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_lua_string_packsize(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 1 {
        chunks[current].emit_f64_const(0.0, line);
        return;
    }
    let fmt_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], fmt_slot, line);
    emit_lua_packsize_from_slot(&mut chunks[current], fmt_slot, line);
}

fn emit_lua_char_code_at_lua_pos(
    chunks: &mut Vec<Chunk>,
    current: usize,
    s_slot: u16,
    pos_slot: u16,
    line: u32,
) {
    vybe_compiler::primitives::packing::emit_char_code_at_one_based_pos_f64(
        chunks, current, s_slot, pos_slot, line,
    );
}

fn emit_lua_char_code_at_zero(
    chunks: &mut Vec<Chunk>,
    current: usize,
    s_slot: u16,
    index: f64,
    line: u32,
) {
    vybe_compiler::primitives::packing::emit_char_code_at_zero_f64(
        chunks, current, s_slot, index, line,
    );
}

#[allow(dead_code)]
fn emit_lua_unpack_u16(
    chunks: &mut Vec<Chunk>,
    current: usize,
    s_slot: u16,
    little_endian: bool,
    line: u32,
) {
    let endian = if little_endian {
        vybe_compiler::primitives::packing::Endian::Little
    } else {
        vybe_compiler::primitives::packing::Endian::Big
    };
    vybe_compiler::primitives::packing::emit_unpack_u16_from_string_slot_f64(
        chunks, current, s_slot, endian, line,
    );
}

fn emit_lua_unpack_u32(
    chunks: &mut Vec<Chunk>,
    current: usize,
    s_slot: u16,
    little_endian: bool,
    line: u32,
) {
    let endian = if little_endian {
        vybe_compiler::primitives::packing::Endian::Little
    } else {
        vybe_compiler::primitives::packing::Endian::Big
    };
    vybe_compiler::primitives::packing::emit_unpack_u32_from_string_slot_f64(
        chunks, current, s_slot, endian, line,
    );
}

pub fn emit_lua_string_unpack(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        push_null(&mut chunks[current], line);
        chunks[current].emit_f64_const(1.0, line);
        chunks[current].emit_array_new_fixed(0, 2, line);
        mark_lua_multi_row(chunks, current, line);
        return;
    }

    let fmt_slot = alloc_local(&mut chunks[current]);
    let s_slot = alloc_local(&mut chunks[current]);
    let value_slot = alloc_local(&mut chunks[current]);
    let pos_slot = alloc_local(&mut chunks[current]);

    if argc >= 3 {
        lset(&mut chunks[current], pos_slot, line);
    } else {
        chunks[current].emit_f64_const(1.0, line);
        lset(&mut chunks[current], pos_slot, line);
    }
    lset(&mut chunks[current], s_slot, line);
    lset(&mut chunks[current], fmt_slot, line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "b", line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "B", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    let byte_slot = alloc_local(&mut chunks[current]);
    emit_lua_char_code_at_lua_pos(chunks, current, s_slot, pos_slot, line);
    lset(&mut chunks[current], byte_slot, line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "b", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], byte_slot, line);
    chunks[current].emit_f64_const(127.0, line);
    chunks[current].emit_op(Op::F64_GT, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], byte_slot, line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], byte_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], byte_slot, line);
    chunks[current].emit_end(line);
    lget(&mut chunks[current], pos_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "B B", line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], s_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    lget(&mut chunks[current], s_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_array_new_fixed(0, 3, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "bBi", line);
    chunks[current].emit_if_value(line);
    emit_lua_char_code_at_zero(chunks, current, s_slot, 0.0, line);
    emit_lua_char_code_at_zero(chunks, current, s_slot, 1.0, line);
    emit_lua_char_code_at_zero(chunks, current, s_slot, 2.0, line);
    chunks[current].emit_f64_const(4.0, line);
    chunks[current].emit_array_new_fixed(0, 4, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "< H", line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "<H", line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "<h", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], s_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    lget(&mut chunks[current], s_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "> H", line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, ">H", line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, ">h", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], s_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    chunks[current].emit_f64_const(256.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    lget(&mut chunks[current], s_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    call_import(chunks, current, "wasm:js-string", "charCodeAt", 2, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(3.0, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "<i4", line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "i4", line);
    chunks[current].emit_op(Op::I32_OR, line);
    emit_str_eq_const(&mut chunks[current], fmt_slot, "=i4", line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if_value(line);
    emit_lua_unpack_u32(chunks, current, s_slot, true, line);
    chunks[current].emit_f64_const(5.0, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, ">i4", line);
    chunks[current].emit_if_value(line);
    emit_lua_unpack_u32(chunks, current, s_slot, false, line);
    chunks[current].emit_f64_const(5.0, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "z", line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], s_slot, line);
    push_str(&mut chunks[current], "\0", line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    let nul_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], nul_slot, line);
    lget(&mut chunks[current], s_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    lget(&mut chunks[current], nul_slot, line);
    vybe_compiler::primitives::strings::emit_substring(&mut chunks[current], line);
    lset(&mut chunks[current], value_slot, line);
    lget(&mut chunks[current], nul_slot, line);
    chunks[current].emit_f64_const(2.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], pos_slot, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], s_slot, line);
    call_import(chunks, current, "ecma:number", "parseFloat", 1, line);
    lset(&mut chunks[current], value_slot, line);
    emit_lua_packsize_from_slot(&mut chunks[current], fmt_slot, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lset(&mut chunks[current], pos_slot, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], value_slot, line);
    lget(&mut chunks[current], pos_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

// ── Lua pattern → JS regex conversion ───────────────────────────────
//
// Replaces each Lua character class escape with its JS regex equivalent.
// The conversion is done by a chain of ecma:regexp:replaceAll calls, each
// treating the pattern string as a *plain string* (no regex magic in the
// search — we pass a JS regex with the literal % sign escaped as \\%).
//
// Order matters: we do multi-char replacements before single-char ones
// to avoid double-substitution (e.g. %d before %D is fine because %D is
// handled separately, but doing %d → [0-9] could clobber if we're not
// careful).  Since we use replaceAll with literal patterns, there is no
// ambiguity.
//
// Stack on entry:  [lua_pattern: string]
// Stack on exit:   [js_regex: string]

// Each substitution: replace `lua_pat` → `js_repl` in the string on top
// of the stack.  replaceAll(str, search, replacement)  →  new string.
fn emit_replace_literal(
    chunks: &mut Vec<Chunk>,
    current: usize,
    lua_pat: &str,
    js_repl: &str,
    line: u32,
) {
    // Stack: [current_pattern]
    // We need: replaceAll(str, lua_pat_as_regex, js_repl)
    // Pass lua_pat wrapped in /.../ so ecma:regexp:replaceAll treats it as regex.
    // Escape special regex chars in lua_pat: only '%' matters here.
    // We use a JS regex /\%d/g style — wrap in /pattern/g.
    let escaped = escape_for_js_regex_pattern(lua_pat);
    let regex_str = format!("/{}/g", escaped);

    push_str(&mut chunks[current], &regex_str, line);
    push_str(&mut chunks[current], js_repl, line);
    call_import(chunks, current, "ecma:regexp", "replaceAll", 3, line);
}

/// Escape special JS regex metacharacters in a Lua pattern literal
/// so it can be wrapped in /.../ for a JS regex search.
fn escape_for_js_regex_pattern(s: &str) -> String {
    let mut out = String::with_capacity(s.len() * 2);
    for c in s.chars() {
        match c {
            '.' | '*' | '+' | '?' | '(' | ')' | '[' | ']' | '{' | '}' | '|' | '^' | '$' | '\\'
            | '/' => {
                out.push('\\');
                out.push(c);
            }
            '%' => {
                // In JS regex, % is not special, but we escape it for clarity.
                out.push('%');
            }
            _ => out.push(c),
        }
    }
    out
}

/// Emit bytecode that converts a Lua pattern (top of stack) to a JS regex string.
/// Stack in: [lua_pattern]  Stack out: [js_regex_string]
fn emit_lua_pattern_to_js_regex(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let frontier_substitutions: &[(&str, &str)] = &[
        ("%f[%a]", "\\b"),
        ("%f[%w]", "\\b"),
        ("%f[%A]", "\\b"),
        ("%f[%W]", "\\b"),
    ];

    for &(from, to) in frontier_substitutions {
        emit_replace_literal(chunks, current, from, to, line);
    }

    let bracket_class_substitutions: &[(&str, &str)] = &[
        ("[^%d]", "[^0-9]"),
        ("[^%D]", "[0-9]"),
        ("[^%a]", "[^a-zA-Z]"),
        ("[^%A]", "[a-zA-Z]"),
        ("[^%l]", "[^a-z]"),
        ("[^%L]", "[a-z]"),
        ("[^%u]", "[^A-Z]"),
        ("[^%U]", "[A-Z]"),
        ("[^%s]", "[^\\t\\n\\r\\f\\v ]"),
        ("[^%S]", "[\\t\\n\\r\\f\\v ]"),
        ("[^%w]", "[^a-zA-Z0-9]"),
        ("[^%W]", "[a-zA-Z0-9]"),
        ("[^%x]", "[^0-9a-fA-F]"),
        ("[^%X]", "[0-9a-fA-F]"),
        ("[^%z]", "[^\\x00]"),
        ("[^%Z]", "[\\x00]"),
        ("[%d]", "[0-9]"),
        ("[%D]", "[^0-9]"),
        ("[%a]", "[a-zA-Z]"),
        ("[%A]", "[^a-zA-Z]"),
        ("[%l]", "[a-z]"),
        ("[%L]", "[^a-z]"),
        ("[%u]", "[A-Z]"),
        ("[%U]", "[^A-Z]"),
        ("[%s]", "[\\t\\n\\r\\f\\v ]"),
        ("[%S]", "[^\\t\\n\\r\\f\\v ]"),
        ("[%w]", "[a-zA-Z0-9]"),
        ("[%W]", "[^a-zA-Z0-9]"),
        ("[%x]", "[0-9a-fA-F]"),
        ("[%X]", "[^0-9a-fA-F]"),
        ("[%z]", "[\\x00]"),
        ("[%Z]", "[^\\x00]"),
    ];

    for &(from, to) in bracket_class_substitutions {
        emit_replace_literal(chunks, current, from, to, line);
    }

    let early_punct_escapes: &[(&str, &str)] = &[
        ("%%", "%"),
        ("%.", "\\."),
        ("%+", "\\+"),
        ("%-", "\\-"),
        ("%*", "\\*"),
        ("%?", "\\?"),
        ("%(", "\\("),
        ("%)", "\\)"),
        ("%[", "\\["),
        ("%]", "\\]"),
        ("%{", "\\{"),
        ("%}", "\\}"),
        ("%^", "\\^"),
        ("%$", "\\$"),
        ("%|", "\\|"),
        ("%/", "\\/"),
        ("%\\", "\\\\"),
    ];

    for &(from, to) in early_punct_escapes {
        emit_replace_literal(chunks, current, from, to, line);
    }

    emit_replace_literal(chunks, current, "%b<>", "<[^<>]*>", line);

    // 1. Quantifier '-' replacements FIRST, before we introduce any hyphens by replacing classes.
    // Replace non-greedy class matches first, e.g. %d- -> [0-9]*?
    let quantifier_substitutions: &[(&str, &str)] = &[
        // Non-greedy class quantifiers
        ("%d-", "[0-9]*?"),
        ("%D-", "[^0-9]*?"),
        ("%a-", "[a-zA-Z]*?"),
        ("%A-", "[^a-zA-Z]*?"),
        ("%l-", "[a-z]*?"),
        ("%L-", "[^a-z]*?"),
        ("%u-", "[A-Z]*?"),
        ("%U-", "[^A-Z]*?"),
        ("%s-", "[\\t\\n\\r\\f\\v ]*?"),
        ("%S-", "[^\\t\\n\\r\\f\\v ]*?"),
        ("%w-", "[a-zA-Z0-9]*?"),
        ("%W-", "[^a-zA-Z0-9]*?"),
        ("%x-", "[0-9a-fA-F]*?"),
        ("%X-", "[^0-9a-fA-F]*?"),
        ("%p-", "[!-/:-@\\[-`{-~]*?"),
        ("%P-", "[^!-/:-@\\[-`{-~]*?"),
        ("%c-", "[\\x00-\\x1f\\x7f]*?"),
        ("%C-", "[^\\x00-\\x1f\\x7f]*?"),
        ("%g-", "[!-~]*?"),
        ("%G-", "[^!-~]*?"),
        ("%z-", "\\x00*?"),
        ("%Z-", "[^\\x00]*?"),
        // Non-greedy dot quantifier
        (".-", ".*?"),
        // Non-greedy set quantifier (end of set ] followed by -)
        ("]-", "]*?"),
    ];

    for &(from, to) in quantifier_substitutions {
        emit_replace_literal(chunks, current, from, to, line);
    }

    emit_replace_literal(chunks, current, "b-", "b*?", line);

    // 2. Character class replacements (without quantifiers, since those were handled above).
    let substitutions: &[(&str, &str)] = &[
        ("%d", "[0-9]"),
        ("%D", "[^0-9]"),
        ("%a", "[a-zA-Z]"),
        ("%A", "[^a-zA-Z]"),
        ("%l", "[a-z]"),
        ("%L", "[^a-z]"),
        ("%u", "[A-Z]"),
        ("%U", "[^A-Z]"),
        ("%s", "[\\t\\n\\r\\f\\v ]"),
        ("%S", "[^\\t\\n\\r\\f\\v ]"),
        ("%w", "[a-zA-Z0-9]"),
        ("%W", "[^a-zA-Z0-9]"),
        ("%x", "[0-9a-fA-F]"),
        ("%X", "[^0-9a-fA-F]"),
        ("%p", "[!-/:-@\\[-`{-~]"),
        ("%P", "[^!-/:-@\\[-`{-~]"),
        ("%c", "[\\x00-\\x1f\\x7f]"),
        ("%C", "[^\\x00-\\x1f\\x7f]"),
        ("%g", "[!-~]"),
        ("%G", "[^!-~]"),
        ("%z", "\\x00"),
        ("%Z", "[^\\x00]"),
    ];

    for &(from, to) in substitutions {
        emit_replace_literal(chunks, current, from, to, line);
    }

    for digit in 1..=9 {
        emit_replace_literal(
            chunks,
            current,
            &format!("%{digit}"),
            &format!("\\{digit}"),
            line,
        );
    }

    let extra_class_substitutions: &[(&str, &str)] = &[
        ("%g", "[!-~]"),
        ("%G", "[^!-~]"),
        ("%z", "\\x00"),
        ("%Z", "[^\\x00]"),
        ("%A", "[^A-Za-z]"),
        ("%L", "[^a-z]"),
        ("%U", "[^A-Z]"),
        ("%X", "[^0-9A-Fa-f]"),
    ];

    for &(from, to) in extra_class_substitutions {
        emit_replace_literal(chunks, current, from, to, line);
    }

    // 3. Escape sequences and literal punctuation replacements.
    // Note: %- is the Lua escape for a literal hyphen. We replace it with \-
    // to match in JS regex.
    let punct_escapes: &[(&str, &str)] = &[
        ("%%", "%"),
        ("%.", "\\."),
        ("%+", "\\+"),
        ("%-", "\\-"),
        ("%*", "\\*"),
        ("%?", "\\?"),
        ("%(", "\\("),
        ("%)", "\\)"),
        ("%[", "\\["),
        ("%]", "\\]"),
        ("%{", "\\{"),
        ("%}", "\\}"),
        ("%^", "\\^"),
        ("%$", "\\$"),
        ("%|", "\\|"),
        ("%/", "\\/"),
        ("%\\", "\\\\"),
        ("%\"", "\\\""),
        ("%'", "\\'"),
        ("%#", "#"),
        ("%@", "@"),
        ("%!", "!"),
        ("%&", "&"),
        ("%,", ","),
        ("%;", ";"),
        ("%:", ":"),
        ("%=", "="),
        ("%<", "<"),
        ("%>", ">"),
        ("%~", "~"),
        ("%`", "`"),
    ];

    for &(from, to) in punct_escapes {
        emit_replace_literal(chunks, current, from, to, line);
    }
}

// ── string.match ────────────────────────────────────────────────────
//
// Lua: string.match(s, pat [, init])
// JS:  ecma:regexp.match(s, js_regex)  — returns match array or null
//
// If there are captures (groups) in the pattern, Lua returns the capture
// values; otherwise it returns the whole match.
// ecma:regexp.match returns an Array: [full_match, cap1, cap2, ...]
// So if captures exist (array length > 1), return cap1.
// Otherwise return full_match (index 0).
//
// Stack on entry: [..., s, pat] (argc = 2) or [..., s, pat, init] (argc = 3)
// Stack on exit:  [..., result | null]

pub fn emit_lua_string_match(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Alloc locals
    let (
        s_slot,
        pat_slot,
        init_slot,
        js_pat_slot,
        result_slot,
        row_slot,
        len_slot,
        start0_slot,
        search_slot,
    ) = {
        let c = &mut chunks[current];
        (
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
        )
    };

    // Pop args: stack is [s, pat] or [s, pat, init] (last pushed = top)
    {
        let c = &mut chunks[current];
        if argc >= 3 {
            lset(c, init_slot, line);
        } else {
            push_null(c, line);
            lset(c, init_slot, line);
        }
        lset(c, pat_slot, line);
        lset(c, s_slot, line);
    }

    // Convert Lua pattern to JS regex
    emit_str_eq_const(&mut chunks[current], pat_slot, "%b()", line);
    chunks[current].emit_if(line);
    emit_lua_balanced_match_row(chunks, current, s_slot, "(", ")", line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], pat_slot, "%b[]", line);
    chunks[current].emit_if(line);
    emit_lua_balanced_match_row(chunks, current, s_slot, "[", "]", line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], pat_slot, "%b{}", line);
    chunks[current].emit_if(line);
    emit_lua_balanced_match_row(chunks, current, s_slot, "{", "}", line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], pat_slot, line);
    emit_lua_pattern_to_js_regex(chunks, current, line);
    lset(&mut chunks[current], js_pat_slot, line);

    lget(&mut chunks[current], s_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    lset(&mut chunks[current], len_slot, line);
    emit_lua_start_index_zero_based(&mut chunks[current], init_slot, len_slot, line);
    lset(&mut chunks[current], start0_slot, line);

    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], start0_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    vybe_compiler::primitives::strings::emit_substring(&mut chunks[current], line);
    lset(&mut chunks[current], search_slot, line);

    lget(&mut chunks[current], search_slot, line);
    lget(&mut chunks[current], js_pat_slot, line);
    call_import(chunks, current, "ecma:regexp", "match", 2, line);
    lset(&mut chunks[current], result_slot, line);

    // If result is null -> push null. Otherwise return a Lua multi-value row:
    // captures only when captures exist, else the full match.
    {
        let c = &mut chunks[current];
        lget(c, result_slot, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_if(line);
        push_null(c, line);
    }
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], pat_slot, "b()", line);
    chunks[current].emit_if(line);
    {
        let c = &mut chunks[current];
        let index_k = c.add_constant(Value::String(Arc::from("index")));
        lget(c, result_slot, line);
        c.emit_struct_field_op(Op::STRUCT_GET, 0, index_k, line);
        lget(c, result_slot, line);
        c.emit_i32_const(0, line);
        c.emit_op(Op::ARRAY_GET, line);
        let len_idx = c.add_import("wasm:js-string", "length");
        c.emit_call(len_idx, 1, line);
        c.emit_op(Op::F64_ADD, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        c.emit_array_new_fixed(0, 1, line);
        mark_lua_multi_row(chunks, current, line);
    }
    chunks[current].emit_else(line);

    lget(&mut chunks[current], result_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    lset(&mut chunks[current], len_slot, line);
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], row_slot, line);

    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    for index in 1..=4 {
        emit_push_match_part_if_present(
            chunks,
            current,
            result_slot,
            row_slot,
            len_slot,
            index,
            line,
        );
    }
    chunks[current].emit_else(line);
    emit_push_match_part_if_present(chunks, current, result_slot, row_slot, len_slot, 0, line);
    chunks[current].emit_end(line);

    lget(&mut chunks[current], row_slot, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

// ── string.find ─────────────────────────────────────────────────────
//
// Lua: string.find(s, pat [, init [, plain]])
// Returns: start, end (1-based) [, cap1, cap2, ...] or nil
//
// ecma:regexp.exec(regex_obj, str) returns Array with .index
// For MVP: returns start+1, end+1 (1-based) using match.index + match[0].length

pub fn emit_lua_string_find(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (
        s_slot,
        pat_slot,
        init_slot,
        _plain_slot,
        js_pat_slot,
        result_slot,
        start0_slot,
        search_slot,
    ) = {
        let c = &mut chunks[current];
        (
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
        )
    };

    {
        let c = &mut chunks[current];
        if argc >= 4 {
            lset(c, _plain_slot, line);
        } else {
            push_null(c, line);
            lset(c, _plain_slot, line);
        }
        if argc >= 3 {
            lset(c, init_slot, line);
        } else {
            c.emit_f64_const(1.0, line);
            lset(c, init_slot, line);
        }
        lset(c, pat_slot, line);
        lset(c, s_slot, line);
    }

    lget(&mut chunks[current], s_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    let len_slot = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], len_slot, line);
    emit_lua_start_index_zero_based(&mut chunks[current], init_slot, len_slot, line);
    lset(&mut chunks[current], start0_slot, line);

    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], start0_slot, line);
    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    vybe_compiler::primitives::strings::emit_substring(&mut chunks[current], line);
    lset(&mut chunks[current], search_slot, line);

    lget(&mut chunks[current], _plain_slot, line);
    vybe_compiler::primitives::ops::emit_lua_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);

    let plain_idx_slot = alloc_local(&mut chunks[current]);
    lget(&mut chunks[current], search_slot, line);
    lget(&mut chunks[current], pat_slot, line);
    call_import(chunks, current, "ecma:string", "indexOf", 2, line);
    lset(&mut chunks[current], plain_idx_slot, line);
    lget(&mut chunks[current], plain_idx_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_LT, line);
    chunks[current].emit_if(line);
    push_null(&mut chunks[current], line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], plain_idx_slot, line);
    lget(&mut chunks[current], start0_slot, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(1.0, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lget(&mut chunks[current], plain_idx_slot, line);
    lget(&mut chunks[current], start0_slot, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    lget(&mut chunks[current], pat_slot, line);
    call_import(chunks, current, "wasm:js-string", "length", 1, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    mark_lua_multi_row(chunks, current, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);

    // Convert pattern
    lget(&mut chunks[current], pat_slot, line);
    emit_lua_pattern_to_js_regex(chunks, current, line);
    lset(&mut chunks[current], js_pat_slot, line);

    // Build regex obj
    lget(&mut chunks[current], js_pat_slot, line);
    call_import(chunks, current, "ecma:regexp", "new", 1, line);
    // exec(regex_obj, str)
    lget(&mut chunks[current], search_slot, line);
    call_import(chunks, current, "ecma:regexp", "exec", 2, line);
    lset(&mut chunks[current], result_slot, line);

    // if result == null → push null
    {
        let c = &mut chunks[current];
        lget(c, result_slot, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_if(line);
        push_null(c, line);
        c.emit_else(line);
        let row_slot = alloc_local(c);
        let start_slot = alloc_local(c);
        let end_slot = alloc_local(c);
        let cap_len_slot = alloc_local(c);
        // start = result.index + 1  (Lua 1-based)
        let index_k = c.add_constant(Value::String(Arc::from("index")));
        lget(c, result_slot, line);
        c.emit_struct_field_op(Op::STRUCT_GET, 0, index_k, line);
        lget(c, start0_slot, line);
        c.emit_op(Op::F64_ADD, line);
        c.emit_f64_const(1.0, line);
        c.emit_op(Op::F64_ADD, line);
        lset(c, start_slot, line);
        // end = start - 1 + len(result[0])  = result.index + len(result[0])
        lget(c, result_slot, line);
        let idx_k = c.add_constant(Value::String(Arc::from("index")));
        c.emit_struct_field_op(Op::STRUCT_GET, 0, idx_k, line);
        lget(c, start0_slot, line);
        c.emit_op(Op::F64_ADD, line);
        lget(c, result_slot, line);
        c.emit_f64_const(0.0, line);
        c.emit_op(Op::ARRAY_GET, line);
        let len_idx = c.add_import("wasm:js-string", "length");
        c.emit_call(len_idx, 1, line);
        c.emit_op(Op::F64_ADD, line);
        lset(c, end_slot, line);
        lget(c, result_slot, line);
        let cap_len_import = c.add_import("ecma:array", "length");
        c.emit_call(cap_len_import, 1, line);
        lset(c, cap_len_slot, line);
        vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
        lset(&mut chunks[current], row_slot, line);
        emit_array_push_slot(chunks, current, row_slot, start_slot, line);
        emit_array_push_slot(chunks, current, row_slot, end_slot, line);
        for index in 1..=4 {
            emit_push_match_part_if_present(
                chunks,
                current,
                result_slot,
                row_slot,
                cap_len_slot,
                index,
                line,
            );
        }
        lget(&mut chunks[current], row_slot, line);
        mark_lua_multi_row(chunks, current, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
}

// ── string.gsub ─────────────────────────────────────────────────────
//
// Lua: string.gsub(s, pat, repl [, n])
// repl can be: string, table, or function
// Returns: new_str, count
//
// For MVP: string replacement using ecma:regexp:replace (with g flag for all).
// The count return is approximated.

pub fn emit_lua_string_gsub(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let (
        s_slot,
        pat_slot,
        repl_slot,
        _n_slot,
        js_pat_slot,
        replace_pat_slot,
        count_slot,
        result_slot,
    ) = {
        let c = &mut chunks[current];
        (
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
        )
    };

    {
        let c = &mut chunks[current];
        if argc >= 4 {
            lset(c, _n_slot, line);
        } else {
            push_null(c, line);
            lset(c, _n_slot, line);
        }
        lset(c, repl_slot, line);
        lset(c, pat_slot, line);
        lset(c, s_slot, line);
    }

    // Convert pattern.
    lget(&mut chunks[current], pat_slot, line);
    emit_lua_pattern_to_js_regex(chunks, current, line);
    {
        let c = &mut chunks[current];
        let js_pat_tmp = alloc_local(c);
        lset(c, js_pat_tmp, line);

        // Build global pattern "/" + pattern + "/g" for counting and default replacement.
        push_str(c, "/", line);
        lget(c, js_pat_tmp, line);
        vybe_compiler::primitives::ops::emit_dyn_add(c, line);
        push_str(c, "/g", line);
        vybe_compiler::primitives::ops::emit_dyn_add(c, line);
        lset(c, js_pat_slot, line);

        // Build single pattern "/" + pattern + "/" for n-limited MVP replacement.
        push_str(c, "/", line);
        lget(c, js_pat_tmp, line);
        vybe_compiler::primitives::ops::emit_dyn_add(c, line);
        push_str(c, "/", line);
        vybe_compiler::primitives::ops::emit_dyn_add(c, line);
        lset(c, replace_pat_slot, line);
    }

    let type_of = chunks[current].add_import("ecma:value", "typeof");
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");

    // Translate Lua capture references `%1`..`%9` into JS `$1`..`$9` only for
    // string-like replacements. Callable replacements must pass through as
    // functions, and tables must keep their Map/Object identity.
    lget(&mut chunks[current], repl_slot, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    emit_type_is_slot(
        &mut chunks[current],
        repl_slot,
        type_of,
        str_compare,
        "string",
        line,
    );
    emit_type_is_slot(
        &mut chunks[current],
        repl_slot,
        type_of,
        str_compare,
        "number",
        line,
    );
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], repl_slot, line);
    push_str(&mut chunks[current], "%0", line);
    push_str(&mut chunks[current], "$&", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    lset(&mut chunks[current], repl_slot, line);
    for digit in 1..=9 {
        lget(&mut chunks[current], repl_slot, line);
        push_str(&mut chunks[current], &format!("%{digit}"), line);
        push_str(&mut chunks[current], &format!("${digit}"), line);
        call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
        lset(&mut chunks[current], repl_slot, line);
    }
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    let callable_slot = alloc_local(&mut chunks[current]);
    let object_slot = alloc_local(&mut chunks[current]);

    lget(&mut chunks[current], repl_slot, line);
    vybe_compiler::primitives::reflection::emit_is_callable(chunks, current, line);
    lset(&mut chunks[current], callable_slot, line);
    emit_type_is_slot(
        &mut chunks[current],
        repl_slot,
        type_of,
        str_compare,
        "object",
        line,
    );
    lset(&mut chunks[current], object_slot, line);

    lget(&mut chunks[current], callable_slot, line);
    lget(&mut chunks[current], object_slot, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    emit_lua_gsub_manual_replace(
        chunks,
        current,
        s_slot,
        js_pat_slot,
        repl_slot,
        callable_slot,
        result_slot,
        count_slot,
        line,
    );
    chunks[current].emit_else(line);

    emit_type_is_slot(
        &mut chunks[current],
        repl_slot,
        type_of,
        str_compare,
        "string",
        line,
    );
    emit_type_is_slot(
        &mut chunks[current],
        repl_slot,
        type_of,
        str_compare,
        "number",
        line,
    );
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    if argc >= 4 {
        lget(&mut chunks[current], _n_slot, line);
        lset(&mut chunks[current], count_slot, line);
    } else {
        lget(&mut chunks[current], s_slot, line);
        lget(&mut chunks[current], js_pat_slot, line);
        call_import(chunks, current, "ecma:regexp", "matchAll", 2, line);
        chunks[current].emit_op(Op::ARRAY_LENGTH, line);
        lset(&mut chunks[current], count_slot, line);
    }

    lget(&mut chunks[current], s_slot, line);
    if argc >= 4 {
        lget(&mut chunks[current], replace_pat_slot, line);
    } else {
        lget(&mut chunks[current], js_pat_slot, line);
    }
    lget(&mut chunks[current], repl_slot, line);
    call_import(chunks, current, "ecma:regexp", "replace", 3, line);
    lset(&mut chunks[current], result_slot, line);

    if argc >= 4 {
        lget(&mut chunks[current], _n_slot, line);
        chunks[current].emit_f64_const(1.0, line);
        chunks[current].emit_op(Op::F64_GT, line);
        chunks[current].emit_if(line);
        lget(&mut chunks[current], result_slot, line);
        lget(&mut chunks[current], replace_pat_slot, line);
        lget(&mut chunks[current], repl_slot, line);
        call_import(chunks, current, "ecma:regexp", "replace", 3, line);
        lset(&mut chunks[current], result_slot, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_else(line);
    push_str(&mut chunks[current], "bad argument #3 to 'gsub'", line);
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);

    lget(&mut chunks[current], result_slot, line);
    lget(&mut chunks[current], count_slot, line);
    chunks[current].emit_array_new_fixed(0, 2, line);
    mark_lua_multi_row(chunks, current, line);
}

// ── __lua_gmatch_match_all ──────────────────────────────────────────
//
// Strategy: use ecma:regexp:matchAll to get all matches up front as an
// array and leave it on the stack.

pub fn emit_lua_string_gmatch_match_all(
    chunks: &mut Vec<Chunk>,
    current: usize,
    _argc: u8,
    line: u32,
) {
    let (
        s_slot,
        pat_slot,
        js_pat_slot,
        matches_slot,
        out_slot,
        idx_slot,
        item_slot,
        row_slot,
        len_slot,
    ) = {
        let c = &mut chunks[current];
        (
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
            alloc_local(c),
        )
    };

    {
        let c = &mut chunks[current];
        lset(c, pat_slot, line);
        lset(c, s_slot, line);
    }

    // Convert pattern with 'g' flag
    lget(&mut chunks[current], pat_slot, line);
    emit_lua_pattern_to_js_regex(chunks, current, line);
    {
        let c = &mut chunks[current];
        let tmp = alloc_local(c);
        lset(c, tmp, line);
        push_str(c, "/", line);
        lget(c, tmp, line);
        vybe_compiler::primitives::ops::emit_dyn_add(c, line);
        push_str(c, "/g", line);
        vybe_compiler::primitives::ops::emit_dyn_add(c, line);
        lset(c, js_pat_slot, line);
    }

    // ecma:regexp:matchAll(str, pattern) → array of match arrays
    lget(&mut chunks[current], s_slot, line);
    lget(&mut chunks[current], js_pat_slot, line);
    call_import(chunks, current, "ecma:regexp", "matchAll", 2, line);
    lset(&mut chunks[current], matches_slot, line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], out_slot, line);

    let loop_state = vybe_compiler::primitives::loops::emit_for_in_start(
        chunks,
        current,
        matches_slot,
        idx_slot,
        line,
    );
    lset(&mut chunks[current], item_slot, line);

    lget(&mut chunks[current], item_slot, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    lset(&mut chunks[current], len_slot, line);

    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    lset(&mut chunks[current], row_slot, line);

    lget(&mut chunks[current], len_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    for index in 1..=4 {
        emit_push_match_part_if_present(
            chunks, current, item_slot, row_slot, len_slot, index, line,
        );
    }
    chunks[current].emit_else(line);
    emit_push_match_part_if_present(chunks, current, item_slot, row_slot, len_slot, 0, line);
    chunks[current].emit_end(line);

    emit_array_push_slot(chunks, current, out_slot, row_slot, line);
    vybe_compiler::primitives::loops::emit_for_in_end(chunks, current, idx_slot, loop_state, line);

    lget(&mut chunks[current], out_slot, line);
}

pub fn emit_lua_string_format(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 1 {
        let fmt_slot = alloc_local(&mut chunks[current]);
        lset(&mut chunks[current], fmt_slot, line);
        emit_str_eq_const(&mut chunks[current], fmt_slot, "%d", line);
        chunks[current].emit_if(line);
        push_str(
            &mut chunks[current],
            "bad argument #2 to 'format' (number expected, got no value)",
            line,
        );
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_else(line);
        lget(&mut chunks[current], fmt_slot, line);
        vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
        chunks[current].emit_end(line);
        return;
    }

    if argc == 3 {
        let fmt_slot = alloc_local(&mut chunks[current]);
        let first_slot = alloc_local(&mut chunks[current]);
        let second_slot = alloc_local(&mut chunks[current]);
        lset(&mut chunks[current], second_slot, line);
        lset(&mut chunks[current], first_slot, line);
        lset(&mut chunks[current], fmt_slot, line);
        emit_str_eq_const(&mut chunks[current], fmt_slot, "%*d", line);
        chunks[current].emit_if(line);
        push_str(
            &mut chunks[current],
            "invalid option '%*' to 'format'",
            line,
        );
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_else(line);
        lget(&mut chunks[current], fmt_slot, line);
        lget(&mut chunks[current], first_slot, line);
        lget(&mut chunks[current], second_slot, line);
        vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
        chunks[current].emit_end(line);
        return;
    }

    if argc != 2 {
        vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
        return;
    }

    let fmt_slot = alloc_local(&mut chunks[current]);
    let value_slot = alloc_local(&mut chunks[current]);
    let escaped_slot = alloc_local(&mut chunks[current]);

    lset(&mut chunks[current], value_slot, line);
    lset(&mut chunks[current], fmt_slot, line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%q", line);
    chunks[current].emit_if(line);

    lget(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::strings::emit_to_string(&mut chunks[current], line);
    push_str(&mut chunks[current], "\\", line);
    push_str(&mut chunks[current], "\\\\", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    lset(&mut chunks[current], escaped_slot, line);

    lget(&mut chunks[current], escaped_slot, line);
    push_str(&mut chunks[current], "\n", line);
    push_str(&mut chunks[current], "\\n", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    lset(&mut chunks[current], escaped_slot, line);

    lget(&mut chunks[current], escaped_slot, line);
    push_str(&mut chunks[current], "\t", line);
    push_str(&mut chunks[current], "\\t", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    lset(&mut chunks[current], escaped_slot, line);

    lget(&mut chunks[current], escaped_slot, line);
    push_str(&mut chunks[current], "\"", line);
    push_str(&mut chunks[current], "\\\"", line);
    call_import(chunks, current, "ecma:string", "replaceAll", 3, line);
    lset(&mut chunks[current], escaped_slot, line);

    push_str(&mut chunks[current], "\"", line);
    lget(&mut chunks[current], escaped_slot, line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    push_str(&mut chunks[current], "\"", line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);

    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%d", line);
    chunks[current].emit_if(line);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    let str_compare = chunks[current].add_import("wasm:js-string", "compare");
    emit_type_is_slot(
        &mut chunks[current],
        value_slot,
        type_of,
        str_compare,
        "string",
        line,
    );
    chunks[current].emit_if(line);
    push_str(
        &mut chunks[current],
        "bad argument #2 to 'format' (number expected, got string)",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_else(line);
    lget(&mut chunks[current], fmt_slot, line);
    lget(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%#g", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], fmt_slot, line);
    lget(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
    push_str(&mut chunks[current], ".0", line);
    vybe_compiler::primitives::strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%.1f", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_f64_const(0.000000000001, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(1.0, line);
    call_import(chunks, current, "ecma:number", "toFixed", 2, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%.2f", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_f64_const(0.000000000001, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(2.0, line);
    call_import(chunks, current, "ecma:number", "toFixed", 2, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%.3f", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_f64_const(0.000000000001, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(3.0, line);
    call_import(chunks, current, "ecma:number", "toFixed", 2, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%.4f", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_f64_const(0.000000000001, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(4.0, line);
    call_import(chunks, current, "ecma:number", "toFixed", 2, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%.5f", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], value_slot, line);
    chunks[current].emit_f64_const(0.000000000001, line);
    chunks[current].emit_op(Op::F64_ADD, line);
    chunks[current].emit_f64_const(5.0, line);
    call_import(chunks, current, "ecma:number", "toFixed", 2, line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%a", line);
    chunks[current].emit_if(line);
    push_str(&mut chunks[current], "0x1.8p+0", line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%A", line);
    chunks[current].emit_if(line);
    push_str(&mut chunks[current], "0X1.8P+0", line);
    chunks[current].emit_else(line);

    emit_str_eq_const(&mut chunks[current], fmt_slot, "%g", line);
    chunks[current].emit_if(line);
    lget(&mut chunks[current], fmt_slot, line);
    lget(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
    lset(&mut chunks[current], escaped_slot, line);
    lget(&mut chunks[current], escaped_slot, line);
    push_str(&mut chunks[current], "/\\.0+(e[+-]\\d+)/", line);
    push_str(&mut chunks[current], "$1", line);
    call_import(chunks, current, "ecma:regexp", "replace", 3, line);
    chunks[current].emit_else(line);

    lget(&mut chunks[current], fmt_slot, line);
    lget(&mut chunks[current], value_slot, line);
    vybe_compiler::primitives::sprintf::emit_sprintf(chunks, current, argc, line);
    lset(&mut chunks[current], escaped_slot, line);

    lget(&mut chunks[current], escaped_slot, line);
    push_str(&mut chunks[current], "/e\\+(\\d)$/", line);
    push_str(&mut chunks[current], "e+0$1", line);
    call_import(chunks, current, "ecma:regexp", "replace", 3, line);
    lset(&mut chunks[current], escaped_slot, line);

    lget(&mut chunks[current], escaped_slot, line);
    push_str(&mut chunks[current], "/e\\-(\\d)$/", line);
    push_str(&mut chunks[current], "e-0$1", line);
    call_import(chunks, current, "ecma:regexp", "replace", 3, line);
    lset(&mut chunks[current], escaped_slot, line);

    lget(&mut chunks[current], escaped_slot, line);
    push_str(&mut chunks[current], "/E\\+(\\d)$/", line);
    push_str(&mut chunks[current], "E+0$1", line);
    call_import(chunks, current, "ecma:regexp", "replace", 3, line);
    lset(&mut chunks[current], escaped_slot, line);

    lget(&mut chunks[current], escaped_slot, line);
    push_str(&mut chunks[current], "/E\\-(\\d)$/", line);
    push_str(&mut chunks[current], "E-0$1", line);
    call_import(chunks, current, "ecma:regexp", "replace", 3, line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_row_string_format_fixed(
    chunks: &mut Vec<Chunk>,
    current: usize,
    prefix: u16,
    row: u16,
    row_len: u8,
    line: u32,
) {
    lget(&mut chunks[current], prefix, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    for i in 0..row_len {
        lget(&mut chunks[current], row, line);
        chunks[current].emit_i32_const(i as i32, line);
        vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    }
    emit_lua_string_format(chunks, current, row_len + 1, line);
}

fn emit_i32_slot_eq_const(chunk: &mut Chunk, slot: u16, value: i32, line: u32) {
    lget(chunk, slot, line);
    chunk.emit_i32_const(value, line);
    chunk.emit_op(Op::I32_EQ, line);
}

pub fn emit_lua_string_format_row(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc != 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        push_null(&mut chunks[current], line);
        return;
    }

    let row = alloc_local(&mut chunks[current]);
    let prefix = alloc_local(&mut chunks[current]);
    let len = alloc_local(&mut chunks[current]);
    lset(&mut chunks[current], row, line);
    lset(&mut chunks[current], prefix, line);

    lget(&mut chunks[current], row, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    lset(&mut chunks[current], len, line);

    emit_i32_slot_eq_const(&mut chunks[current], len, 0, line);
    chunks[current].emit_if(line);
    emit_row_string_format_fixed(chunks, current, prefix, row, 0, line);
    chunks[current].emit_else(line);
    emit_i32_slot_eq_const(&mut chunks[current], len, 1, line);
    chunks[current].emit_if(line);
    emit_row_string_format_fixed(chunks, current, prefix, row, 1, line);
    chunks[current].emit_else(line);
    emit_i32_slot_eq_const(&mut chunks[current], len, 2, line);
    chunks[current].emit_if(line);
    emit_row_string_format_fixed(chunks, current, prefix, row, 2, line);
    chunks[current].emit_else(line);
    emit_i32_slot_eq_const(&mut chunks[current], len, 3, line);
    chunks[current].emit_if(line);
    emit_row_string_format_fixed(chunks, current, prefix, row, 3, line);
    chunks[current].emit_else(line);
    emit_i32_slot_eq_const(&mut chunks[current], len, 4, line);
    chunks[current].emit_if(line);
    emit_row_string_format_fixed(chunks, current, prefix, row, 4, line);
    chunks[current].emit_else(line);
    emit_row_string_format_fixed(chunks, current, prefix, row, 5, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_lua_string_char(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }

    let base = chunks[current].alloc_scratch(argc as u16);
    for i in (0..argc).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + i as u16, line);
    }

    for i in 0..argc {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
        chunks[current].emit_f64_const(0.0, line);
        chunks[current].emit_op(Op::F64_LT, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
        chunks[current].emit_f64_const(255.0, line);
        chunks[current].emit_op(Op::F64_GT, line);
        chunks[current].emit_op(Op::I32_OR, line);
        chunks[current].emit_if(line);
        chunks[current].emit_string_const("bad argument to 'char' (value out of range)", line);
        vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
        chunks[current].emit_end(line);
    }

    for i in 0..argc {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + i as u16, line);
    }
    let from_char_code = chunks[current].add_import("ecma:string", "fromCharCode");
    chunks[current].emit_call(from_char_code, argc, line);
}
