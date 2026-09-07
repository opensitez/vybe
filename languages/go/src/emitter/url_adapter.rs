use vybe_compiler::primitives::{collections, ops, strings};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

pub fn emit_helper(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    match name {
        "go.url_parse_qs" => emit_parse_qs(chunks, current, argc, line),
        "go.url_encode_values" => emit_urlencode(chunks, current, argc, line),
        "go.url_values_set" => emit_values_set(chunks, current, argc, line),
        "go.url_values_add" => emit_values_add(chunks, current, argc, line),
        _ => return false,
    }
    true
}

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn lget(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

pub fn emit_parse_qs(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        call_import(chunks, current, "ecma:map", "new", 0, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let params = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let values = chunks[current].alloc_scratch(1);
    let keep_blank = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        lget(chunks, current, base + 1, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        chunks[current].emit_i32_const(1, line);
    }
    lset(chunks, current, keep_blank, line);

    lget(chunks, current, base, line);
    call_import(chunks, current, "web:url", "searchParamsNew", 1, line);
    lset(chunks, current, params, line);
    call_import(chunks, current, "ecma:map", "new", 0, line);
    lset(chunks, current, out, line);

    lget(chunks, current, params, line);
    call_import(chunks, current, "web:url", "searchParamsKeys", 1, line);
    lset(chunks, current, keys, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    lget(chunks, current, keys, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    lset(chunks, current, n, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(chunks, current, keys, line);
    lget(chunks, current, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(chunks, current, key, line);

    lget(chunks, current, params, line);
    lget(chunks, current, key, line);
    call_import(chunks, current, "web:url", "searchParamsGetAll", 2, line);
    lset(chunks, current, values, line);

    lget(chunks, current, keep_blank, line);
    lget(chunks, current, values, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    strings::emit_length(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_NE, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    lget(chunks, current, out, line);
    lget(chunks, current, key, line);
    lget(chunks, current, values, line);
    call_import(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    lget(chunks, current, out, line);
}

pub fn emit_urlencode(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let params = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let value = chunks[current].alloc_scratch(1);
    let doseq = chunks[current].alloc_scratch(1);

    if argc >= 2 {
        lget(chunks, current, base + 1, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
    } else {
        chunks[current].emit_i32_const(1, line);
    }
    lset(chunks, current, doseq, line);

    call_import(chunks, current, "web:url", "searchParamsNew", 0, line);
    lset(chunks, current, params, line);

    lget(chunks, current, base, line);
    call_import(chunks, current, "ecma:object", "keys", 1, line);
    collections::emit_sort(chunks, current, line);
    lset(chunks, current, keys, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    lget(chunks, current, keys, line);
    call_import(chunks, current, "ecma:array", "length", 1, line);
    lset(chunks, current, n, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(chunks, current, keys, line);
    lget(chunks, current, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(chunks, current, key, line);
    lget(chunks, current, base, line);
    lget(chunks, current, key, line);
    collections::emit_get(chunks, current, line);
    lset(chunks, current, value, line);

    lget(chunks, current, doseq, line);
    lget(chunks, current, value, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);
    append_sequence_values(chunks, current, params, key, value, line);
    chunks[current].emit_else(line);
    append_query_value(chunks, current, params, key, value, line);
    chunks[current].emit_end(line);

    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    lget(chunks, current, params, line);
    call_import(chunks, current, "web:url", "searchParamsToString", 1, line);
}

pub fn emit_values_set(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    lget(chunks, current, base + 2, line);
    collections::emit_array_new(chunks, current, 1, line);
    let arr = chunks[current].alloc_scratch(1);
    lset(chunks, current, arr, line);
    lget(chunks, current, base, line);
    lget(chunks, current, base + 1, line);
    lget(chunks, current, arr, line);
    collections::emit_set(chunks, current, line);
}

pub fn emit_values_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 3 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let current_values = chunks[current].alloc_scratch(1);
    let next_values = chunks[current].alloc_scratch(1);

    lget(chunks, current, base, line);
    lget(chunks, current, base + 1, line);
    collections::emit_get(chunks, current, line);
    lset(chunks, current, current_values, line);

    lget(chunks, current, current_values, line);
    call_import(chunks, current, "ecma:array", "isArray", 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    lget(chunks, current, current_values, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_end(line);
    lset(chunks, current, next_values, line);

    lget(chunks, current, next_values, line);
    lget(chunks, current, base + 2, line);
    call_import(chunks, current, "ecma:array", "push", 2, line);
    chunks[current].emit_op(Op::DROP, line);

    lget(chunks, current, base, line);
    lget(chunks, current, base + 1, line);
    lget(chunks, current, next_values, line);
    collections::emit_set(chunks, current, line);
}

fn append_query_value(
    chunks: &mut [Chunk],
    current: usize,
    params: u16,
    key: u16,
    value: u16,
    line: u32,
) {
    lget(chunks, current, params, line);
    lget(chunks, current, key, line);
    lget(chunks, current, value, line);
    call_import(chunks, current, "web:url", "searchParamsAppend", 3, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn append_sequence_values(
    chunks: &mut [Chunk],
    current: usize,
    params: u16,
    key: u16,
    values: u16,
    line: u32,
) {
    let j = chunks[current].alloc_scratch(1);
    let n = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);

    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, j, line);
    lget(chunks, current, values, line);
    collections::emit_len(chunks, current, line);
    lset(chunks, current, n, line);

    let block = chunks[current].emit_block(line);
    let lp = chunks[current].emit_loop_s(line).0;
    lget(chunks, current, j, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    lget(chunks, current, values, line);
    lget(chunks, current, j, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    lset(chunks, current, item, line);
    append_query_value(chunks, current, params, key, item, line);

    lget(chunks, current, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
}
