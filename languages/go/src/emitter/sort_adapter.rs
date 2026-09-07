use vybe_compiler::primitives::{callable, class_context, collections, loops, ops, tuples};
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
        "go.sort_search" if argc == 2 => emit_sort_search(chunks, current, line),
        "go.sort_search_ordered" if argc == 2 => emit_sort_search_ordered(chunks, current, line),
        "go.sort_find" if argc == 2 => emit_sort_find(chunks, current, line),
        "go.sort_slice" if argc == 3 => emit_sort_slice(chunks, current, line),
        "go.sort_is_sorted" if argc == 2 => emit_sort_is_sorted(chunks, current, line),
        "go.sort_reverse" if argc == 1 => emit_sort_reverse(chunks, current, line),
        _ => return false,
    }
    true
}

fn emit_sort_search_ordered(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(9);
    let slice = base;
    let needle = base + 1;
    let n = base + 2;
    let lo = base + 3;
    let hi = base + 4;
    let mid = base + 5;
    let i = base + 6;
    let result = base + 7;
    let found = base + 8;

    lset(chunks, current, needle, line);
    lset(chunks, current, slice, line);
    lget(chunks, current, slice, line);
    collections::emit_len(chunks, current, line);
    lset(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, lo, line);
    lget(chunks, current, n, line);
    lset(chunks, current, hi, line);

    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, lo, line);
    lget(chunks, current, hi, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    lget(chunks, current, lo, line);
    lget(chunks, current, hi, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SHR_S, line);
    lset(chunks, current, mid, line);

    emit_slice_get(chunks, current, slice, mid, line);
    lget(chunks, current, needle, line);
    ops::emit_dyn_ge(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(chunks, current, mid, line);
    lset(chunks, current, hi, line);
    chunks[current].emit_else(line);
    lget(chunks, current, mid, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, lo, line);
    chunks[current].emit_end(line);

    loops::emit_loop_end(chunks, current, state, line);

    lget(chunks, current, lo, line);
    lset(chunks, current, result, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, found, line);

    lget(chunks, current, lo, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    emit_slice_get(chunks, current, slice, lo, line);
    lget(chunks, current, needle, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_i32_const(1, line);
    lset(chunks, current, found, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, i, line);
    let scan = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    lget(chunks, current, found, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);
    emit_slice_get(chunks, current, slice, i, line);
    lget(chunks, current, needle, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(chunks, current, i, line);
    lset(chunks, current, result, line);
    chunks[current].emit_i32_const(1, line);
    lset(chunks, current, found, line);
    chunks[current].emit_end(line);
    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, scan, line);
    lget(chunks, current, result, line);
}

fn emit_sort_search(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(5);
    let n = base;
    let pred = base + 1;
    let lo = base + 2;
    let hi = base + 3;
    let mid = base + 4;

    lset(chunks, current, pred, line);
    lset(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, lo, line);
    lget(chunks, current, n, line);
    lset(chunks, current, hi, line);

    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, lo, line);
    lget(chunks, current, hi, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    lget(chunks, current, lo, line);
    lget(chunks, current, hi, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SHR_S, line);
    lset(chunks, current, mid, line);

    lget(chunks, current, pred, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    lget(chunks, current, mid, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    lget(chunks, current, mid, line);
    lset(chunks, current, hi, line);
    chunks[current].emit_else(line);
    lget(chunks, current, mid, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, lo, line);
    chunks[current].emit_end(line);

    loops::emit_loop_end(chunks, current, state, line);
    lget(chunks, current, lo, line);
}

fn emit_sort_find(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let n = base;
    let cmp = base + 1;
    let idx = base + 2;
    let found = base + 3;
    let lo = base + 4;
    let hi = base + 5;

    lset(chunks, current, cmp, line);
    lset(chunks, current, n, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, found, line);

    lget(chunks, current, n, line);
    lget(chunks, current, cmp, line);
    emit_sort_search_from_stack(chunks, current, lo, hi, idx, line);
    lset(chunks, current, idx, line);

    lget(chunks, current, idx, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    lget(chunks, current, cmp, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    lget(chunks, current, idx, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    lset(chunks, current, found, line);
    chunks[current].emit_end(line);

    lget(chunks, current, idx, line);
    lget(chunks, current, found, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    tuples::emit_tuple(chunks, current, 2, line);
}

fn emit_sort_slice(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(7);
    let slice = base;
    let less = base + 1;
    let _stable = base + 2;
    let n = base + 3;
    let i = base + 4;
    let j = base + 5;
    let tmp = base + 6;

    lset(chunks, current, _stable, line);
    lset(chunks, current, less, line);
    lset(chunks, current, slice, line);
    lget(chunks, current, slice, line);
    collections::emit_len(chunks, current, line);
    lset(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    lset(chunks, current, i, line);

    let outer = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    lget(chunks, current, i, line);
    lset(chunks, current, j, line);

    let inner = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, j, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    lget(chunks, current, less, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    lget(chunks, current, j, line);
    lget(chunks, current, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 2 + recv, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    emit_slice_swap(chunks, current, slice, j, tmp, line);

    lget(chunks, current, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, j, line);
    loops::emit_loop_end(chunks, current, inner, line);

    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, outer, line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_sort_reverse(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(6);
    let slice = base;
    let n = base + 1;
    let i = base + 2;
    let j = base + 3;
    let tmp = base + 4;
    let prev = base + 5;

    lset(chunks, current, slice, line);
    lget(chunks, current, slice, line);
    collections::emit_len(chunks, current, line);
    lset(chunks, current, n, line);
    chunks[current].emit_i32_const(1, line);
    lset(chunks, current, i, line);

    let outer = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    lget(chunks, current, i, line);
    lset(chunks, current, j, line);

    let inner = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, j, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    lget(chunks, current, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, prev, line);
    emit_slice_get(chunks, current, slice, j, line);
    emit_slice_get(chunks, current, slice, prev, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_end(line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    emit_slice_swap_with_prev(chunks, current, slice, j, prev, tmp, line);
    lget(chunks, current, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, j, line);
    loops::emit_loop_end(chunks, current, inner, line);

    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, outer, line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_slice_get(chunks: &mut [Chunk], current: usize, slice: u16, index: u16, line: u32) {
    lget(chunks, current, slice, line);
    lget(chunks, current, index, line);
    collections::emit_get(chunks, current, line);
}

fn emit_slice_set_from_local(
    chunks: &mut [Chunk],
    current: usize,
    slice: u16,
    index: u16,
    value: u16,
    line: u32,
) {
    lget(chunks, current, slice, line);
    lget(chunks, current, index, line);
    lget(chunks, current, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn emit_slice_swap(chunks: &mut [Chunk], current: usize, slice: u16, j: u16, tmp: u16, line: u32) {
    let prev = chunks[current].alloc_scratch(1);
    lget(chunks, current, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    lset(chunks, current, prev, line);
    emit_slice_swap_with_prev(chunks, current, slice, j, prev, tmp, line);
}

fn emit_slice_swap_with_prev(
    chunks: &mut [Chunk],
    current: usize,
    slice: u16,
    j: u16,
    prev: u16,
    tmp: u16,
    line: u32,
) {
    emit_slice_get(chunks, current, slice, j, line);
    lset(chunks, current, tmp, line);
    lget(chunks, current, slice, line);
    lget(chunks, current, j, line);
    emit_slice_get(chunks, current, slice, prev, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_slice_set_from_local(chunks, current, slice, prev, tmp, line);
}

fn emit_sort_is_sorted(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(4);
    let n = base;
    let less = base + 1;
    let i = base + 2;
    let result = base + 3;

    lset(chunks, current, less, line);
    lset(chunks, current, n, line);
    chunks[current].emit_bool_const(true, line);
    lset(chunks, current, result, line);
    chunks[current].emit_i32_const(1, line);
    lset(chunks, current, i, line);

    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, i, line);
    lget(chunks, current, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    lget(chunks, current, less, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    lget(chunks, current, i, line);
    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 2 + recv, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    lset(chunks, current, result, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    lget(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, i, line);
    loops::emit_loop_end(chunks, current, state, line);

    lget(chunks, current, result, line);
}

fn emit_sort_search_from_stack(
    chunks: &mut [Chunk],
    current: usize,
    lo: u16,
    hi: u16,
    mid: u16,
    line: u32,
) {
    let pred = chunks[current].alloc_scratch(1);
    lset(chunks, current, pred, line);
    lset(chunks, current, hi, line);
    chunks[current].emit_i32_const(0, line);
    lset(chunks, current, lo, line);

    let state = loops::emit_loop_start(chunks, current, line);
    lget(chunks, current, lo, line);
    lget(chunks, current, hi, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond_from_i32(chunks, current, line);

    lget(chunks, current, lo, line);
    lget(chunks, current, hi, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SHR_S, line);
    lset(chunks, current, mid, line);

    lget(chunks, current, pred, line);
    let abi = class_context::module_receiver_abi(chunks);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    lget(chunks, current, mid, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LE_S, line);
    chunks[current].emit_if(line);
    lget(chunks, current, mid, line);
    lset(chunks, current, hi, line);
    chunks[current].emit_else(line);
    lget(chunks, current, mid, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    lset(chunks, current, lo, line);
    chunks[current].emit_end(line);

    loops::emit_loop_end(chunks, current, state, line);
    lget(chunks, current, lo, line);
}

fn lget(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn lset(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}
