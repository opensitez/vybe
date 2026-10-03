//! Shared heap primitives.
//!
//! The core representation is an array kept sorted ascending. That is a valid
//! min-heap invariant, and it gives compatible observable behaviour for
//! language surfaces such as Python `heapq` and priority-queue adapters.

use crate::primitives::collections;
use crate::primitives::instructions::core_wasm;
use crate::primitives::ops;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn emit_sort_in_place(c: &mut Chunk, h: u16, line: u32) {
    let mut imports = Chunk::new("imports");
    collections::emit_stable_merge_sort(&mut imports, c, h, line, |_, c, lhs, rhs, line| {
        get(c, lhs, line);
        get(c, rhs, line);
        ops::emit_dyn_gt(c, line);
    });
}

// Upper-bound insertion preserves equal-item order and the sorted-array
// representation used by the queue adapters. Only O(log n) comparisons are
// needed; the common splice operation moves the trailing elements once.
fn emit_binary_insert(
    c: &mut Chunk,
    h: u16,
    x: u16,
    line: u32,
    mut greater: impl FnMut(&mut Chunk, u16, u16, u32),
) {
    c.local_count = c.local_count.max(h + 1).max(x + 1);
    let base = c.alloc_scratch(4);
    let lo = base;
    let hi = base + 1;
    let mid = base + 2;
    let item = base + 3;
    c.emit_f64_const(0.0, line);
    set(c, lo, line);
    get(c, h, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    set(c, hi, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, lo, line);
    get(c, hi, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    get(c, hi, line);
    get(c, lo, line);
    c.emit_op(Op::F64_SUB, line);
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_DIV, line);
    c.emit_op(Op::F64_FLOOR, line);
    get(c, lo, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, mid, line);
    read(c, h, mid, item, line);
    greater(c, item, x, line);
    c.emit_if(line);
    get(c, mid, line);
    set(c, hi, line);
    c.emit_else(line);
    get(c, mid, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, lo, line);
    c.emit_end(line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    get(c, h, line);
    get(c, lo, line);
    c.emit_i32_const(0, line);
    get(c, x, line);
    let splice = c.add_import("ecma:array", "splice");
    c.emit_call(splice, 4, line);
    c.emit_op(Op::DROP, line);
}

fn emit_comparator_greater(c: &mut Chunk, line: u32) {
    // Comparator results undergo ToNumber before comparing with zero, as in
    // the existing sort surface. NaN and signed zero both preserve ties.
    let to_number = c.add_import("ecma:value", "toNumber");
    c.emit_call(to_number, 1, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_GT, line);
}

/// Establish the heap invariant in place. Stack: `[heap]` -> `[null]`.
pub fn emit_heapify(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);
    emit_sort_in_place(chunk, h, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_push_sorted(c: &mut Chunk, h: u16, x: u16, line: u32) {
    emit_binary_insert(c, h, x, line, |c, lhs, rhs, line| {
        get(c, lhs, line);
        get(c, rhs, line);
        ops::emit_dyn_gt(c, line);
    });
}

pub fn emit_push_sorted_with_comparator_func(
    c: &mut Chunk,
    h: u16,
    x: u16,
    comparator_idx: usize,
    line: u32,
) {
    emit_binary_insert(c, h, x, line, |c, lhs, rhs, line| {
        c.emit_op_u16(Op::REF_FUNC, comparator_idx as u16, line);
        c.emit(0, line);
        get(c, lhs, line);
        get(c, rhs, line);
        // This API names an explicit two-argument helper declaration (the SPL
        // comparators), rather than a source callback with an implicit receiver.
        crate::primitives::callable::emit_direct_invoke_chunk(c, 2, line);
        emit_comparator_greater(c, line);
    });
}

/// Push one item into the heap. Stack: `[heap, value]` -> `[null]`.
pub fn emit_push(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let x = chunk.alloc_scratch(1);
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);
    emit_push_sorted(chunk, h, x, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Push one item into a comparator-backed heap.
/// Stack: `[heap, value, comparator]` -> `[null]`.
pub fn emit_push_with_comparator(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let abi = crate::primitives::class_context::module_receiver_abi(chunks);
    let c = &mut chunks[current];
    let base = c.alloc_scratch(3);
    let cmp = base;
    let x = base + 1;
    let h = base + 2;
    set(c, cmp, line);
    set(c, x, line);
    set(c, h, line);
    emit_binary_insert(c, h, x, line, |c, lhs, rhs, line| {
        get(c, cmp, line);
        let recv = crate::primitives::callable::emit_callback_receiver(c, abi, line);
        get(c, lhs, line);
        get(c, rhs, line);
        crate::primitives::callable::emit_direct_invoke_chunk(c, 2 + recv, line);
        emit_comparator_greater(c, line);
    });
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_pop_front(chunk: &mut Chunk, h: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    core_wasm::i32_const(chunk, line, 0);
    core_wasm::i32_const(chunk, line, 1);
    let splice = chunk.add_import("ecma:array", "splice");
    chunk.emit_call(splice, 3, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// Pop and return the minimum item. Stack: `[heap]` -> `[value]`.
pub fn emit_pop(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);
    emit_pop_front(chunk, h, line);
}

/// Pop then push, returning the removed item. Stack: `[heap, value]` -> `[value]`.
pub fn emit_replace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let x = chunk.alloc_scratch(1);
    let h = chunk.alloc_scratch(1);
    let out = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);
    emit_pop_front(chunk, h, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    emit_push_sorted(chunk, h, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Push then pop, with the usual heap push-pop fast path.
/// Stack: `[heap, value]` -> `[value]`.
pub fn emit_push_pop(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let x = chunk.alloc_scratch(1);
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);

    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    // Only replace when root < x. Negating x <= root is different for NaN.
    ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_else(line);
    emit_push_sorted(chunk, h, x, line);
    emit_pop_front(chunk, h, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn get(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_GET, slot, line);
}
fn set(c: &mut Chunk, slot: u16, line: u32) {
    c.emit_op_u16(Op::LOCAL_SET, slot, line);
}
fn field(c: &mut Chunk, pair: u16, index: i32, line: u32) {
    get(c, pair, line);
    c.emit_i32_const(index, line);
    c.emit_op(Op::ARRAY_GET, line);
}
fn read(c: &mut Chunk, arr: u16, index: u16, value: u16, line: u32) {
    get(c, arr, line);
    get(c, index, line);
    // Dynamic arrays can exceed the signed-i32 range. Keep their numeric
    // indices on the common array surface; GC gets are only used for the
    // three fixed, proven decoration fields.
    let get = c.add_import("ecma:array", "get");
    c.emit_call(get, 2, line);
    set(c, value, line);
}
fn write(c: &mut Chunk, arr: u16, index: u16, value: u16, line: u32) {
    get(c, arr, line);
    get(c, index, line);
    get(c, value, line);
    let set = c.add_import("ecma:array", "set");
    c.emit_call(set, 3, line);
    c.emit_op(Op::DROP, line);
}
fn increment(c: &mut Chunk, index: u16, line: u32) {
    get(c, index, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, index, line);
}
fn key_better(c: &mut Chunk, largest: bool, line: u32) {
    if largest {
        ops::emit_dyn_gt(c, line);
    } else {
        ops::emit_dyn_lt(c, line);
    }
}

// A decoration is [key, original index, value]. The index is a proven
// numeric counter, and makes equal-key selection stable in either direction.
fn pair_better(c: &mut Chunk, lhs: u16, rhs: u16, largest: bool, line: u32) {
    field(c, lhs, 0, line);
    field(c, rhs, 0, line);
    key_better(c, largest, line);
    c.emit_if_i32(line);
    c.emit_i32_const(1, line);
    c.emit_else(line);
    field(c, rhs, 0, line);
    field(c, lhs, 0, line);
    key_better(c, largest, line);
    c.emit_if_i32(line);
    c.emit_i32_const(0, line);
    c.emit_else(line);
    field(c, lhs, 1, line);
    field(c, rhs, 1, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_end(line);
    c.emit_end(line);
}

// Keep the worst retained decoration at the root of the private heap.
fn sift_up(c: &mut Chunk, arr: u16, pos: u16, item: u16, largest: bool, line: u32) {
    let base = c.alloc_scratch(2);
    let parent = base;
    let value = base + 1;
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, pos, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_LE, line);
    c.emit_br_if(1, line);
    get(c, pos, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_SUB, line);
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_DIV, line);
    c.emit_op(Op::F64_FLOOR, line);
    set(c, parent, line);
    read(c, arr, parent, value, line);
    pair_better(c, value, item, largest, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    write(c, arr, pos, value, line);
    get(c, parent, line);
    set(c, pos, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    write(c, arr, pos, item, line);
}

fn sift_down(c: &mut Chunk, arr: u16, len: u16, item: u16, largest: bool, line: u32) {
    let base = c.alloc_scratch(5);
    let pos = base;
    let child = base + 1;
    let right = base + 2;
    let value = base + 3;
    let other = base + 4;
    c.emit_f64_const(0.0, line);
    set(c, pos, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, pos, line);
    c.emit_f64_const(2.0, line);
    c.emit_op(Op::F64_MUL, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, child, line);
    get(c, child, line);
    get(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    read(c, arr, child, value, line);
    get(c, child, line);
    c.emit_f64_const(1.0, line);
    c.emit_op(Op::F64_ADD, line);
    set(c, right, line);
    get(c, right, line);
    get(c, len, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    read(c, arr, right, other, line);
    pair_better(c, value, other, largest, line);
    c.emit_if(line);
    get(c, right, line);
    set(c, child, line);
    get(c, other, line);
    set(c, value, line);
    c.emit_end(line);
    c.emit_end(line);
    pair_better(c, item, value, largest, line);
    c.emit_op(Op::I32_EQZ, line);
    c.emit_br_if(1, line);
    write(c, arr, pos, value, line);
    get(c, child, line);
    set(c, pos, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    write(c, arr, pos, item, line);
}

fn emit_n_of(chunks: &mut [Chunk], current: usize, argc: u8, largest: bool, line: u32) {
    let abi = crate::primitives::class_context::module_receiver_abi(chunks);
    let c = &mut chunks[current];
    let base = c.alloc_scratch(12);
    let n = base;
    let data = base + 1;
    let key_fn = base + 2;
    let len = base + 3;
    let result = base + 4;
    let index = base + 5;
    let filled = base + 6;
    let pos = base + 7;
    let value = base + 8;
    let key = base + 9;
    let item = base + 10;
    let root = base + 11;
    if argc >= 3 {
        set(c, key_fn, line);
    }
    set(c, data, line);
    set(c, n, line);
    get(c, data, line);
    c.emit_op(Op::ARRAY_LENGTH, line);
    set(c, len, line);
    // Normalize the dynamic count once; subsequent operations consume a
    // proven number. Private indices are numeric counters throughout.
    get(c, n, line);
    let to_number = c.add_import("ecma:value", "toNumber");
    c.emit_call(to_number, 1, line);
    set(c, n, line);
    get(c, n, line);
    get(c, n, line);
    c.emit_op(Op::F64_EQ, line);
    c.emit_if_value(line);
    get(c, n, line);
    c.emit_op(Op::F64_FLOOR, line);
    c.emit_else(line);
    c.emit_f64_const(0.0, line);
    c.emit_end(line);
    get(c, len, line);
    c.emit_op(Op::F64_MIN, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_MAX, line);
    set(c, n, line);
    let mut imports = Chunk::new("imports");
    get(c, n, line);
    collections::emit_new_with_length_into(&mut imports, c, line);
    set(c, result, line);
    c.emit_f64_const(0.0, line);
    set(c, index, line);
    c.emit_f64_const(0.0, line);
    set(c, filled, line);
    let done = c.emit_block(line);
    get(c, n, line);
    c.emit_f64_const(0.0, line);
    c.emit_op(Op::F64_LE, line);
    c.emit_br_if(0, line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, index, line);
    get(c, len, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    read(c, data, index, value, line);
    if argc >= 3 {
        get(c, key_fn, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_if_value(line);
        get(c, value, line);
        c.emit_else(line);
        get(c, key_fn, line);
        let recv = crate::primitives::callable::emit_callback_receiver(c, abi, line);
        get(c, value, line);
        crate::primitives::callable::emit_direct_invoke_chunk(c, 1 + recv, line);
        c.emit_end(line);
    } else {
        get(c, value, line);
    }
    set(c, key, line);
    get(c, filled, line);
    get(c, n, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    get(c, key, line);
    get(c, index, line);
    get(c, value, line);
    c.emit_array_new_fixed(0, 3, line);
    set(c, item, line);
    get(c, filled, line);
    set(c, pos, line);
    // When all elements are requested, decorate directly without maintaining
    // a heap that would immediately be sorted again.
    get(c, n, line);
    get(c, len, line);
    c.emit_op(Op::F64_LT, line);
    c.emit_if(line);
    sift_up(c, result, pos, item, largest, line);
    c.emit_else(line);
    write(c, result, pos, item, line);
    c.emit_end(line);
    increment(c, filled, line);
    c.emit_else(line);
    get(c, result, line);
    c.emit_i32_const(0, line);
    c.emit_op(Op::ARRAY_GET, line);
    set(c, root, line);
    get(c, key, line);
    field(c, root, 0, line);
    key_better(c, largest, line);
    c.emit_if(line);
    get(c, key, line);
    get(c, index, line);
    get(c, value, line);
    c.emit_array_new_fixed(0, 3, line);
    set(c, item, line);
    sift_down(c, result, n, item, largest, line);
    c.emit_end(line);
    c.emit_end(line);
    increment(c, index, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);

    collections::emit_stable_merge_sort(&mut imports, c, result, line, |_, c, lhs, rhs, line| {
        pair_better(c, rhs, lhs, largest, line);
    });
    c.emit_f64_const(0.0, line);
    set(c, index, line);
    let done = c.emit_block(line);
    let (repeat, _) = c.emit_loop_s(line);
    get(c, index, line);
    get(c, n, line);
    c.emit_op(Op::F64_GE, line);
    c.emit_br_if(1, line);
    read(c, result, index, item, line);
    field(c, item, 2, line);
    set(c, value, line);
    write(c, result, index, value, line);
    increment(c, index, line);
    c.emit_br(0, line);
    c.emit_end(line);
    c.patch_loop(repeat);
    c.emit_end(line);
    c.patch_block(done);
    get(c, result, line);
}

/// Return the `n` smallest items from data. Stack: `[n, data]` -> `[array]`.
pub fn emit_nsmallest(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_n_of(chunks, current, argc, false, line);
}

/// Return the `n` largest items from data. Stack: `[n, data]` -> `[array]`.
pub fn emit_nlargest(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_n_of(chunks, current, argc, true, line);
}

/// Merge sorted iterables by concatenating and sorting. Stack: `[a, b, ...]` -> `[array]`.
pub fn emit_merge(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(argc.max(1) as u16);
    for i in (0..argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
    }
    if argc == 0 {
        collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    for i in 1..argc as u16 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + i, line);
        let concat = chunks[current].add_import("ecma:array", "concat");
        chunks[current].emit_call(concat, 2, line);
    }
    let sorted = chunks[current].add_import("ecma:array", "toSorted");
    chunks[current].emit_call(sorted, 1, line);
}
