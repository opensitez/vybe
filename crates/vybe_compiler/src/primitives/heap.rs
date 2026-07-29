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

fn emit_sort_in_place(chunk: &mut Chunk, h: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    let sort = chunk.add_import("ecma:array", "sort");
    chunk.emit_call(sort, 1, line);
    chunk.emit_op(Op::DROP, line);
}

/// Establish the heap invariant in place. Stack: `[heap]` -> `[null]`.
pub fn emit_heapify(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);
    emit_sort_in_place(chunk, h, line);
    chunk.emit_op(Op::NULL, line);
}

fn emit_push_sorted(chunk: &mut Chunk, h: u16, x: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    emit_sort_in_place(chunk, h, line);
}

pub fn emit_push_sorted_with_comparator_func(
    chunk: &mut Chunk,
    h: u16,
    x: u16,
    comparator_idx: usize,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    chunk.emit_op_u16(Op::REF_FUNC, comparator_idx as u16, line);
    chunk.emit(0, line);
    collections::emit_sort_with_comparator_in_chunk(chunk, line);
    chunk.emit_op(Op::DROP, line);
}

/// Push one item into the heap. Stack: `[heap, value]` -> `[null]`.
pub fn emit_push(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let x = chunk.alloc_scratch(1);
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);
    emit_push_sorted(chunk, h, x, line);
    chunk.emit_op(Op::NULL, line);
}

/// Push one item into a comparator-backed heap.
/// Stack: `[heap, value, comparator]` -> `[null]`.
pub fn emit_push_with_comparator(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let cmp = chunk.alloc_scratch(1);
    let x = chunk.alloc_scratch(1);
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, cmp, line);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);

    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cmp, line);
    let _ = chunk;
    collections::emit_sort_with_comparator(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
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
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op(Op::ARRAY_GET, line);
    ops::emit_dyn_le(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_else(line);
    emit_push_sorted(chunk, h, x, line);
    emit_pop_front(chunk, h, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_n_of(chunks: &mut [Chunk], current: usize, argc: u8, largest: bool, line: u32) {
    let n = chunks[current].alloc_scratch(1);
    let data = chunks[current].alloc_scratch(1);
    let key_fn = chunks[current].alloc_scratch(1);
    if argc >= 3 {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, key_fn, line);
        chunk.emit_op_u16(Op::LOCAL_SET, data, line);
        chunk.emit_op_u16(Op::LOCAL_SET, n, line);
        chunk.emit_op_u16(Op::LOCAL_GET, data, line);
        chunk.emit_op_u16(Op::LOCAL_GET, key_fn, line);
        let _ = chunk;
        collections::emit_sort_by_key_in_place(chunks, current, line);
    } else {
        let chunk = &mut chunks[current];
        chunk.emit_op_u16(Op::LOCAL_SET, data, line);
        chunk.emit_op_u16(Op::LOCAL_SET, n, line);
        chunk.emit_op_u16(Op::LOCAL_GET, data, line);
        let sorted = chunk.add_import("ecma:array", "toSorted");
        chunk.emit_call(sorted, 1, line);
    }

    let chunk = &mut chunks[current];
    if largest {
        let rev = chunk.add_import("ecma:array", "toReversed");
        chunk.emit_call(rev, 1, line);
    }
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    let slice = chunk.add_import("ecma:array", "slice");
    chunk.emit_call(slice, 3, line);
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
