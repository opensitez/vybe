//! Python `heapq` adapter — bytecode-only.
//!
//! The list is kept **sorted ascending**, which satisfies the min-heap
//! invariant `heap[k] <= heap[2k+1]` outright. Everything `heapq` guarantees
//! then falls out directly: `heap[0]` is the smallest, `heappop` returns items
//! in order, and a heapified list is a valid heap. `heappush` costs O(n) rather
//! than O(log n) — the observable behaviour is what `heapq` specifies, and this
//! reuses the sort/splice the runtime already has instead of open-coding
//! sift-up/sift-down.
//!
//! No new host fns: `ecma:array.{sort,splice,slice,concat}` do the work.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_emitter::instructions::core_wasm;

/// Sort a list in place. Stack: `[]` → `[]`.
fn emit_sort_in_place(chunk: &mut Chunk, h: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    let sort = chunk.add_import("ecma:array", "sort");
    chunk.emit_call(sort, 1, line);
    chunk.emit_op(Op::DROP, line);
}

/// `heapq.heapify(h)` — sorting establishes the invariant. In-place, returns None.
/// Stack: `[h]` → `[null]`.
pub fn emit_heapify(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);
    emit_sort_in_place(chunk, h, line);
    chunk.emit_op(Op::NULL, line);
}

/// Push `x` onto the sorted list, keeping it sorted. Stack: `[]` → `[]`.
fn emit_push_sorted(chunk: &mut Chunk, h: u16, x: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    emit_sort_in_place(chunk, h, line);
}

/// `heapq.heappush(h, x)`. Stack: `[h, x]` → `[null]`.
pub fn emit_heappush(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let x = chunk.alloc_scratch(1);
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);
    emit_push_sorted(chunk, h, x, line);
    chunk.emit_op(Op::NULL, line);
}

/// Remove and return `h[0]`. Stack: `[]` → `[value]`.
fn emit_pop_front(chunk: &mut Chunk, h: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, h, line);
    core_wasm::i32_const(chunk, line, 0);
    core_wasm::i32_const(chunk, line, 1);
    let splice = chunk.add_import("ecma:array", "splice");
    chunk.emit_call(splice, 3, line);
    // splice returns the removed slice — the item is its only element.
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// `heapq.heappop(h)` — the smallest item, which a sorted list keeps at 0.
/// Stack: `[h]` → `[value]`.
pub fn emit_heappop(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);
    emit_pop_front(chunk, h, line);
}

/// `heapq.heapreplace(h, x)` — pop then push, unconditionally.
/// Stack: `[h, x]` → `[value]`.
pub fn emit_heapreplace(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
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

/// `heapq.heappushpop(h, x)` — push then pop, but when `x` is already the
/// smallest it is returned untouched and the heap is left alone.
/// Stack: `[h, x]` → `[value]`.
pub fn emit_heappushpop(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let x = chunk.alloc_scratch(1);
    let h = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, x, line);
    chunk.emit_op_u16(Op::LOCAL_SET, h, line);

    // An empty heap, or x <= h[0] → x is the answer.
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
    vybe_emitter::ops::emit_dyn_le(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, x, line);
    chunk.emit_else(line);
    emit_push_sorted(chunk, h, x, line);
    emit_pop_front(chunk, h, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `nsmallest(n, data)` / `nlargest(n, data)` — the first `n` of the sorted
/// data, from whichever end. Stack: `[n, data]` → `[array]`.
fn emit_n_of(chunks: &mut [Chunk], current: usize, largest: bool, line: u32) {
    let chunk = &mut chunks[current];
    let data = chunk.alloc_scratch(1);
    let n = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, data, line);
    chunk.emit_op_u16(Op::LOCAL_SET, n, line);

    chunk.emit_op_u16(Op::LOCAL_GET, data, line);
    let sorted = chunk.add_import("ecma:array", "toSorted");
    chunk.emit_call(sorted, 1, line);
    if largest {
        let rev = chunk.add_import("ecma:array", "reversed");
        chunk.emit_call(rev, 1, line);
    }
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, n, line);
    let slice = chunk.add_import("ecma:array", "slice");
    chunk.emit_call(slice, 3, line);
}

pub fn emit_nsmallest(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_n_of(chunks, current, false, line);
}

pub fn emit_nlargest(chunks: &mut [Chunk], current: usize, _argc: u8, line: u32) {
    emit_n_of(chunks, current, true, line);
}

/// `heapq.merge(*iterables)` — the inputs are each already sorted, so their
/// concatenation sorted is the merge. Stack: `[a, b, …]` → `[array]`.
pub fn emit_merge(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(argc.max(1) as u16);
    for i in (0..argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
    }
    if argc == 0 {
        vybe_emitter::collections::emit_array_new(chunks, current, 0, line);
        return;
    }
    chunk.emit_op_u16(Op::LOCAL_GET, base, line);
    for i in 1..argc as u16 {
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        let concat = chunk.add_import("ecma:array", "concat");
        chunk.emit_call(concat, 2, line);
    }
    let sorted = chunk.add_import("ecma:array", "toSorted");
    chunk.emit_call(sorted, 1, line);
}
