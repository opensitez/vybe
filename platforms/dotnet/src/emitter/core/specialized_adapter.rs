//! Emits for `BitVector32`, `OrderedDictionary` and `PriorityQueue`.
//!
//! Only the members that cannot be expressed by an existing leaf live here.
//! Everything else on those types resolves straight to `ecma:map` or
//! `collections.*` from the descriptors, because an ordered dictionary IS a
//! map and a read-only wrapper IS its wrapped collection.

use vybe_compiler::primitives::{callable, collections, ops};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// Where a `PriorityQueue` parks its comparer.
///
/// A property on the backing array rather than a parallel structure: the queue
/// IS the array of pairs, and a comparer stored anywhere else would have to be
/// found again from the receiver alone.
const PQ_COMPARER: &str = "__dotnet_pq_comparer";

fn stash(chunks: &mut [Chunk], current: usize, argc: u16, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc);
    for offset in (0..argc).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

fn get(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn get_field(chunks: &mut [Chunk], current: usize, obj: u16, field: &str, line: u32) {
    get(chunks, current, obj, line);
    chunks[current].emit_string_const(field, line);
    collections::emit_get(chunks, current, line);
}

fn set_field(chunks: &mut [Chunk], current: usize, obj: u16, field: &str, value: u16, line: u32) {
    get(chunks, current, obj, line);
    chunks[current].emit_string_const(field, line);
    get(chunks, current, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

// ── BitVector32 ───────────────────────────────────────────────────────────

/// `BitVector32.CreateMask()` — the first mask, which §is always bit 0.
pub fn emit_create_mask(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(1, line);
}

/// `BitVector32.CreateMask(previous)` — the next bit up.
pub fn emit_create_mask_next(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SHL, line);
}

// ── OrderedDictionary ─────────────────────────────────────────────────────

/// `OrderedDictionary` — an array of VALUES that also carries each value under
/// its own key as a property.
///
/// ⛔ The declared `Item` member is NEVER CONSULTED: `od[0]` compiles to native
/// indexing on whatever the object is, so the indexer cannot be intercepted.
/// That is the whole design constraint here, and it was only visible after two
/// wrong backings — a Map answered `od["k"]` but not `od[0]`, and an array of
/// PAIRS answered `od[0]` with the pair.
///
/// Storing each value twice — positionally and under its key — makes BOTH
/// spellings resolve through native indexing, which is the only path that runs.
/// `Count` stays right because a string-keyed property is not an element.
/// `od.Add(key, value)` — stack `[recv, key, value]`.
pub fn emit_ordered_dictionary_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash(chunks, current, 3, line);
    let (recv, key, value) = (base, base + 1, base + 2);

    // Positional half — the element, which is what `od[0]` reads.
    get(chunks, current, recv, line);
    get(chunks, current, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // Keyed half — the same value under its own name, which is what
    // `od["key"]` reads.
    get(chunks, current, recv, line);
    get(chunks, current, key, line);
    get(chunks, current, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

// ── PriorityQueue ─────────────────────────────────────────────────────────

/// `new PriorityQueue<E, P>()` / `new PriorityQueue<E, P>(comparer)`.
///
/// The queue is an array of `[element, priority]` pairs with the comparer
/// parked on it. `argc` distinguishes the two constructors — with one
/// argument the comparer is on the stack, with none there is nothing to pop.
pub fn emit_priority_queue_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let comparer = chunks[current].alloc_scratch(1);
    if argc >= 1 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, comparer, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, comparer, line);
    }
    let queue = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, queue, line);
    set_field(chunks, current, queue, PQ_COMPARER, comparer, line);
    get(chunks, current, queue, line);
}

/// `pq.Enqueue(element, priority)`.
pub fn emit_priority_queue_enqueue(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash(chunks, current, 3, line);
    let (recv, element, priority) = (base, base + 1, base + 2);

    get(chunks, current, recv, line);
    get(chunks, current, element, line);
    get(chunks, current, priority, line);
    collections::emit_array_new(chunks, current, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Leaves the index of the minimum-priority pair in a fresh scratch slot, and
/// returns that slot. Answers -1 for an empty queue.
///
/// "Minimum" is the comparer's order when one was supplied, and natural
/// ascending order otherwise — which is what makes the default a MIN-queue.
fn emit_find_min_index(chunks: &mut [Chunk], current: usize, recv: u16, line: u32) -> u16 {
    let best = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    let cmp = chunks[current].alloc_scratch(1);
    let a = chunks[current].alloc_scratch(1);
    let b = chunks[current].alloc_scratch(1);

    get_field(chunks, current, recv, PQ_COMPARER, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cmp, line);

    get(chunks, current, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);

    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let block = chunks[current].emit_block(line);
    let (lp, _) = chunks[current].emit_loop_s(line);
    get(chunks, current, i, line);
    get(chunks, current, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    // a = priority at i
    get(chunks, current, recv, line);
    get(chunks, current, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);

    get(chunks, current, best, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    // first candidate always wins
    get(chunks, current, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best, line);
    chunks[current].emit_else(line);

    // b = priority at best
    get(chunks, current, recv, line);
    get(chunks, current, best, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);

    get(chunks, current, cmp, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    get(chunks, current, a, line);
    get(chunks, current, b, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    // A `Comparer<T>.Create(fn)` IS `fn`, so the comparer is invoked directly.
    get(chunks, current, cmp, line);
    get(chunks, current, a, line);
    get(chunks, current, b, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 2, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_if(line);
    get(chunks, current, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);

    get(chunks, current, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(lp);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);

    best
}

/// `pq.Peek()` — the minimum-priority ELEMENT, left in the queue.
pub fn emit_priority_queue_peek(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash(chunks, current, 1, line);
    let best = emit_find_min_index(chunks, current, recv, line);

    get(chunks, current, best, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    get(chunks, current, best, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `pq.Dequeue()` — the minimum-priority element, REMOVED.
pub fn emit_priority_queue_dequeue(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash(chunks, current, 1, line);
    let best = emit_find_min_index(chunks, current, recv, line);
    let out = chunks[current].alloc_scratch(1);

    get(chunks, current, best, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_else(line);
    get(chunks, current, recv, line);
    get(chunks, current, best, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    // Remove the pair AFTER reading it — `remove_at` answers the removed
    // value, not the element, and the pair is gone by the time it returns.
    get(chunks, current, recv, line);
    get(chunks, current, best, line);
    collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    get(chunks, current, out, line);
}

/// `pq.Count` — pairs enqueued.
pub fn emit_priority_queue_count(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_len(chunks, current, line);
}

/// `pq.Clear()`.
pub fn emit_priority_queue_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_clear(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

// ── NameValueCollection ───────────────────────────────────────────────────

/// `nvc.Add(name, value)` — stack `[recv, name, value]`.
///
/// ⛔ NOT `Map.set`. §A NameValueCollection holds MULTIPLE values per name and
/// `Add` APPENDS: adding `val1` then `val2` under one name makes the indexer
/// read `"val1,val2"`. Aliasing this to `set` overwrote instead, which reads as
/// a plausible answer (`"val2"`) rather than an error.
pub fn emit_name_value_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash(chunks, current, 3, line);
    let (recv, name, value) = (base, base + 1, base + 2);
    let existing = chunks[current].alloc_scratch(1);

    get(chunks, current, recv, line);
    get(chunks, current, name, line);
    let map_get = chunks[current].add_import("ecma:map", "get");
    chunks[current].emit_call(map_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, existing, line);

    get(chunks, current, recv, line);
    get(chunks, current, name, line);

    get(chunks, current, existing, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    get(chunks, current, value, line);
    chunks[current].emit_else(line);
    get(chunks, current, existing, line);
    chunks[current].emit_string_const(",", line);
    ops::emit_dyn_add(&mut chunks[current], line);
    get(chunks, current, value, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_end(line);

    let map_set = chunks[current].add_import("ecma:map", "set");
    chunks[current].emit_call(map_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// A fresh empty pair array — the backing store for `OrderedDictionary`.
pub fn emit_seq_new(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
}
