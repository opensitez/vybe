//! `System.Collections.Immutable` emits — copy-on-write over ordinary storage.
//!
//! Every "mutation" here is COPY-then-mutate and returns the copy. That is the
//! entire contract of these types, and it is why they cannot simply alias the
//! mutable leaves: `dotnet.list_add` pushes in place, so `a2 = a1.Add(2)` would
//! leave `a1.Length == 2` when .NET guarantees 1.
//!
//! The copy is eager, not structural sharing. A persistent list in .NET shares
//! its spine; here it does not, so `Add` is O(n) rather than O(log n). That is
//! a performance difference, never an observable one — no program can tell the
//! two apart through the public surface — and buying the real structure would
//! mean a second array representation the rest of the compiler cannot index.

use vybe_compiler::primitives::{collections, sets};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// Pop `argc` values into consecutive scratch slots, returning the base slot.
///
/// The stack holds them in call order, so they are popped in REVERSE — the
/// same shape `collections_adapter::stash_args` uses. Kept private rather than
/// shared because a scratch base is only meaningful inside one emit.
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

// ── sequences: ImmutableArray / ImmutableList / Queue / Stack ─────────────

/// `ImmutableX.Empty` — a fresh empty sequence.
///
/// A fresh one per read, not a shared singleton. `Empty` is only ever the
/// START of a copy chain, and handing out one shared array would make two
/// unrelated `Empty.Add(…)` chains able to observe each other if any future
/// leaf ever mutated in place.
pub fn emit_seq_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_array_new(chunks, current, 0, line);
}

/// `ImmutableX.Create(a, b, …)` — the arguments are already on the stack.
pub fn emit_seq_create(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    collections::emit_array_new(chunks, current, argc as u16, line);
}

/// `seq.Add(v)` / `queue.Enqueue(v)` / `stack.Push(v)` — one append, three
/// .NET names.
pub fn emit_seq_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash(chunks, current, 2, line);
    let (recv, value) = (base, base + 1);
    let copy = chunks[current].alloc_scratch(1);

    get(chunks, current, recv, line);
    collections::emit_clone(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, copy, line);

    get(chunks, current, copy, line);
    get(chunks, current, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line); // push answers the new length

    get(chunks, current, copy, line);
}

/// `seq.AddRange(other)` — concat already answers a NEW array, so no copy.
pub fn emit_seq_add_range(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_concat(chunks, current, line);
}

/// `seq.SetItem(i, v)`.
pub fn emit_seq_set_item(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash(chunks, current, 3, line);
    let (recv, index, value) = (base, base + 1, base + 2);
    let copy = chunks[current].alloc_scratch(1);

    get(chunks, current, recv, line);
    collections::emit_clone(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, copy, line);

    get(chunks, current, copy, line);
    get(chunks, current, index, line);
    get(chunks, current, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(chunks, current, copy, line);
}

/// `seq.RemoveAt(i)` — `[..i] ++ [i+1..]`, which never touches the receiver.
pub fn emit_seq_remove_at(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash(chunks, current, 2, line);
    let (recv, index) = (base, base + 1);

    get(chunks, current, recv, line);
    chunks[current].emit_i32_const(0, line);
    get(chunks, current, index, line);
    collections::emit_slice(chunks, current, line);

    get(chunks, current, recv, line);
    get(chunks, current, index, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_i32_const(i32::MAX, line);
    collections::emit_slice(chunks, current, line);

    collections::emit_concat(chunks, current, line);
}

/// `seq.IsEmpty` — length == 0. Also serves the queue and the stack.
pub fn emit_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
}

/// `queue.Peek()` — the FRONT. A queue and a stack share storage and differ
/// only in which end they read, which is the whole reason both exist here.
pub fn emit_queue_peek(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
}

/// `queue.Dequeue()` — everything after the front.
pub fn emit_queue_dequeue(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_i32_const(i32::MAX, line);
    collections::emit_slice(chunks, current, line);
}

/// `stack.Peek()` — the BACK.
pub fn emit_stack_peek(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash(chunks, current, 1, line);
    get(chunks, current, recv, line);
    get(chunks, current, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
}

/// `stack.Pop()` — everything before the back.
pub fn emit_stack_pop(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = stash(chunks, current, 1, line);
    get(chunks, current, recv, line);
    chunks[current].emit_i32_const(0, line);
    get(chunks, current, recv, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_slice(chunks, current, line);
}

// ── ImmutableDictionary ───────────────────────────────────────────────────

pub fn emit_map_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_map_new(chunks, current, line);
}

/// `map.Add(k, v)` / `map.SetItem(k, v)`.
///
/// .NET's `Add` throws on a duplicate key and `SetItem` overwrites. Both are
/// the same emit here because the copy is what these tests observe and a
/// duplicate-key throw needs a `ContainsKey` probe this surface does not yet
/// declare — noted rather than faked, so the gap is visible.
pub fn emit_map_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash(chunks, current, 3, line);
    let (recv, key, value) = (base, base + 1, base + 2);
    let copy = chunks[current].alloc_scratch(1);

    get(chunks, current, recv, line);
    collections::emit_map_clone(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, copy, line);

    get(chunks, current, copy, line);
    get(chunks, current, key, line);
    get(chunks, current, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(chunks, current, copy, line);
}

// ── ImmutableHashSet ──────────────────────────────────────────────────────

pub fn emit_set_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    sets::emit_new(chunks, current, line);
}

pub fn emit_set_create(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    sets::emit_literal(chunks, current, argc, line);
}

/// `set.Add(v)` — copy, then add to the copy.
pub fn emit_set_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash(chunks, current, 2, line);
    let (recv, value) = (base, base + 1);
    let copy = chunks[current].alloc_scratch(1);

    get(chunks, current, recv, line);
    sets::emit_from_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, copy, line);

    get(chunks, current, copy, line);
    get(chunks, current, value, line);
    sets::emit_add(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(chunks, current, copy, line);
}

pub fn emit_set_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = stash(chunks, current, 2, line);
    let (recv, value) = (base, base + 1);
    let copy = chunks[current].alloc_scratch(1);

    get(chunks, current, recv, line);
    sets::emit_from_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, copy, line);

    get(chunks, current, copy, line);
    get(chunks, current, value, line);
    sets::emit_delete(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(chunks, current, copy, line);
}

/// `Union` / `Intersect` / `Except` — the ECMA set operations already ANSWER a
/// new set, unlike their `…With` mutating cousins that `HashSet` uses. That
/// distinction is exactly the immutable/mutable split, so these need no copy.
pub fn emit_set_union(chunks: &mut [Chunk], current: usize, line: u32) {
    sets::emit_union(chunks, current, line);
}

pub fn emit_set_intersect(chunks: &mut [Chunk], current: usize, line: u32) {
    sets::emit_intersection(chunks, current, line);
}

pub fn emit_set_except(chunks: &mut [Chunk], current: usize, line: u32) {
    sets::emit_difference(chunks, current, line);
}

pub fn emit_set_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    sets::emit_size(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
}
