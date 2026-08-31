//! `System.Threading.Tasks.Parallel`, `ParallelOptions`, `AsyncLocal(Of T)` and
//! `System.Collections.Concurrent.Partitioner` — the data-parallel surface every
//! .NET language shares.
//!
//! ⛔ **SEQUENTIAL IS A CONFORMANT SCHEDULE, NOT A STAND-IN.** `Parallel.For`
//! promises that the body runs once per index; it never promises an order, a
//! thread, or an overlap, and `MaxDegreeOfParallelism = 1` is a legal .NET
//! configuration for every one of these calls. Running the body in index order
//! on the calling thread satisfies the contract exactly.
//!
//! The runtime this compiles to has no thread source to do otherwise: the WASM
//! threads opcodes exist, but nothing spawns, and `primitives/channels.rs`
//! records the same boundary for Go's blocking channel ops. A `worker_threads`
//! host would not help either — the bodies here close over CALLER LOCALS
//! (`lock (obj) { sum += k; }`), and a worker cannot share that frame.
//!
//! `lock` needs nothing from this module: it already compiles on any object.

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::{collections, delegates, ops};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use super::object_fields::field_slot;

const VALUE: &str = "Value";
const MAX_DOP: &str = "MaxDegreeOfParallelism";
const IS_COMPLETED: &str = "IsCompleted";
const LOWEST_BREAK: &str = "LowestBreakIteration";
const RANGE_FROM: &str = "__from";
const RANGE_TO: &str = "__to";
const RANGE_SIZE: &str = "__range_size";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn field_set_drop(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

fn field_get(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_get(chunk, ObjSource::Stack, &field_slot(key), Dest::Stack, line);
}

fn stash_args(chunk: &mut Chunk, argc: u8, line: u32) -> u16 {
    let base = chunk.alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// `New AsyncLocal(Of T)()` — a one-slot ambient context.
///
/// The value flows with the logical call context, and on a single flow of
/// execution that context is the object itself. The `valueChangedHandler`
/// overload is accepted and dropped: it fires when the context SWITCHES, which
/// one flow never does.
///
/// Stack: `[args…]` → `[asyncLocal]`.
pub fn emit_async_local_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    set(chunk, obj, line);
    get(chunk, obj, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    field_set_drop(chunk, VALUE, line);
    get(chunk, obj, line);
}

/// `New ParallelOptions()` — the knobs `Parallel.*` reads.
///
/// `MaxDegreeOfParallelism` defaults to `-1` ("unbounded"), which is what .NET
/// reports on a fresh instance; an object initializer overwrites it as an
/// ordinary field write.
///
/// Stack: `[]` → `[options]`.
pub fn emit_parallel_options_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    set(chunk, obj, line);
    get(chunk, obj, line);
    chunk.emit_f64_const(-1.0, line);
    field_set_drop(chunk, MAX_DOP, line);
    get(chunk, obj, line);
}

/// A completed `ParallelLoopResult`, left on the stack.
///
/// .NET returns one from every `Parallel.For`/`ForEach`. A loop that ran to
/// completion reports `IsCompleted = True` and a null `LowestBreakIteration`.
fn emit_loop_result(chunk: &mut Chunk, line: u32) {
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    set(chunk, obj, line);
    get(chunk, obj, line);
    chunk.emit_bool_const(true, line);
    field_set_drop(chunk, IS_COMPLETED, line);
    get(chunk, obj, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    field_set_drop(chunk, LOWEST_BREAK, line);
    get(chunk, obj, line);
}

/// Call `body(arg)` and discard whatever it answers.
fn emit_call_body(chunks: &mut [Chunk], current: usize, body: u16, arg: u16, line: u32) {
    get(&mut chunks[current], body, line);
    get(&mut chunks[current], arg, line);
    delegates::emit_invoke(chunks, current, 2, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `Parallel.For(fromInclusive, toExclusive, body)` and its `ParallelOptions`
/// overload.
///
/// Stack: `[from, to, body]` (or `[from, to, options, body]`) →
/// `[ParallelLoopResult]`.
pub fn emit_parallel_for(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(&mut chunks[current], argc, line);
    let from = base;
    let to = base + 1;
    // The body is always LAST; an options argument sits between the bounds and
    // it, and selects a degree of parallelism this schedule already satisfies.
    let body = base + argc as u16 - 1;

    let idx = chunks[current].alloc_scratch(1);
    get(&mut chunks[current], from, line);
    set(&mut chunks[current], idx, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], idx, line);
    get(&mut chunks[current], to, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    emit_call_body(chunks, current, body, idx, line);

    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    emit_loop_result(&mut chunks[current], line);
}

/// `Parallel.ForEach(source, body)` and its `ParallelOptions` overload, plus
/// `Parallel.ForEachAsync` — which differs only in awaiting each body, and the
/// bodies here are already run to completion before the next index starts.
///
/// Stack: `[source, body]` (or `[source, options, body]`) →
/// `[ParallelLoopResult]`.
pub fn emit_parallel_for_each(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(&mut chunks[current], argc, line);
    let source = base;
    let body = base + argc as u16 - 1;

    let idx = chunks[current].alloc_scratch(3);
    let len = idx + 1;
    let elem = idx + 2;

    get(&mut chunks[current], source, line);
    collections::emit_len(chunks, current, line);
    set(&mut chunks[current], len, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    set(&mut chunks[current], idx, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], idx, line);
    get(&mut chunks[current], len, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    get(&mut chunks[current], source, line);
    get(&mut chunks[current], idx, line);
    collections::emit_get(chunks, current, line);
    set(&mut chunks[current], elem, line);
    emit_call_body(chunks, current, body, elem, line);

    get(&mut chunks[current], idx, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], idx, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    emit_loop_result(&mut chunks[current], line);
}

/// `Parallel.Invoke(action1, action2, …)` — run each action once.
///
/// Stack: `[a1 … aN]` → `[]`.
pub fn emit_parallel_invoke(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(&mut chunks[current], argc, line);
    for offset in 0..argc as u16 {
        get(&mut chunks[current], base + offset, line);
        delegates::emit_invoke(chunks, current, 1, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `Partitioner.Create(fromInclusive, toExclusive, rangeSize)`.
///
/// The range form is the one the corpus builds. The bounds are kept so the
/// partitions can be COMPUTED on demand rather than guessed at.
///
/// Stack: `[from, to, rangeSize]` → `[partitioner]`.
pub fn emit_partitioner_create(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(&mut chunks[current], argc, line);
    let chunk = &mut chunks[current];
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    set(chunk, obj, line);

    get(chunk, obj, line);
    get(chunk, base, line);
    field_set_drop(chunk, RANGE_FROM, line);

    get(chunk, obj, line);
    if argc >= 2 {
        get(chunk, base + 1, line);
    } else {
        chunk.emit_f64_const(0.0, line);
    }
    field_set_drop(chunk, RANGE_TO, line);

    get(chunk, obj, line);
    if argc >= 3 {
        get(chunk, base + 2, line);
    } else {
        chunk.emit_f64_const(1.0, line);
    }
    field_set_drop(chunk, RANGE_SIZE, line);

    get(chunk, obj, line);
}

/// `partitioner.GetPartitions(n)` / `GetOrderablePartitions(n)`.
///
/// Both answer the SAME ranges here — the "orderable" form differs only by
/// carrying each range's index, and the ranges are yielded in order — so the
/// partitioning is computed once: `[from, from+size)`, `[from+size, …)`, up to
/// `to`, exactly as .NET's range partitioner splits.
///
/// Stack: `[partitioner, n]` → `[array of [from, to] pairs]`.
pub fn emit_partitioner_get_partitions(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_args(&mut chunks[current], argc, line);
    let recv = base;

    let out = chunks[current].alloc_scratch(4);
    let cursor = out + 1;
    let stop = out + 2;
    let next = out + 3;

    collections::emit_array_new(chunks, current, 0, line);
    set(&mut chunks[current], out, line);

    get(&mut chunks[current], recv, line);
    field_get(&mut chunks[current], RANGE_FROM, line);
    set(&mut chunks[current], cursor, line);
    get(&mut chunks[current], recv, line);
    field_get(&mut chunks[current], RANGE_TO, line);
    set(&mut chunks[current], stop, line);

    let outer = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    get(&mut chunks[current], cursor, line);
    get(&mut chunks[current], stop, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    // next = min(cursor + rangeSize, stop)
    get(&mut chunks[current], cursor, line);
    get(&mut chunks[current], recv, line);
    field_get(&mut chunks[current], RANGE_SIZE, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    set(&mut chunks[current], next, line);
    get(&mut chunks[current], next, line);
    get(&mut chunks[current], stop, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    get(&mut chunks[current], next, line);
    chunks[current].emit_else(line);
    get(&mut chunks[current], stop, line);
    chunks[current].emit_end(line);
    set(&mut chunks[current], next, line);

    get(&mut chunks[current], out, line);
    get(&mut chunks[current], cursor, line);
    get(&mut chunks[current], next, line);
    collections::emit_array_new(chunks, current, 2, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    get(&mut chunks[current], next, line);
    set(&mut chunks[current], cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer);

    get(&mut chunks[current], out, line);
}
