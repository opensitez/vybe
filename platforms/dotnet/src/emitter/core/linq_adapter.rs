//! .NET `System.Linq` instance-method adapter — bytecode-only.
//!
//! Every .NET-shape language (C#, VB, F#, …) ships the same LINQ
//! surface on `IEnumerable<T>` / `List<T>` / arrays. Each adapter
//! emits composed bytecode that ECMA-shape array opcodes already
//! deliver, so VB and C# get one implementation regardless of how
//! the surface syntax differs.
//!
//! Each `emit_linq_*` is invoked through `value_methods` dispatch.
//! Stack on entry is `[receiver, arg1, ..., argN]` (per the
//! `compile_call` value-method contract); each emitter leaves a
//! single result on the stack.
//!
//! Pure WASM, no `vybe:*` involvement. Composes existing
//! `compiler_common::collections` / `compiler_common::loops`
//! emitters wherever possible so semantics stay aligned with the
//! rest of the standard library.

use vybe_bytecode::opcode::Op;
use vybe_bytecode::Chunk;
use vybe_emitter::instructions::core_wasm;

use vybe_emitter::collections;
use vybe_emitter::generators;
use vybe_emitter::loops;
use vybe_emitter::ops;

/// Allocate `count` consecutive scratch locals; returns the first slot.
fn alloc_locals(chunk: &mut Chunk, count: u16) -> u16 {
    chunk.alloc_scratch(count)
}

/// Drain the receiver in `slot` through the shared ECMA §7.4 iterator protocol
/// (`emit_spread_iterable`) and store the materialized array back. This is what
/// lets every LINQ operator work uniformly over arrays, `List<T>`, `yield`
/// generators, custom `IEnumerable`, and — because the drain is the common
/// cross-language iterator emitter — iterables produced by any Vybe frontend.
/// For an array receiver it is effectively identity.
fn materialize_receiver_slot(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Bounded counterpart to `materialize_receiver_slot`: for lazy operators
/// (`Take`, `First`, …) that only need a prefix of the sequence.
///
/// A **generator** receiver is drained LAZILY to at most `limit_slot` elements
/// via the bounded stack-switching take (`generators::emit_take_into_array`),
/// so it terminates on infinite sequences — C#'s deferred-execution semantics.
/// Any other iterable is finite by construction and materialized in full; the
/// caller slices to `limit`. The bounded array is stored back into `slot`.
fn materialize_bounded_slot(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    limit_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_gen = chunks[current].add_import("ecma:value", "isGenerator");
    chunks[current].emit_call(is_gen, 1, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    // Generator → bounded lazy drain (safe for infinite sequences).
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, limit_slot, line);
    generators::emit_take_into_array(chunks, current, line);

    chunks[current].emit_else(line);

    // Other iterable → full materialization (caller slices to `limit`).
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_spread_iterable(chunks, current, line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_import_call(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    // Per-chunk import tables: the import must be registered on the SAME chunk
    // that emits the CALL_IMPORT, or the index is out of range when the adapter
    // runs in a non-script chunk (e.g. compiled into a `__linq_*` vtable chunk).
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

// ── Pure reductions (no fn arg) ──────────────────────────────────────────

/// `arr.First()` — returns `arr[0]`. Stack: [arr] → [first].
/// Bounded: a generator only advances once (deferred execution).
pub fn emit_linq_first(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 2);
    let one_slot = arr_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, one_slot, line);
    materialize_bounded_slot(chunks, current, arr_slot, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
}

/// `arr.Last()` — returns `arr[arr.length - 1]`. Stack: [arr] → [last].
pub fn emit_linq_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
}

/// `arr.Skip(n)` — `arr.slice(n, arr.length)`. Stack: [arr, n] → [array].
pub fn emit_linq_skip(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 2);
    let n_slot = arr_slot + 1;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, n_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, n_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);
}

/// `arr.Take(n)` — `arr.slice(0, n)`. Stack: [arr, n] → [array].
pub fn emit_linq_take(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 2);
    let n_slot = arr_slot + 1;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, n_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    // Bounded drain: a generator yields at most `n` elements (deferred
    // execution — terminates on infinite sequences); other iterables
    // materialize fully and are sliced below.
    materialize_bounded_slot(chunks, current, arr_slot, n_slot, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, n_slot, line);
    collections::emit_slice(chunks, current, line);
}

/// `arr.ToList()` / `arr.ToArray()` — materialize the sequence into a concrete
/// array (draining generators / custom iterables). Stack: [seq] → [array].
pub fn emit_linq_identity(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_spread_iterable(chunks, current, line);
}

/// `arr.Average()` — `Sum(arr) / Length(arr)`. Stack: [arr] → [number].
pub fn emit_linq_average(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_sum(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

/// `arr.Sum()` — materialize then sum. Stack: [seq] → [number].
pub fn emit_linq_sum(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_sum(chunks, current, line);
}

/// `arr.Count()` — materialize then length. Stack: [seq] → [number].
pub fn emit_linq_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
}

/// `arr.FirstOrDefault()` — `arr.length === 0 ? default : arr[0]`.
/// `default` is `0` (numeric) — generalising requires a per-T hook.
/// Stack: [arr] → [first | 0].
pub fn emit_linq_first_or_default(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 2);
    let one_slot = arr_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, one_slot, line);
    materialize_bounded_slot(chunks, current, arr_slot, one_slot, line);

    // Test arr.length == 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);

    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

/// `arr.Distinct()` — array with duplicates removed.
/// Stack: [arr] → [array].
///
/// Lowers to:
/// ```text
/// result = [];
/// for elem in arr {
///     if result.indexOf(elem) < 0 { result.push(elem); }
/// }
/// ```
pub fn emit_linq_distinct(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 4);
    let result_slot = arr_slot + 1;
    let idx_slot = arr_slot + 2;
    let elem_slot = arr_slot + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    // result = []
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    // for_in_start leaves arr[i] on the stack — stash to elem_slot.
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // result.indexOf(elem) — push elem only if it is NOT already in result.
    // Structured WASM block matches the pattern used by `emit_filter`.
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_index_of(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    let if_block = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line); // skip push if duplicate (>= 0)
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line); // discard push's return value
    chunks[current].emit_end(line);
    chunks[current].patch_block(if_block);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.DistinctBy(keyFn)` — first element for each distinct `keyFn(elem)`.
/// Stack: [arr, keyFn] → [array]. Dedupes on the projected key (tracked in a
/// separate `keys` array) while emitting the original elements.
pub fn emit_linq_distinct_by(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 7);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let keys_slot = arr_slot + 3;
    let idx_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;
    let key_slot = arr_slot + 6;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    // result = []; keys = []
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // key = keyFn(elem)
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    // if keys.indexOf(key) < 0 { keys.push(key); result.push(elem); }
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_index_of(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    let if_block = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line); // skip if key already seen
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(if_block);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.SequenceEqual(other)` — same length and pairwise equal values.
/// Stack: [arr, other] → [bool].
pub fn emit_linq_sequence_equal(chunks: &mut [Chunk], current: usize, line: u32) {
    let left_slot = alloc_locals(&mut chunks[current], 6);
    let right_slot = left_slot + 1;
    let len_slot = left_slot + 2;
    let idx_slot = left_slot + 3;
    let right_elem_slot = left_slot + 4;
    let result_slot = left_slot + 5;

    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);

    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    let done = chunks[current].emit_block(line);
    let lengths_match = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(1, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(lengths_match);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_elem_slot, line);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    let equal_values = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);

    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(2, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(equal_values);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    chunks[current].emit_end(line);
    chunks[current].patch_block(done);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

// ── HOFs (one fn arg) ────────────────────────────────────────────────────
//
// Each emitter receives the receiver and predicate / selector / reducer
// already on the stack (per the value_methods dispatch contract). The
// emitter does the per-element CALL_REF inline.

/// `arr.Count(pred)` — count elements where `pred(elem)` is truthy.
/// Stack: [arr, pred] → [count].
///
/// `arr.Count()` (0-arg) defers to the runtime collection registry
/// (List<T>.Count is a per-type property). The 1-arg form is the
/// LINQ overload — `compiler/calls.rs::compile_call`'s
/// `prefer_dotnet_adapter` check routes any `common:dotnet.*`
/// value-method overload around the registry intercept so this
/// emitter actually runs.
pub fn emit_linq_count_pred(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = arr_slot + 1;
    let count_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    let chunk = &mut chunks[current];
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // pred(elem) → if-true increment count.  Use a structured WASM
    // block (same pattern as `emit_filter`) — byte-offset
    // Structured skip blocks keep the predicate guard interleaved with the
    // outer `for_in` body block, so we open an inner block and `br_if`
    // out of it on the false branch.
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let if_block = chunks[current].emit_block(line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // skip increment if false
    // count++
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(if_block);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
}

/// `arr.Where(pred)` — elements for which `pred(elem)` is truthy.
/// Stack: [arr, pred] → [array].
pub fn emit_linq_where(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // if pred(elem): result.push(elem)  (structured skip block, cf. count_pred)
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let if_block = chunks[current].emit_block(line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // skip push if pred false
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(if_block);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.Any()` — true iff the sequence has any elements. Emits a proper
/// boxed bool (`i32_to_bool` of the length) so it prints `True`/`False`.
/// Stack: [seq] → [bool].
pub fn emit_linq_any(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    vybe_emitter::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `arr.Contains(x)` — `arr.includes(x)`. Stack: [arr, x] → [bool].
pub fn emit_linq_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 2);
    let x_slot = arr_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, x_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x_slot, line);
    collections::emit_contains(chunks, current, line);
}

/// `arr.Reverse()` — a new reversed array (LINQ Reverse is non-mutating).
/// Stack: [seq] → [array].
pub fn emit_linq_reverse(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_reversed(chunks, current, line);
}

/// `arr.SkipWhile(pred)` — drop leading elements while `pred` holds, keep the
/// rest (including the first element that fails `pred`). Stack: [arr, pred] →
/// [array]. A `skipping` flag is cleared at the first failing element.
pub fn emit_linq_skip_while(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 6);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let skipping_slot = arr_slot + 5;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_i32_const(1, line); // skipping = true
    chunks[current].emit_op_u16(Op::LOCAL_SET, skipping_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // if skipping && !pred(elem): skipping = false
    let stop_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, skipping_slot, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // not skipping → leave flag
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // pred still true → keep skipping
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, skipping_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(stop_block);

    // if !skipping: result.push(elem)
    let push_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, skipping_slot, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // still skipping → no push
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(push_block);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.TakeWhile(pred)` — keep leading elements while `pred` holds, stop at
/// the first failing element. Stack: [arr, pred] → [array]. A `taking` flag is
/// cleared at the first failing element and suppresses all later pushes.
pub fn emit_linq_take_while(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 6);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let taking_slot = arr_slot + 5;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_i32_const(1, line); // taking = true
    chunks[current].emit_op_u16(Op::LOCAL_SET, taking_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // if taking && !pred(elem): taking = false
    let stop_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, taking_slot, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // already stopped → leave flag
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // pred true → keep taking
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, taking_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(stop_block);

    // if taking: result.push(elem)
    let push_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, taking_slot, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // stopped → no push
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(push_block);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.Chunk(size)` — split into consecutive sub-arrays of length `size`
/// (the final batch may be shorter). Stack: [arr, size] → [array of arrays].
pub fn emit_linq_chunk(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 6);
    let size_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let batch_slot = arr_slot + 3;
    let idx_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;

    chunks[current].emit_op_u16(Op::LOCAL_SET, size_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, batch_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // batch.push(elem)
    chunks[current].emit_op_u16(Op::LOCAL_GET, batch_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // if batch.length >= size: result.push(batch); batch = []
    let flush_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, batch_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, size_slot, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // batch not full yet → keep filling
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, batch_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, batch_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(flush_block);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    // trailing partial batch: if batch.length >= 1: result.push(batch)
    let tail_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, batch_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(1, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // empty → nothing to append
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, batch_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(tail_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.Aggregate(seed, fn)` — fold from `seed` calling `fn(acc, x)`.
/// .NET argument order is `(seed, fn)`; we swap to call the shared
/// `emit_reduce` helper which expects `acc` already initialised.
/// Stack: [arr, seed, fn] → [acc].
pub fn emit_linq_aggregate(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 4);
    let fn_slot = arr_slot + 1;
    let acc_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    // for elem in arr: acc = fn(acc, elem)
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    {
        let chunk = &mut chunks[current];
        let elem_local = chunk.alloc_scratch(1);
        chunk.emit_op_u16(Op::LOCAL_SET, elem_local, line);

        chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, acc_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, elem_local, line);
        chunk.emit_op_u8(Op::CALL_REF, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, acc_slot, line);
    }
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_slot, line);
}

/// `arr.OrderByDescending(keyFn)` — same as `OrderBy` then `reverse`.
/// Stack: [arr, keyFn] → [array].
pub fn emit_linq_order_by_descending(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_sort_by_key_in_place(chunks, current, line);
    collections::emit_reverse(chunks, current, line);
}

/// `arr.OrderBy(keyFn)` — ascending stable sort by projected key.
/// Stack: [arr, keyFn] → [sorted array].
pub fn emit_linq_order_by(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_sort_by_key_in_place(chunks, current, line);
}

/// `arr.Select(fn)` — invoke `map` on the receiver.
/// Stack: [arr, fn] → [array].
pub fn emit_linq_select(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // Normalize receiver to an indexable values array.
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // result = []
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // result.push(fn(elem))
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.SelectMany(fn)` — invoke `flatMap` on the receiver.
/// Stack: [arr, fn] → [array].
pub fn emit_linq_select_many(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 6);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let mapped_slot = arr_slot + 5;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // Normalize receiver to an indexable values array.
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // result = []
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // mapped = fn(elem)
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_slot, line);

    // result = result.concat(mapped)
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    collections::emit_concat(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.GroupBy(keyFn)` — returns an array of grouping objects.
/// Each group has `Key`, `Items`, and `Count` properties.
/// Stack: [arr, keyFn] → [groups].
pub fn emit_linq_group_by(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 9);
    let fn_slot = arr_slot + 1;
    let map_slot = arr_slot + 2;
    let out_slot = arr_slot + 3;
    let idx_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;
    let key_slot = arr_slot + 6;
    let group_slot = arr_slot + 7;
    let items_slot = arr_slot + 8;

    // Stack: [arr, keyFn]
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // groupMap = new Map()
    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_slot, line);

    // out = []
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // key = keyFn(elem)
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    // if !groupMap.has(key) { create group object, initialize fields, save map, out.push(group) }
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    emit_import_call(chunks, current, "ecma:map", "has", 2, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let maybe_new = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line); // already exists

    emit_import_call(chunks, current, "ecma:object", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, group_slot, line);

    // group["Key"] = key
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Key", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // group["Items"] = []
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Items", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // group["Count"] = 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Count", line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // groupMap[key] = group
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // out.push(group)
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(maybe_new);

    // group = groupMap[key]
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, group_slot, line);

    // items = group["Items"]
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Items", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_slot, line);

    // items.push(elem)
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // group["Count"] = items.length
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Count", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_slot, line);
    collections::emit_len(chunks, current, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// `arr.ToDictionary(keyFn, valueFn)`.
/// Stack: [arr, keyFn, valueFn] → [map].
pub fn emit_linq_to_dictionary(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 8);
    let key_fn_slot = arr_slot + 1;
    let val_fn_slot = arr_slot + 2;
    let map_slot = arr_slot + 3;
    let idx_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;
    let key_slot = arr_slot + 6;
    let val_slot = arr_slot + 7;

    chunks[current].emit_op_u16(Op::LOCAL_SET, val_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, val_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, val_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
}

/// `arr.Zip(other, zipperFn)`.
/// Stack: [arr, other, fn] → [array].
pub fn emit_linq_zip(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 8);
    let other_slot = arr_slot + 1;
    let fn_slot = arr_slot + 2;
    let out_slot = arr_slot + 3;
    let idx_slot = arr_slot + 4;
    let left_slot = arr_slot + 5;
    let right_slot = arr_slot + 6;
    let zipped_slot = arr_slot + 7;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);

    // Skip iteration body if idx >= other.length
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    collections::emit_len(chunks, current, line);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    let too_short = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, zipped_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, zipped_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(too_short);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

// ── ElementAt / Single / defaults ────────────────────────────────────────

/// `arr.ElementAt(i)` — `arr[i]`, throwing `ArgumentOutOfRangeException` when
/// `i` is out of range. Stack: [arr, i] → [elem].
pub fn emit_linq_element_at(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = alloc_locals(&mut chunks[current], 2);
    let arr_slot = base;
    let idx_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    // if !(i >= 0 && i < len) throw
    emit_index_in_range(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Index was out of range. Must be non-negative and less than the size of the collection.", line);
    vybe_emitter::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "ArgumentOutOfRangeException",
        line,
    );
    vybe_emitter::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
}

/// `arr.ElementAtOrDefault(i)` — `arr[i]` when in range, else `default` (`0`).
/// Stack: [arr, i] → [elem | 0].
pub fn emit_linq_element_at_or_default(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = alloc_locals(&mut chunks[current], 2);
    let arr_slot = base;
    let idx_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    emit_index_in_range(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

/// Leaves an i32 bool: `idx >= 0 && idx < arr.length`.
fn emit_index_in_range(
    chunks: &mut [Chunk],
    current: usize,
    arr_slot: u16,
    idx_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
}

/// `arr.Single()` — the sole element, throwing `InvalidOperationException`
/// unless the sequence has exactly one. Stack: [arr] → [elem].
pub fn emit_linq_single(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // if len != 1 throw
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_ne(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_string_const("Sequence contains no elements or more than one element.", line);
    vybe_emitter::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "InvalidOperationException",
        line,
    );
    vybe_emitter::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
}

/// `arr.SingleOrDefault()` — the sole element when the sequence has exactly
/// one, else `default` (`0`). Stack: [arr] → [elem | 0].
pub fn emit_linq_single_or_default(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_end(line);
}

// ── MaxBy / MinBy ────────────────────────────────────────────────────────

/// `arr.MaxBy(keyFn)` — element whose `keyFn(elem)` is greatest (first on ties).
/// Stack: [arr, keyFn] → [elem].
pub fn emit_linq_max_by(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_by_extreme(chunks, current, line, true);
}

/// `arr.MinBy(keyFn)` — element whose `keyFn(elem)` is smallest (first on ties).
/// Stack: [arr, keyFn] → [elem].
pub fn emit_linq_min_by(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_by_extreme(chunks, current, line, false);
}

fn emit_linq_by_extreme(chunks: &mut [Chunk], current: usize, line: u32, want_max: bool) {
    let base = alloc_locals(&mut chunks[current], 7);
    let arr_slot = base;
    let fn_slot = base + 1;
    let best_slot = base + 2;
    let bestkey_slot = base + 3;
    let idx_slot = base + 4;
    let elem_slot = base + 5;
    let key_slot = base + 6;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // Normalize receiver to an indexable values array.
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // best = arr[0]; bestKey = fn(best)
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bestkey_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // key = fn(elem)
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    // if key >/< bestKey { best = elem; bestKey = key }
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bestkey_slot, line);
    if want_max {
        vybe_emitter::ops::emit_dyn_gt(&mut chunks[current], line);
    } else {
        vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    }
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, best_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bestkey_slot, line);
    chunks[current].emit_end(line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, best_slot, line);
}

// ── Aggregate without seed / Append / Prepend ────────────────────────────

/// `arr.Aggregate(fn)` (no seed) — fold starting from `arr[0]`.
/// Stack: [arr, fn] → [acc].
pub fn emit_linq_aggregate_no_seed(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = alloc_locals(&mut chunks[current], 4);
    let arr_slot = base;
    let fn_slot = base + 1;
    let acc_slot = base + 2;
    let idx_slot = base + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // Normalize receiver to an indexable values array.
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // acc = arr[0]
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc_slot, line);

    // rest = arr.slice(1, len)
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // for elem in rest: acc = fn(acc, elem)
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    {
        let elem_local = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_SET, elem_local, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, acc_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem_local, line);
        chunks[current].emit_op_u8(Op::CALL_REF, 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc_slot, line);
    }
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, acc_slot, line);
}

/// `arr.Append(x)` — new sequence `[...arr, x]`. Stack: [arr, x] → [array].
pub fn emit_linq_append(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = alloc_locals(&mut chunks[current], 2);
    let arr_slot = base;
    let x_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, x_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // [...arr].concat([x])
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x_slot, line);
    collections::emit_array_new(chunks, current, 1, line);
    collections::emit_concat(chunks, current, line);
}

/// `arr.Prepend(x)` — new sequence `[x, ...arr]`. Stack: [arr, x] → [array].
pub fn emit_linq_prepend(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = alloc_locals(&mut chunks[current], 2);
    let arr_slot = base;
    let x_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, x_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // [x].concat([...arr])
    chunks[current].emit_op_u16(Op::LOCAL_GET, x_slot, line);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    collections::emit_concat(chunks, current, line);
}

// ── SkipLast / TakeLast / DefaultIfEmpty ─────────────────────────────────

/// Leaves an f64 on the stack: `max(0, arr.length - n)`.
fn emit_len_minus_n_clamped(chunks: &mut [Chunk], current: usize, arr_slot: u16, n_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::F64_FROM_I32, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n_slot, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_f64_const(0.0, line);
    chunks[current].emit_op(Op::F64_MAX, line);
}

/// `arr.SkipLast(n)` — `arr.slice(0, max(0, len - n))`. Stack: [arr, n] → [array].
pub fn emit_linq_skip_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = alloc_locals(&mut chunks[current], 2);
    let arr_slot = base;
    let n_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    emit_len_minus_n_clamped(chunks, current, arr_slot, n_slot, line);
    collections::emit_slice(chunks, current, line);
}

/// `arr.TakeLast(n)` — `arr.slice(max(0, len - n), len)`. Stack: [arr, n] → [array].
pub fn emit_linq_take_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = alloc_locals(&mut chunks[current], 2);
    let arr_slot = base;
    let n_slot = base + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, n_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    emit_len_minus_n_clamped(chunks, current, arr_slot, n_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    collections::emit_slice(chunks, current, line);
}

/// `arr.DefaultIfEmpty()` — `arr` when non-empty, else `[default]` (`[0]`).
/// Stack: [arr] → [array].
pub fn emit_linq_default_if_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_end(line);
}
