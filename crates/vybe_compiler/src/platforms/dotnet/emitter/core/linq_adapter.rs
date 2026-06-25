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

use crate::emitter::instructions::core_wasm;
use std::sync::Arc;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::emitter::collections;
use crate::emitter::loops;

/// Allocate `count` consecutive scratch locals; returns the first slot.
fn alloc_locals(chunk: &mut Chunk, count: u16) -> u16 {
    chunk.alloc_scratch(count)
}

fn emit_import_call(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[0].add_import(module, name);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(argc, line);
}

// ── Pure reductions (no fn arg) ──────────────────────────────────────────

/// `arr.First()` — returns `arr[0]`. Stack: [arr] → [first].
pub fn emit_linq_first(chunks: &mut [Chunk], current: usize, line: u32) {
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_get(chunks, current, line);
}

/// `arr.Last()` — returns `arr[arr.length - 1]`. Stack: [arr] → [last].
pub fn emit_linq_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
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
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, n_slot, line);
    collections::emit_slice(chunks, current, line);
}

/// `arr.ToList()` / `arr.ToArray()` — identity. Stack: [arr] → [arr].
pub fn emit_linq_identity(_chunks: &mut [Chunk], _current: usize, _line: u32) {
    // arr is already on the stack; no-op.
}

/// `arr.Average()` — `Sum(arr) / Length(arr)`. Stack: [arr] → [number].
pub fn emit_linq_average(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_sum(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::F64_DIV, line);
}

/// `arr.FirstOrDefault()` — `arr.length === 0 ? default : arr[0]`.
/// `default` is `0` (numeric) — generalising requires a per-T hook.
/// Stack: [arr] → [first | 0].
pub fn emit_linq_first_or_default(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // Test arr.length == 0
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);

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
    crate::emitter::ops::emit_dyn_ge(&mut chunks[current], line);
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
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    collections::emit_iter_values(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);

    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
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
    crate::emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_elem_slot, line);
    crate::emitter::ops::emit_dyn_eq(&mut chunks[current], line);
    let equal_values = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);

    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(2, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(equal_values);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
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

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);

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
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let if_block = chunks[current].emit_block(line);
    crate::emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // skip increment if false
    // count++
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunks[current].emit_i32_const(1, line);
    crate::emitter::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(if_block);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
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

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);

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
    // `__vybe_sort_by_key(arr, keyFn)` is the same routine `OrderBy`
    // already uses — call it then reverse the result in place.
    let global = chunks[current].add_constant(Value::String(Arc::from("__vybe_sort_by_key")));
    // Stack: [arr, keyFn] → need [globalfn, arr, keyFn].
    // Pop both into locals so we can stage the call cleanly.
    let arr_slot = alloc_locals(&mut chunks[current], 2);
    let fn_slot = arr_slot + 1;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    chunk.emit_op_u16(Op::GLOBAL_GET, global, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    // Reverse in place — emit_reverse is the array reverse opcode.
    collections::emit_reverse(chunks, current, line);
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
    collections::emit_iter_values(chunks, current, line);
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
    collections::emit_iter_values(chunks, current, line);
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
    crate::emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    crate::emitter::ops::emit_dyn_ge(&mut chunks[current], line);
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
