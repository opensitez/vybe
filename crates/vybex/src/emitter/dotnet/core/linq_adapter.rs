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

use std::sync::Arc;
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

use crate::emitter::collections;
use crate::emitter::loops;

/// Allocate `count` consecutive scratch locals; returns the first slot.
fn alloc_locals(chunk: &mut Chunk, count: u16) -> u16 {
    let base = chunk.local_count;
    chunk.local_count = base + count;
    base
}

// ── Pure reductions (no fn arg) ──────────────────────────────────────────

/// `arr.First()` — returns `arr[0]`. Stack: [arr] → [first].
pub fn emit_linq_first(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_op(Op::I32_CONST_0, line);
    collections::emit_get(chunks, current, line);
}

/// `arr.Last()` — returns `arr[arr.length - 1]`. Stack: [arr] → [last].
pub fn emit_linq_last(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    let one = chunks[current].add_constant(Value::I32(1));
    chunks[current].emit_op_u16(Op::CONST, one, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
}

/// `arr.Skip(n)` — `arr.slice(n, arr.length)`. Stack: [arr, n] → [array].
pub fn emit_linq_skip(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 2);
    let n_slot = arr_slot + 1;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, n_slot, line);   chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
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
    chunk.emit_op_u16(Op::LOCAL_SET, n_slot, line);   chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op(Op::I32_CONST_0, line);
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
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
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
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);

    // Test arr.length == 0
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op(Op::DYN_EQ, line);

    // Branch: if empty push default(0), else push arr[0].
    chunks[current].emit_op(Op::DYN_TO_BOOL, line);
    let empty = chunks[current].emit_jump(Op::BR_IF_TRUE, line);
    // non-empty path
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    collections::emit_get(chunks, current, line);
    let done = chunks[current].emit_jump(Op::BR, line);
    // empty path
    chunks[current].patch_jump(empty);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].patch_jump(done);
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
    chunks[current].emit_op(Op::DROP, line);

    // result = []
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    // for_in_start leaves arr[i] on the stack — stash to elem_slot.
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    // result.indexOf(elem) — push elem only if it is NOT already in result.
    // Structured WASM block matches the pattern used by `emit_filter`.
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_index_of(chunks, current, line);
    chunks[current].emit_op(Op::I32_CONST_0, line);
    chunks[current].emit_op(Op::DYN_GE, line);
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
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);  chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::I32_CONST_0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line); chunk.emit_op(Op::DROP, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op(Op::DROP, line);

    // pred(elem) → if-true increment count.  Use a structured WASM
    // block (same pattern as `emit_filter`) — byte-offset
    // `emit_jump(BR_IF_FALSE)` does not interleave correctly with the
    // outer `for_in` body block, so we open an inner block and `br_if`
    // out of it on the false branch.
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    chunks[current].emit_op(Op::DYN_TO_BOOL, line);
    let if_block = chunks[current].emit_block(line);
    chunks[current].emit_op(Op::DYN_NOT, line);
    chunks[current].emit_br_if(0, line); // skip increment if false
    // count++
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    let one = chunks[current].add_constant(Value::I32(1));
    chunks[current].emit_op_u16(Op::CONST, one, line);
    chunks[current].emit_op(Op::DYN_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_op(Op::DROP, line);
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
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);  chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc_slot, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);

    // for elem in arr: acc = fn(acc, elem)
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    {
        let chunk = &mut chunks[current];
        let elem_local = {
            let s = chunk.local_count;
            chunk.local_count = s + 1;
            s
        };
        chunk.emit_op_u16(Op::LOCAL_SET, elem_local, line);
        chunk.emit_op(Op::DROP, line);

        chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, acc_slot, line);
        chunk.emit_op_u16(Op::LOCAL_GET, elem_local, line);
        chunk.emit_op_u8(Op::CALL_REF, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, acc_slot, line);
        chunk.emit_op(Op::DROP, line);
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
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);  chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::GLOBAL_GET, global, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunk.emit_op_u8(Op::CALL_REF, 2, line);
    // Reverse in place — emit_reverse is the array reverse opcode.
    collections::emit_reverse(chunks, current, line);
}

