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

use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use vybe_compiler::primitives::collections;
use vybe_compiler::primitives::generators;
use vybe_compiler::primitives::loops;
use vybe_compiler::primitives::ops;
use vybe_compiler::primitives::class_slots;

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
    // that emits the spec `call`, or the index is out of range when the adapter
    // runs in a non-script chunk (e.g. compiled into a `__linq_*` vtable chunk).
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_linq_structural_key(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:json", "stringify", 1, line);
}

fn emit_value_eq_stamp_test(chunks: &mut [Chunk], current: usize, value_slot: u16, line: u32) {
    let result_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_import_call(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_string_const("__value_eq", line);
    collections::emit_get(chunks, current, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_linq_value_equals_slots(
    chunks: &mut [Chunk],
    current: usize,
    left_slot: u16,
    right_slot: u16,
    line: u32,
) {
    let left_is_value_slot = alloc_locals(&mut chunks[current], 2);
    let right_is_value_slot = left_is_value_slot + 1;

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_else(line);

    emit_value_eq_stamp_test(chunks, current, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_is_value_slot, line);
    emit_value_eq_stamp_test(chunks, current, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_is_value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_is_value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_is_value_slot, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    emit_linq_structural_key(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_linq_structural_key(chunks, current, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    emit_linq_structural_key(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_linq_structural_key(chunks, current, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

fn emit_linq_array_contains_value_slot(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    needle_slot: u16,
    line: u32,
) {
    let base = alloc_locals(&mut chunks[current], 3);
    let result_slot = base;
    let idx_slot = base + 1;
    let elem_slot = base + 2;

    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, array_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    let already_found = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);
    emit_linq_value_equals_slots(chunks, current, elem_slot, needle_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(already_found);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

fn emit_linq_comparer_equals_slots(
    chunks: &mut [Chunk],
    current: usize,
    comparer_slot: u16,
    left_slot: u16,
    right_slot: u16,
    line: u32,
) {
    let equals_fn_slot = alloc_locals(&mut chunks[current], 1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_string_const("__dotnet_stringcomparer_ordinalignorecase", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_linq_value_equals_slots(chunks, current, left_slot, right_slot, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_string_const("equals", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, equals_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, equals_fn_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_string_const("Equals", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, equals_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, equals_fn_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_string_const("__vb_iface_iequalitycomparer_equals", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, equals_fn_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, equals_fn_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    emit_linq_value_equals_slots(chunks, current, left_slot, right_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, equals_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 3, 1, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_linq_key_with_comparer(
    chunks: &mut [Chunk],
    current: usize,
    comparer_slot: u16,
    line: u32,
) {
    let value_slot = alloc_locals(&mut chunks[current], 2);
    let hash_fn_slot = value_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_string_const("__dotnet_stringcomparer_ordinalignorecase", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_linq_structural_key(chunks, current, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_string_const("hash", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hash_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hash_fn_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_string_const("gethashcode", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hash_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hash_fn_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_string_const("GetHashCode", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, hash_fn_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hash_fn_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_linq_structural_key(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, hash_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_invalid_operation_exception(
    chunks: &mut [Chunk],
    current: usize,
    message: &str,
    line: u32,
) {
    // ⛔ `emit_exception_new_finalize` consumes `[obj, obj, message]` — it
    // STAMPS an object that the caller has already created and duplicated.
    // Handing it a bare message throws something with no type and no message,
    // so `Catch ex As InvalidOperationException` cannot match it and `ex.Message`
    // is empty.
    vybe_compiler::primitives::errors::emit_exception_new(
        &mut chunks[current],
        "InvalidOperationException",
        class_slots::ValueSource::ConstStr(message.to_string()),
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
}

/// Resolve the `First`/`Last`/`Single` overload set and leave the sequence to
/// reduce in one slot and the fallback in another.
///
/// .NET declares four shapes — `()`, `(predicate)`, `(defaultValue)` and
/// `(predicate, defaultValue)` — and selects between the two one-argument
/// forms on the ARGUMENT TYPE. There is no static type here, so the choice is
/// made on whether the argument is CALLABLE, which answers the same question
/// for every case a caller can actually write.
///
/// ⛔ `argc` counts the RECEIVER, so a bare `First()` arrives as 1.
///
/// Stack on entry: `[seq, arg1?, arg2?]`; on exit the stack is empty and the
/// sequence lives in the returned `arr_slot`, the fallback in `default_slot`.
fn emit_resolve_sequence_overloads(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) -> (u16, u16) {
    let base = alloc_locals(&mut chunks[current], 8);
    let arr_slot = base;
    let default_slot = base + 1;
    let pred_slot = base + 2;
    let has_pred = base + 3;
    let arg_slot = base + 4;
    let idx_slot = base + 5;
    let elem_slot = base + 6;
    let out_slot = base + 7;

    {
        let chunk = &mut chunks[current];
        // `default(T)` is 0 for the numeric cases this surface reduces over,
        // which is the convention the existing `*OrDefault` bodies already use.
        chunk.emit_i32_const(0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, default_slot, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, has_pred, line);

        if argc >= 3 {
            chunk.emit_op_u16(Op::LOCAL_SET, default_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, pred_slot, line);
            chunk.emit_i32_const(1, line);
            chunk.emit_op_u16(Op::LOCAL_SET, has_pred, line);
        } else if argc == 2 {
            chunk.emit_op_u16(Op::LOCAL_SET, arg_slot, line);
            let typeof_fn = chunk.add_import("ecma:value", "typeof");
            chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
            chunk.emit_call(typeof_fn, 1, line);
            chunk.emit_string_const("function", line);
            vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
            chunk.emit_if(line);
            chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, pred_slot, line);
            chunk.emit_i32_const(1, line);
            chunk.emit_op_u16(Op::LOCAL_SET, has_pred, line);
            chunk.emit_else(line);
            chunk.emit_op_u16(Op::LOCAL_GET, arg_slot, line);
            chunk.emit_op_u16(Op::LOCAL_SET, default_slot, line);
            chunk.emit_end(line);
        }
        chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    }
    materialize_receiver_slot(chunks, current, arr_slot, line);

    // With a predicate the sequence has to be filtered before the reduction —
    // `First(p)` is the first element SATISFYING p, not a test of the first.
    chunks[current].emit_op_u16(Op::LOCAL_GET, has_pred, line);
    chunks[current].emit_if(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pred_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let if_block = chunks[current].emit_block(line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(if_block);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunks[current].emit_end(line);

    (arr_slot, default_slot)
}

/// `[] → [len]` for the sequence in `slot`.
fn emit_len_of(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    collections::emit_len(chunks, current, line);
}

/// `[] → [i32 0/1]` — whether the sequence in `slot` has exactly `n` elements.
fn emit_len_eq(chunks: &mut [Chunk], current: usize, slot: u16, n: i32, line: u32) {
    emit_len_of(chunks, current, slot, line);
    chunks[current].emit_i32_const(n, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
}

/// `[] → [element]` at `index` of the sequence in `slot`.
fn emit_element_of(chunks: &mut [Chunk], current: usize, slot: u16, index: i32, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_i32_const(index, line);
    collections::emit_get(chunks, current, line);
}

/// `[] → [last element]` of the sequence in `slot`.
fn emit_last_of(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    emit_len_of(chunks, current, slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
}

// ── Pure reductions (no fn arg) ──────────────────────────────────────────

/// `seq.First([predicate])` — the first element, or the first satisfying
/// `predicate`. Throws when there is none, which is what separates it from
/// `FirstOrDefault`.
pub fn emit_linq_first(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (arr_slot, _) = emit_resolve_sequence_overloads(chunks, current, argc, line);
    emit_len_eq(chunks, current, arr_slot, 0, line);
    chunks[current].emit_if(line);
    emit_invalid_operation_exception(chunks, current, "Sequence contains no elements.", line);
    chunks[current].emit_end(line);
    emit_element_of(chunks, current, arr_slot, 0, line);
}

/// `seq.Last([predicate])`.
pub fn emit_linq_last(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (arr_slot, _) = emit_resolve_sequence_overloads(chunks, current, argc, line);
    emit_len_eq(chunks, current, arr_slot, 0, line);
    chunks[current].emit_if(line);
    emit_invalid_operation_exception(chunks, current, "Sequence contains no elements.", line);
    chunks[current].emit_end(line);
    emit_last_of(chunks, current, arr_slot, line);
}

/// `seq.LastOrDefault([predicate][, defaultValue])`.
pub fn emit_linq_last_or_default(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (arr_slot, default_slot) = emit_resolve_sequence_overloads(chunks, current, argc, line);
    emit_len_eq(chunks, current, arr_slot, 0, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, default_slot, line);
    chunks[current].emit_else(line);
    emit_last_of(chunks, current, arr_slot, line);
    chunks[current].emit_end(line);
}

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

/// `arr.Sum(fn)` — map through selector then sum. Stack: [seq, fn] → [number].
pub fn emit_linq_sum_selector(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = arr_slot + 1;
    let mapped_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
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

/// `seq.FirstOrDefault([predicate][, defaultValue])`.
pub fn emit_linq_first_or_default(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (arr_slot, default_slot) = emit_resolve_sequence_overloads(chunks, current, argc, line);
    emit_len_eq(chunks, current, arr_slot, 0, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, default_slot, line);
    chunks[current].emit_else(line);
    emit_element_of(chunks, current, arr_slot, 0, line);
    chunks[current].emit_end(line);
}

pub fn emit_linq_distinct(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let result_slot = arr_slot + 1;
    let idx_slot = arr_slot + 2;
    let elem_slot = arr_slot + 3;
    let duplicate_slot = arr_slot + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    // result = []
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    // for_in_start leaves arr[i] on the stack — stash to elem_slot.
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    emit_linq_array_contains_value_slot(chunks, current, result_slot, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, duplicate_slot, line);

    let if_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, duplicate_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
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

/// `arr.Distinct(comparer)` — comparer-aware distinct. Currently recognises
/// the shared ordinal-ignore-case string comparer sentinel used by .NET-shape
/// walkers; other comparer values fall back to ordinary equality.
/// Stack: [arr, comparer] → [array].
pub fn emit_linq_distinct_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 8);
    let comparer_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let scan_idx_slot = arr_slot + 5;
    let existing_slot = arr_slot + 6;
    let duplicate_slot = arr_slot + 7;

    chunks[current].emit_op_u16(Op::LOCAL_SET, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, duplicate_slot, line);

    let scan_state = loops::emit_for_in_start(chunks, current, result_slot, scan_idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, existing_slot, line);
    let skip_after_duplicate = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, duplicate_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);
    emit_linq_comparer_equals_slots(
        chunks,
        current,
        comparer_slot,
        existing_slot,
        elem_slot,
        line,
    );
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, duplicate_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(skip_after_duplicate);
    loops::emit_for_in_end(chunks, current, scan_idx_slot, scan_state, line);

    let push_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, duplicate_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(push_block);

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
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    // if keys.indexOf(key) < 0 { keys.push(key); result.push(elem); }
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_index_of(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
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

/// `arr.DistinctBy(keyFn, comparer)` — first element for each comparer-distinct
/// projected key.
/// Stack: [arr, keyFn, comparer] → [array].
pub fn emit_linq_distinct_by_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 8);
    let fn_slot = arr_slot + 1;
    let comparer_slot = arr_slot + 2;
    let result_slot = arr_slot + 3;
    let keys_slot = arr_slot + 4;
    let idx_slot = arr_slot + 5;
    let elem_slot = arr_slot + 6;
    let key_slot = arr_slot + 7;

    chunks[current].emit_op_u16(Op::LOCAL_SET, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    emit_linq_key_with_comparer(chunks, current, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_index_of(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    let if_block = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);
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
    let left_slot = alloc_locals(&mut chunks[current], 7);
    let right_slot = left_slot + 1;
    let len_slot = left_slot + 2;
    let idx_slot = left_slot + 3;
    let right_elem_slot = left_slot + 4;
    let result_slot = left_slot + 5;
    let left_elem_slot = left_slot + 6;

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
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_elem_slot, line);
    emit_linq_value_equals_slots(chunks, current, left_elem_slot, right_elem_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    let equal_values = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);

    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(2, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(equal_values);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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

/// `arr.SequenceEqual(other, comparer)`.
/// Stack: [arr, other, comparer] → [bool].
pub fn emit_linq_sequence_equal_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    let left_slot = alloc_locals(&mut chunks[current], 8);
    let right_slot = left_slot + 1;
    let comparer_slot = left_slot + 2;
    let len_slot = left_slot + 3;
    let idx_slot = left_slot + 4;
    let right_key_slot = left_slot + 5;
    let result_slot = left_slot + 6;
    let left_key_slot = left_slot + 7;

    chunks[current].emit_op_u16(Op::LOCAL_SET, comparer_slot, line);
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
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    emit_linq_key_with_comparer(chunks, current, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_key_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    emit_linq_key_with_comparer(chunks, current, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_key_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    let equal_values = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);

    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_br(2, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(equal_values);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
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
/// LINQ overload — `primitives/calls.rs::compile_call`'s
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
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let if_block = chunks[current].emit_block(line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // skip increment if false
    // count++
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(if_block);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
}

/// `arr.All(pred)` — true iff every element satisfies `pred`.
/// Stack: [arr, pred] → [bool].
pub fn emit_linq_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    let skip_after_false = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(skip_after_false);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
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
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let if_block = chunks[current].emit_block(line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// `arr.Any(pred)` — true iff any element satisfies `pred`.
/// Stack: [arr, pred] → [bool].
pub fn emit_linq_any_pred(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    let skip_after_true = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(skip_after_true);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.Contains(x)` — .NET equality over primitive/structural keys.
/// Stack: [arr, x] → [bool].
pub fn emit_linq_contains(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let x_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    chunks[current].emit_op_u16(Op::LOCAL_SET, x_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    let skip_after_true = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);
    emit_linq_value_equals_slots(chunks, current, elem_slot, x_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(skip_after_true);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.Contains(x, comparer)`.
/// Stack: [arr, x, comparer] → [bool].
pub fn emit_linq_contains_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 6);
    let x_slot = arr_slot + 1;
    let comparer_slot = arr_slot + 2;
    let keys_slot = arr_slot + 3;
    let idx_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;
    chunks[current].emit_op_u16(Op::LOCAL_SET, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, x_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    emit_linq_key_with_comparer(chunks, current, comparer_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, x_slot, line);
    emit_linq_key_with_comparer(chunks, current, comparer_slot, line);
    collections::emit_index_of(chunks, current, line);
    chunks[current].emit_i32_const(-1, line);
    chunks[current].emit_op(Op::I32_NE, line);
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // not skipping → leave flag
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // pred still true → keep skipping
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, skipping_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(stop_block);

    // if !skipping: result.push(elem)
    let push_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, skipping_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
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

/// `arr.SkipWhile(Function(elem, index) ...)`.
/// Stack: [arr, pred] → [array].
pub fn emit_linq_skip_while_indexed(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_skip_take_while_indexed(chunks, current, line, false);
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
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // already stopped → leave flag
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line); // pred true → keep taking
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, taking_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(stop_block);

    // if taking: result.push(elem)
    let push_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, taking_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
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

/// `arr.TakeWhile(Function(elem, index) ...)`.
/// Stack: [arr, pred] → [array].
pub fn emit_linq_take_while_indexed(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_skip_take_while_indexed(chunks, current, line, true);
}

fn emit_linq_skip_take_while_indexed(chunks: &mut [Chunk], current: usize, line: u32, take: bool) {
    let arr_slot = alloc_locals(&mut chunks[current], 6);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let active_slot = arr_slot + 5;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, active_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    let stop_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, active_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, active_slot, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(stop_block);

    let push_block = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, active_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    if take {
        vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    }
    chunks[current].emit_br_if(0, line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
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
        chunk.emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
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
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Invoke a selector that may or may not declare an index parameter.
///
/// `SelectMany`, `Select` and `Where` each carry a second .NET overload whose
/// delegate takes `(element, index)`, and the overload is chosen on the
/// DELEGATE'S SHAPE. `ecma:function.length` answers that shape at run time —
/// the same question C# answers at compile time — because a lambda's chunk
/// carries `params.len()` as its arity and `Function::arity` is copied from it.
///
/// ⛔ A non-function receiver answers `0`, which takes the one-argument path.
/// That is the pre-existing behaviour, so a value that is not introspectable
/// degrades rather than trapping.
///
/// Stack: [] → [result].
fn emit_call_selector_indexed(
    chunks: &mut [Chunk],
    current: usize,
    fn_slot: u16,
    elem_slot: u16,
    idx_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    emit_import_call(chunks, current, "ecma:function", "length", 1, line);
    core_wasm::i32_const(&mut chunks[current], line, 2);
    ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_if_value(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);

    chunks[current].emit_end(line);
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

    // mapped = fn(elem) — or fn(elem, idx) for the indexed overload.
    emit_call_selector_indexed(chunks, current, fn_slot, elem_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_slot, line);
    // The inner sequence can be any iterable — `Enumerable.Range`, a `List<T>`,
    // a generator — and `concat` needs an indexable array.
    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_slot, line);

    // result = result.concat(mapped)
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, mapped_slot, line);
    collections::emit_concat(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.SelectMany(collectionFn, resultFn)` — the projecting overload.
///
/// ⛔ `resultFn` is `Func<TSource, TCollection, TResult>`: it is called once per
/// **(outer, inner-element) PAIR**, not once per flattened element. Mapping it
/// over the flattened array instead gives the right answer for any selector
/// that ignores its first parameter and a silently wrong one for every other.
///
/// Stack: [arr, collectionFn, resultFn] → [array].
pub fn emit_linq_select_many_result(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 8);
    let collection_fn_slot = arr_slot + 1;
    let result_fn_slot = arr_slot + 2;
    let result_slot = arr_slot + 3;
    let idx_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;
    let mapped_slot = arr_slot + 6;
    let inner_idx_slot = arr_slot + 7;

    chunks[current].emit_op_u16(Op::LOCAL_SET, result_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, collection_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let outer = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // mapped = collectionFn(elem) — or collectionFn(elem, idx).
    emit_call_selector_indexed(chunks, current, collection_fn_slot, elem_slot, idx_slot, line);
    collections::emit_spread_iterable(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, mapped_slot, line);

    // The outer element travels into the inner loop as `resultFn`'s first
    // argument, which is what makes this overload different from `Select`
    // over a flattened sequence.
    let inner_elem_slot = alloc_locals(&mut chunks[current], 1);
    let inner = loops::emit_for_in_start(chunks, current, mapped_slot, inner_idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, inner_idx_slot, inner, line);
    loops::emit_for_in_end(chunks, current, idx_slot, outer, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `outer.Join(inner, outerKeyFn, innerKeyFn, resultFn)` — the inner equijoin.
///
/// One row per MATCHING PAIR, in outer order and then inner order, which is the
/// order .NET's own implementation produces. An outer element with no match
/// contributes nothing — that is what separates `Join` from `GroupJoin`.
///
/// Stack: [outer, inner, outerKeyFn, innerKeyFn, resultFn] → [array].
pub fn emit_linq_join(chunks: &mut [Chunk], current: usize, line: u32) {
    let outer_slot = alloc_locals(&mut chunks[current], 12);
    let inner_slot = outer_slot + 1;
    let outer_key_fn_slot = outer_slot + 2;
    let inner_key_fn_slot = outer_slot + 3;
    let result_fn_slot = outer_slot + 4;
    let out_slot = outer_slot + 5;
    let outer_idx_slot = outer_slot + 6;
    let inner_idx_slot = outer_slot + 7;
    let outer_elem_slot = outer_slot + 8;
    let inner_elem_slot = outer_slot + 9;
    let outer_key_slot = outer_slot + 10;
    let inner_key_slot = outer_slot + 11;

    chunks[current].emit_op_u16(Op::LOCAL_SET, result_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_slot, line);
    materialize_receiver_slot(chunks, current, outer_slot, line);
    materialize_receiver_slot(chunks, current, inner_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    let outer_loop = loops::emit_for_in_start(chunks, current, outer_slot, outer_idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, outer_key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, outer_elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_key_slot, line);

    let inner_loop = loops::emit_for_in_start(chunks, current, inner_slot, inner_idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_key_slot, line);

    // Keys join on .NET VALUE equality, so a struct key (a tuple, a
    // `KeyValuePair`, a `DateTime`) matches by content, not by identity.
    emit_linq_value_equals_slots(chunks, current, outer_key_slot, inner_key_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    let no_match = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, outer_elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(no_match);

    loops::emit_for_in_end(chunks, current, inner_idx_slot, inner_loop, line);
    loops::emit_for_in_end(chunks, current, outer_idx_slot, outer_loop, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// `outer.GroupJoin(inner, outerKeyFn, innerKeyFn, resultFn)`.
///
/// One row per OUTER element — including outer elements with no match, which
/// receive an empty group. `resultFn` is `(outerElement, matchingInners)`.
///
/// Stack: [outer, inner, outerKeyFn, innerKeyFn, resultFn] → [array].
pub fn emit_linq_group_join(chunks: &mut [Chunk], current: usize, line: u32) {
    let outer_slot = alloc_locals(&mut chunks[current], 13);
    let inner_slot = outer_slot + 1;
    let outer_key_fn_slot = outer_slot + 2;
    let inner_key_fn_slot = outer_slot + 3;
    let result_fn_slot = outer_slot + 4;
    let out_slot = outer_slot + 5;
    let outer_idx_slot = outer_slot + 6;
    let inner_idx_slot = outer_slot + 7;
    let outer_elem_slot = outer_slot + 8;
    let inner_elem_slot = outer_slot + 9;
    let outer_key_slot = outer_slot + 10;
    let inner_key_slot = outer_slot + 11;
    let group_slot = outer_slot + 12;

    chunks[current].emit_op_u16(Op::LOCAL_SET, result_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_slot, line);
    materialize_receiver_slot(chunks, current, outer_slot, line);
    materialize_receiver_slot(chunks, current, inner_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);

    let outer_loop = loops::emit_for_in_start(chunks, current, outer_slot, outer_idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, outer_key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, outer_elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, outer_key_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, group_slot, line);

    let inner_loop = loops::emit_for_in_start(chunks, current, inner_slot, inner_idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, inner_key_slot, line);

    emit_linq_value_equals_slots(chunks, current, outer_key_slot, inner_key_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    let no_match = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, inner_elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);
    chunks[current].patch_block(no_match);

    loops::emit_for_in_end(chunks, current, inner_idx_slot, inner_loop, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, outer_elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, outer_idx_slot, outer_loop, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// `Enumerable.Range(start, count)` — the counted integer sequence.
/// Stack: [start, count] → [array].
pub fn emit_linq_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let start_slot = alloc_locals(&mut chunks[current], 4);
    let count_slot = start_slot + 1;
    let out_slot = start_slot + 2;
    let idx_slot = start_slot + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    let state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// `Enumerable.Repeat(element, count)`.
/// Stack: [element, count] → [array].
pub fn emit_linq_repeat(chunks: &mut [Chunk], current: usize, line: u32) {
    let elem_slot = alloc_locals(&mut chunks[current], 4);
    let count_slot = elem_slot + 1;
    let out_slot = elem_slot + 2;
    let idx_slot = elem_slot + 3;

    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    let state = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    loops::emit_loop_end(chunks, current, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

/// `arr.GroupBy(keyFn)` — returns an array of grouping objects.
/// Each group has `Key`, `Items`, and `Count` properties.
/// Stack: [arr, keyFn] → [groups].
pub fn emit_linq_group_by(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_group_by_core(chunks, current, false, line);
}

/// `arr.GroupBy(keyFn, elementFn)` — grouping whose members are PROJECTED.
/// `Key` still comes from `keyFn`; every group member is `elementFn(elem)`.
/// Stack: [arr, keyFn, elementFn] → [groups].
pub fn emit_linq_group_by_element(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_group_by_core(chunks, current, true, line);
}

fn emit_group_by_core(
    chunks: &mut [Chunk],
    current: usize,
    has_element_selector: bool,
    line: u32,
) {
    let arr_slot = alloc_locals(&mut chunks[current], 13);
    let fn_slot = arr_slot + 1;
    let map_slot = arr_slot + 2;
    let out_slot = arr_slot + 3;
    let idx_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;
    let key_slot = arr_slot + 6;
    let group_slot = arr_slot + 7;
    let items_slot = arr_slot + 8;
    let map_key_slot = arr_slot + 9;
    let sum_slot = arr_slot + 10;
    let element_fn_slot = arr_slot + 11;
    // What lands in the group: the source element, or its projection.
    let value_slot = arr_slot + 12;

    // Stack: [arr, keyFn] — or [arr, keyFn, elementFn].
    if has_element_selector {
        chunks[current].emit_op_u16(Op::LOCAL_SET, element_fn_slot, line);
    }
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
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    // value = elementFn(elem), or the element itself.
    if has_element_selector {
        chunks[current].emit_op_u16(Op::LOCAL_GET, element_fn_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
        chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    emit_linq_structural_key(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_key_slot, line);

    // if !groupMap.has(key) { create group object, initialize fields, save map, out.push(group) }
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_key_slot, line);
    emit_import_call(chunks, current, "ecma:map", "has", 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let maybe_new = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line); // already exists

    // ⛔ The group IS the sequence. `IGrouping<K,T> : IEnumerable<T>`, so
    // `foreach (var v in g)` and `String.Join(",", g)` iterate the MEMBERS —
    // an object with an `Items` field satisfies neither. The named fields ride
    // along as string-keyed properties, which an Array exotic object carries
    // per ECMA-262 §10.4.2.2 (`ecma:array.set` falls through to the property
    // bag for a non-index key).
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, group_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_slot, line);

    // group["Key"] = key
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Key", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("key", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // `Items` names the group itself, so the VB query lowering's
    // `__vb_group.Items` and a direct iteration of the group see one array.
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Items", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("items", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // group["Count"] = 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Count", line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("count", line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // group["Sum"] = 0
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Sum", line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("sum", line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // groupMap[key] = group
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("count", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_slot, line);
    collections::emit_len(chunks, current, line);
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
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_key_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, group_slot, line);

    // The group IS the member array — no `Items` indirection to follow.
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items_slot, line);

    // items.push(value)
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // group["Count"] = items.length
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Count", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items_slot, line);
    collections::emit_len(chunks, current, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    // group["Sum"] += elem for numeric elements. Non-numeric element groups
    // still support selector-based Sum over Items; the cached Sum field stays 0.
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Sum", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sum_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_import_call(chunks, current, "wasm:js-number", "test", 1, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sum_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("Sum", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, group_slot, line);
    chunks[current].emit_string_const("sum", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, sum_slot, line);
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
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, val_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, val_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
}

/// `arr.ToDictionary(keyFn)` — map each projected key to the original element.
/// Stack: [arr, keyFn] → [map].
pub fn emit_linq_to_dictionary_key(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 6);
    let key_fn_slot = arr_slot + 1;
    let map_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let key_slot = arr_slot + 5;

    chunks[current].emit_op_u16(Op::LOCAL_SET, key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
}

/// `arr.ToLookup(keyFn)` — map each projected key to an array of matching
/// elements. Stack: [arr, keyFn] → [map].
pub fn emit_linq_to_lookup(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 8);
    let key_fn_slot = arr_slot + 1;
    let map_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let key_slot = arr_slot + 5;
    let bucket_slot = arr_slot + 6;
    let has_slot = arr_slot + 7;

    chunks[current].emit_op_u16(Op::LOCAL_SET, key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, map_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    emit_import_call(chunks, current, "ecma:map", "has", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, has_slot, line);

    let existing = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, has_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bucket_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(existing);

    chunks[current].emit_op_u16(Op::LOCAL_GET, map_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bucket_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, bucket_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
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
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    let too_short = chunks[current].emit_block(line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
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

/// `arr.Concat(other)` — materialize both sequences and concatenate.
/// Stack: [arr, other] → [array].
pub fn emit_linq_concat(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 2);
    let other_slot = arr_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);
    materialize_receiver_slot(chunks, current, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    collections::emit_concat(chunks, current, line);
}

/// `arr.Union(other)` — concatenation with duplicate removal.
/// Stack: [arr, other] → [array].
pub fn emit_linq_union(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_concat(chunks, current, line);
    emit_linq_distinct(chunks, current, line);
}

/// `arr.Union(other, comparer)`.
/// Stack: [arr, other, comparer] → [array].
pub fn emit_linq_union_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = alloc_locals(&mut chunks[current], 3);
    let comparer_slot = base;
    let other_slot = base + 1;
    let left_slot = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    emit_linq_concat(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, comparer_slot, line);
    emit_linq_distinct_comparer(chunks, current, line);
}

fn emit_linq_set_filter(chunks: &mut [Chunk], current: usize, line: u32, keep_matches: bool) {
    let left_slot = alloc_locals(&mut chunks[current], 7);
    let right_slot = left_slot + 1;
    let result_slot = left_slot + 2;
    let idx_slot = left_slot + 3;
    let elem_slot = left_slot + 4;
    let contains_slot = left_slot + 5;
    let duplicate_slot = left_slot + 6;

    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);
    materialize_receiver_slot(chunks, current, left_slot, line);
    materialize_receiver_slot(chunks, current, right_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, left_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    emit_linq_array_contains_value_slot(chunks, current, right_slot, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, contains_slot, line);

    let skip = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, contains_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    if !keep_matches {
        vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    }
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);

    emit_linq_array_contains_value_slot(chunks, current, result_slot, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, duplicate_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, duplicate_slot, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(skip);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.Intersect(other)` — distinct values present in both sequences.
/// Stack: [arr, other] → [array].
pub fn emit_linq_intersect(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_set_filter(chunks, current, line, true);
}

/// `arr.Except(other)` — distinct values present in the receiver but not other.
/// Stack: [arr, other] → [array].
pub fn emit_linq_except(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_set_filter(chunks, current, line, false);
}

fn emit_linq_set_filter_comparer(
    chunks: &mut [Chunk],
    current: usize,
    line: u32,
    keep_matches: bool,
) {
    let left_slot = alloc_locals(&mut chunks[current], 10);
    let right_slot = left_slot + 1;
    let comparer_slot = left_slot + 2;
    let result_slot = left_slot + 3;
    let result_keys_slot = left_slot + 4;
    let right_keys_slot = left_slot + 5;
    let idx_slot = left_slot + 6;
    let elem_slot = left_slot + 7;
    let elem_key_slot = left_slot + 8;
    let contains_slot = left_slot + 9;

    chunks[current].emit_op_u16(Op::LOCAL_SET, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);
    materialize_receiver_slot(chunks, current, left_slot, line);
    materialize_receiver_slot(chunks, current, right_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_keys_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_keys_slot, line);

    let right_state = loops::emit_for_in_start(chunks, current, right_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    emit_linq_key_with_comparer(chunks, current, comparer_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, right_state, line);

    let state = loops::emit_for_in_start(chunks, current, left_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    emit_linq_key_with_comparer(chunks, current, comparer_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_key_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_key_slot, line);
    collections::emit_contains(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, contains_slot, line);

    let skip = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, contains_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    if !keep_matches {
        vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    }
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_key_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_key_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(skip);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_linq_intersect_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_set_filter_comparer(chunks, current, line, true);
}

pub fn emit_linq_except_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_set_filter_comparer(chunks, current, line, false);
}

fn emit_linq_set_filter_by(chunks: &mut [Chunk], current: usize, line: u32, keep_matches: bool) {
    let left_slot = alloc_locals(&mut chunks[current], 10);
    let keys_slot = left_slot + 1;
    let key_fn_slot = left_slot + 2;
    let result_slot = left_slot + 3;
    let result_keys_slot = left_slot + 4;
    let right_keys_slot = left_slot + 5;
    let idx_slot = left_slot + 6;
    let elem_slot = left_slot + 7;
    let elem_key_slot = left_slot + 8;
    let contains_slot = left_slot + 9;

    chunks[current].emit_op_u16(Op::LOCAL_SET, key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);
    materialize_receiver_slot(chunks, current, left_slot, line);
    materialize_receiver_slot(chunks, current, keys_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_keys_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_keys_slot, line);

    let key_state = loops::emit_for_in_start(chunks, current, keys_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    emit_linq_structural_key(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, key_state, line);

    let state = loops::emit_for_in_start(chunks, current, left_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    emit_linq_structural_key(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_key_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, right_keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_key_slot, line);
    collections::emit_contains(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, contains_slot, line);

    let skip = chunks[current].emit_block(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, contains_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    if !keep_matches {
        vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    }
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_key_slot, line);
    collections::emit_contains(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_keys_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_key_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(skip);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

pub fn emit_linq_union_by(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = alloc_locals(&mut chunks[current], 3);
    let key_fn_slot = base;
    let other_slot = base + 1;
    let left_slot = base + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, other_slot, line);
    emit_linq_concat(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_fn_slot, line);
    emit_linq_distinct_by(chunks, current, line);
}

pub fn emit_linq_intersect_by(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_set_filter_by(chunks, current, line, true);
}

pub fn emit_linq_except_by(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_linq_set_filter_by(chunks, current, line, false);
}

fn emit_type_name_is_any(
    chunks: &mut [Chunk],
    current: usize,
    type_slot: u16,
    names: &[&str],
    line: u32,
) {
    let out_slot = alloc_locals(&mut chunks[current], 1);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    for name in names {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, type_slot, line);
        chunks[current].emit_string_const(name, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_op(Op::I32_OR, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
}

fn emit_linq_of_type_match(
    chunks: &mut [Chunk],
    current: usize,
    type_slot: u16,
    elem_slot: u16,
    line: u32,
) {
    let result_slot = alloc_locals(&mut chunks[current], 1);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    for (aliases, ecma_type) in [
        (
            &[
                "Integer", "Int32", "Long", "Int64", "Short", "Int16", "Byte", "Double", "Single",
                "Decimal", "int", "long", "short", "byte", "uint", "ulong", "ushort", "sbyte",
                "double", "float", "decimal",
            ][..],
            "number",
        ),
        (&["String", "string"][..], "string"),
        (&["Boolean", "Bool", "bool"][..], "boolean"),
    ] {
        emit_type_name_is_any(chunks, current, type_slot, aliases, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
        emit_import_call(chunks, current, "ecma:value", "typeof", 1, line);
        chunks[current].emit_string_const(ecma_type, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
        chunks[current].emit_end(line);
    }

    emit_type_name_is_any(chunks, current, type_slot, &["Object"], line);
    chunks[current].emit_if(line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `arr.OfType<T>()` — filter by normalized .NET primitive type name.
/// Stack: [arr, typeName] → [array].
pub fn emit_linq_of_type(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let type_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, type_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    emit_linq_of_type_match(chunks, current, type_slot, elem_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let skip = chunks[current].emit_block(line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].patch_block(skip);

    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
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
    chunks[current].emit_string_const(
        "Index was out of range. Must be non-negative and less than the size of the collection.",
        line,
    );
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        &mut chunks[current],
        "ArgumentOutOfRangeException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(&mut chunks[current], line);
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
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
}

/// `seq.Single([predicate])` — the sole element, throwing
/// `InvalidOperationException` unless there is exactly one.
pub fn emit_linq_single(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (arr_slot, _) = emit_resolve_sequence_overloads(chunks, current, argc, line);
    emit_len_eq(chunks, current, arr_slot, 0, line);
    chunks[current].emit_if(line);
    emit_invalid_operation_exception(chunks, current, "Sequence contains no elements.", line);
    chunks[current].emit_end(line);
    emit_len_of(chunks, current, arr_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_invalid_operation_exception(
        chunks,
        current,
        "Sequence contains more than one element.",
        line,
    );
    chunks[current].emit_end(line);
    emit_element_of(chunks, current, arr_slot, 0, line);
}

/// `seq.SingleOrDefault([predicate][, defaultValue])`.
///
/// ⛔ An EMPTY sequence answers the default, but MORE THAN ONE still throws —
/// `SingleOrDefault` only forgives the empty case.
pub fn emit_linq_single_or_default(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (arr_slot, default_slot) = emit_resolve_sequence_overloads(chunks, current, argc, line);
    emit_len_of(chunks, current, arr_slot, line);
    chunks[current].emit_i32_const(1, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_invalid_operation_exception(
        chunks,
        current,
        "Sequence contains more than one element.",
        line,
    );
    chunks[current].emit_end(line);
    emit_len_eq(chunks, current, arr_slot, 1, line);
    chunks[current].emit_if_value(line);
    emit_element_of(chunks, current, arr_slot, 0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, default_slot, line);
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
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, bestkey_slot, line);

    let state = loops::emit_for_in_start(chunks, current, arr_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // key = fn(elem)
    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunks[current].emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);

    // if key >/< bestKey { best = elem; bestKey = key }
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, bestkey_slot, line);
    if want_max {
        vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    } else {
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
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
        chunks[current].emit_op_u8_u8(Op::CALL_REF, 2, 1, line);
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
fn emit_len_minus_n_clamped(
    chunks: &mut [Chunk],
    current: usize,
    arr_slot: u16,
    n_slot: u16,
    line: u32,
) {
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
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_end(line);
}

/// `arr.DefaultIfEmpty(value)` — `arr` when non-empty, else `[value]`.
/// Stack: [arr, value] → [array].
pub fn emit_linq_default_if_empty_value(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 2);
    let value_slot = arr_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    materialize_receiver_slot(chunks, current, arr_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    collections::emit_len(chunks, current, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_end(line);
}

/// `seq.OrderDescending()` — ascending sort, reversed.
///
/// Composed from the two existing leaves rather than given its own comparison:
/// a descending order that disagreed with `collections.sorted` about ties or
/// mixed types would be a second answer to the same question.
pub fn emit_linq_order_descending(chunks: &mut [Chunk], current: usize, line: u32) {
    collections::emit_sorted(chunks, current, line);
    collections::emit_reverse(chunks, current, line);
}
