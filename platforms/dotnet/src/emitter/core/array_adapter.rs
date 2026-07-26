//! .NET `System.Array` static-method adapter — bytecode-only.
//!
//! `Array.Clear(arr, idx, count)` / `Copy(src, dst, count)` /
//! `Resize(arr, newSize)` / `Sort(arr)` are .NET-shape range
//! operations on plain Arrays. None of them have a 1:1 ECMA-262
//! §23.1 mirror (the closest analogues — `toSpliced`, `toSorted` —
//! return new arrays; .NET `Array.*` mutate in place). Each adapter
//! lowers to a stdlib bytecode chunk (`__vybe_*` global) that
//! composes the right `ecma:array.*` primitives, or to an inline
//! loop when no chunk fits.
//!
//! Pure WASM, zero `vybe:types` involvement.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;
use vybe_emitter::instructions::{core_wasm, host};

fn emit_throw_dotnet_exception(chunk: &mut Chunk, exception_name: &str, message: &str, line: u32) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_dup(line);
    chunk.emit_string_const(message, line);
    vybe_emitter::errors::emit_exception_new_finalize(chunk, exception_name, line);
    vybe_emitter::errors::emit_throw(chunk, line);
}

fn emit_index_bounds_check(
    chunk: &mut Chunk,
    arr_slot: u16,
    index_slot: u16,
    exception_name: &str,
    message: &str,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_ge(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(chunk, exception_name, message, line);
    chunk.emit_end(line);
}

/// `Array.Clear(arr, idx, count)` — reset `count` elements starting at
/// `idx` to a .NET-style default. Until the runtime carries per-array
/// element metadata, infer the default from the current element's
/// runtime category: numbers clear to `0`, booleans to `false`, and
/// reference-like values to `null`.
///
/// Stack on entry: `[arr, idx, count]` ; Stack on exit: `[null]`
pub fn emit_array_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    // Allocate scratch slots: arr, idx, count, i (counter), target index,
    // current element.
    let arr_slot = chunk.alloc_scratch(6);
    let idx_slot = arr_slot + 1;
    let count_slot = arr_slot + 2;
    let i_slot = arr_slot + 3;
    let target_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;

    // Stash args (top of stack first → reverse order)
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // i = 0
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    // Loop: while i < count { arr[idx + i] = default(existing_value); i++ }
    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    vybe_emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    host::emit(chunk, "wasm:js-boolean", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    core_wasm::bool_const(chunk, line, false);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    host::emit(chunk, "wasm:js-number", "test", 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_else(line);
    chunk.emit_op(Op::NULL, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line); // ARRAY_SET pushes the value; drop it

    // i++
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line); // end loop
    chunk.patch_loop(loop_p);
    chunk.emit_end(line); // end block
    chunk.patch_block(block_p);

    chunk.emit_op(Op::NULL, line);
}

/// `Array.Copy(src, dst, count)` — copy first `count` elements from
/// `src` to `dst`. Lowers to `__vybe_array_copy` runtime helper
/// (already bundled).
///
/// Stack on entry: `[src, dst, count]` ; Stack on exit: `[null]`
pub fn emit_array_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let count_slot = chunk.alloc_scratch(3);
    let dst_slot = count_slot + 1;
    let src_slot = count_slot + 2;
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, dst_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, src_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_ge(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, src_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    vybe_emitter::ops::emit_dyn_le(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, dst_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    vybe_emitter::ops::emit_dyn_le(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "ArgumentException",
        "Destination array was not long enough.",
        line,
    );
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, src_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, dst_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    vybe_emitter::collections::emit_runtime_helper_call(
        chunks,
        current,
        "__vybe_array_copy",
        3,
        line,
    );
    // Stdlib chunk returns null; leave it on stack.
}

/// Bounds-checked `Array.GetValue(index)` / VB array read.
/// Stack: `[arr, index]` -> `[value]`.
pub fn emit_array_get_checked(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let index_slot = chunk.alloc_scratch(2);
    let arr_slot = index_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    emit_index_bounds_check(
        chunk,
        arr_slot,
        index_slot,
        "IndexOutOfRangeException",
        "Index was outside the bounds of the array.",
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// Bounds-checked `List<T>.Item(index)`.
/// Stack: `[list, index]` -> `[value]`.
pub fn emit_list_get_checked(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let index_slot = chunk.alloc_scratch(2);
    let arr_slot = index_slot + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    emit_index_bounds_check(
        chunk,
        arr_slot,
        index_slot,
        "ArgumentOutOfRangeException",
        "Index was out of range. Must be non-negative and less than the size of the collection.",
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
}

/// Bounds-checked .NET range slice.
/// Stack: `[array, start, length]` -> `[array_slice]`.
pub fn emit_get_range_checked(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let len_slot = chunk.alloc_scratch(4);
    let start_slot = len_slot + 1;
    let arr_slot = len_slot + 2;
    let end_slot = len_slot + 3;
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, end_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunk.emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_ge(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunk.emit_i32_const(0, line);
    vybe_emitter::ops::emit_dyn_ge(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    vybe_emitter::ops::emit_dyn_le(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    emit_throw_dotnet_exception(
        chunk,
        "ArgumentOutOfRangeException",
        "Specified argument was out of the range of valid values.",
        line,
    );
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, start_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_emitter::collections::emit_get_range(chunks, current, line);
}

/// Bounds-checked `Array.SetValue(value, index)`.
/// Stack: `[arr, value, index]` -> `[null]`.
pub fn emit_array_set_checked(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let index_slot = chunk.alloc_scratch(3);
    let value_slot = index_slot + 1;
    let arr_slot = index_slot + 2;
    chunk.emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    emit_index_bounds_check(
        chunk,
        arr_slot,
        index_slot,
        "IndexOutOfRangeException",
        "Index was outside the bounds of the array.",
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op(Op::NULL, line);
}

/// `Array.Resize(arr, newSize)` — extend or truncate `arr` to
/// `newSize` elements. Lowers to `__vybe_redim` runtime helper.
///
/// Stack on entry: `[arr, newSize]` ; Stack on exit: `[arr]` (the
/// runtime helper returns the resized array; .NET `Array.Resize`
/// signature is by-ref but the bytecode propagates the value).
pub fn emit_array_resize(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_emitter::collections::emit_runtime_helper_call(chunks, current, "__vybe_redim", 2, line);
}

/// `Array.Sort(arr)` — in-place sort. Lowers to `__vybe_sort_in_place`
/// runtime helper.
///
/// Stack on entry: `[arr]` ; Stack on exit: `[null]` (sort is void in
/// .NET; the runtime helper returns the array but we drop it).
pub fn emit_array_sort(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_emitter::collections::emit_runtime_helper_call(
        chunks,
        current,
        "__vybe_sort_in_place",
        1,
        line,
    );
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `Array.Reverse(arr)` — in-place reverse. Stack: `[arr]` → `[null]`.
/// Reverse mutates the array; .NET's signature returns void.
pub fn emit_array_reverse(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_emitter::collections::emit_reverse(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `List<T>.RemoveAll(pred)` — remove each matching element in place and
/// return the number removed. Stack: `[list, pred]` → `[removed_count]`.
pub fn emit_list_remove_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let list_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = list_slot + 1;
    let idx_slot = list_slot + 2;
    let removed_slot = list_slot + 3;
    let matched_slot = list_slot + 4;

    chunks[current].emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, list_slot, line);

    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, removed_slot, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    vybe_emitter::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_emitter::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, fn_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u8(Op::CALL_REF, 1, line);
    vybe_emitter::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, matched_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, matched_slot, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    vybe_emitter::collections::emit_remove_at(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, removed_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, removed_slot, line);

    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 1);
    vybe_emitter::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, idx_slot, line);

    chunks[current].emit_end(line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    chunks[current].emit_op_u16(Op::LOCAL_GET, removed_slot, line);
}

/// `Array.IndexOf(arr, value)` — search for `value`, return index or -1.
/// Stack: `[arr, value]` → `[index]`.
///
/// The C# profile previously routed this through `opcode:str_index_of`,
/// which only works for strings. Routing through this adapter uses the
/// shared array `indexOf` opcode that ECMA-262 §23.1.3.16 specifies.
pub fn emit_array_index_of(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_emitter::collections::emit_index_of(chunks, current, line);
}

// ── Static `Array.<HOF>(arr, fn)` predicates ──────────────────────────────
//
// .NET exposes `Array.Exists` / `Array.Find` / `Array.FindAll` /
// `Array.TrueForAll` / `Array.ConvertAll` / `Array.ForEach` as static
// methods that take the array as first arg; ECMA-262 expresses the
// equivalents as instance HOFs (`some`, `find`, `filter`, `every`,
// `map`, `forEach`). Each adapter inlines the loop using
// `compiler_common::loops` so the implementation matches what the
// instance-form HOF dispatch produces — same semantics, same bytecode
// shape.

fn alloc_locals(chunk: &mut Chunk, count: u16) -> u16 {
    chunk.alloc_scratch(count)
}

/// `Array.Exists(arr, pred)` → `arr.some(pred)`. Stack: `[arr, pred]` → `[bool]`.
pub fn emit_array_exists(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 4);
    let fn_slot = arr_slot + 1;
    let idx_slot = arr_slot + 2;
    let _result_slot = arr_slot + 3;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    vybe_emitter::loops::emit_any_every(
        chunks, current, fn_slot, arr_slot, idx_slot, /* is_some= */ true, line,
    );
}

/// `Array.TrueForAll(arr, pred)` → `arr.every(pred)`. Stack: `[arr, pred]` → `[bool]`.
pub fn emit_array_true_for_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 4);
    let fn_slot = arr_slot + 1;
    let idx_slot = arr_slot + 2;
    let _result_slot = arr_slot + 3;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    vybe_emitter::loops::emit_any_every(
        chunks, current, fn_slot, arr_slot, idx_slot, /* is_some= */ false, line,
    );
}

/// `Array.Find(arr, pred)` → `arr.find(pred)`. Stack: `[arr, pred]` → `[elem | default]`.
pub fn emit_array_find(chunks: &mut [Chunk], current: usize, line: u32) {
    // Use existing find loop pattern: iterate, return first matching elem.
    // We compose: filter then index 0 (matches the LINQ adapter's
    // `Array.Find` walker rewrite).
    let arr_slot = alloc_locals(&mut chunks[current], 6);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let found_slot = arr_slot + 5;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    vybe_emitter::loops::emit_filter(
        chunks,
        current,
        fn_slot,
        arr_slot,
        result_slot,
        idx_slot,
        elem_slot,
        line,
    );
    // The result is on the stack as an array — take element 0.
    core_wasm::i32_const(&mut chunks[current], line, 0);
    vybe_emitter::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_slot, line);
    let undef_idx = chunks[current].add_import("wasm:js-undefined", "test");
    chunks[current].emit_op_u16(Op::CALL_IMPORT, undef_idx, line);
    chunks[current].emit(1, line);
    chunks[current].emit_if(line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `Array.FindAll(arr, pred)` → `arr.filter(pred)`. Stack: `[arr, pred]` → `[array]`.
pub fn emit_array_find_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    vybe_emitter::loops::emit_filter(
        chunks,
        current,
        fn_slot,
        arr_slot,
        result_slot,
        idx_slot,
        elem_slot,
        line,
    );
}

/// `Array.BinarySearch(arr, value)` — .NET returns the index of `value` in
/// the sorted `arr`, or the bitwise complement of the insertion point (`~i`,
/// a negative) when it is absent. The shared `collections.binary_search`
/// delegates to `indexOf`, which can only signal a miss with `-1` and so
/// loses the insertion point — drive a spec-correct scan here instead.
/// Stack: `[arr, value]` → `[i32]`.
pub fn emit_array_binary_search(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let arr_slot = chunk.alloc_scratch(6);
    let value_slot = arr_slot + 1;
    let i_slot = arr_slot + 2;
    let len_slot = arr_slot + 3;
    let result_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;

    // Stash args (value on top).
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // len = arr.length
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op(Op::ARRAY_LENGTH, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len_slot, line);

    // i = 0
    core_wasm::i32_const(chunk, line, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    // result = ~len (value sorts after every element → insertion point == len)
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_emitter::ops::emit_dyn_neg(chunk, line);
    core_wasm::i32_const(chunk, line, -1);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);

    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);

    // while i < len — else break, keeping result = ~len
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len_slot, line);
    vybe_emitter::ops::emit_dyn_lt(chunk, line);
    vybe_emitter::ops::emit_dyn_not(chunk, line);
    chunk.emit_br_if(1, line);

    // elem = arr[i]
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, elem_slot, line);

    // if elem == value { result = i; break }  (br 2: if → loop → block)
    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_emitter::ops::emit_dyn_eq(chunk, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_br(2, line);
    chunk.emit_end(line);

    // if elem > value { result = ~i; break }  (insertion point is i)
    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_emitter::ops::emit_dyn_gt(chunk, line);
    vybe_emitter::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    vybe_emitter::ops::emit_dyn_neg(chunk, line);
    core_wasm::i32_const(chunk, line, -1);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunk.emit_br(2, line);
    chunk.emit_end(line);

    // i++
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_i32_const(1, line);
    vybe_emitter::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line); // loop
    chunk.patch_loop(loop_p);
    chunk.emit_end(line); // block
    chunk.patch_block(block_p);

    chunk.emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// `Array.ConvertAll(arr, selector)` → `arr.map(selector)`. Stack: `[arr, fn]` → `[array]`.
pub fn emit_array_convert_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 4);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    vybe_emitter::loops::emit_map(
        chunks,
        current,
        fn_slot,
        arr_slot,
        result_slot,
        idx_slot,
        line,
    );
}

/// `list.AddRange(other)` — append every element of `other` to `list`.
/// Stack: `[list, other]` → `[null]`. ECMA-262 has no direct mirror;
/// the closest is `list.push(...other)` which we lower to a for-each
/// loop because spread isn't always free in the call site.
pub fn emit_list_add_range(chunks: &mut [Chunk], current: usize, line: u32) {
    let list_slot = alloc_locals(&mut chunks[current], 3);
    let other_slot = list_slot + 1;
    let idx_slot = list_slot + 2;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, other_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, list_slot, line);

    // for elem in other: list.push(elem)
    let state = vybe_emitter::loops::emit_for_in_start(chunks, current, other_slot, idx_slot, line);
    let chunk = &mut chunks[current];
    let elem_slot = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    vybe_emitter::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line); // discard new length
    vybe_emitter::loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `Array.ForEach(arr, action)` → `arr.forEach(action)`. Stack: `[arr, fn]` → `[null]`.
pub fn emit_array_for_each(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 3);
    let fn_slot = arr_slot + 1;
    let idx_slot = arr_slot + 2;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);
    vybe_emitter::loops::emit_foreach(chunks, current, fn_slot, arr_slot, idx_slot, line);
}
