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

use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;

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
    let arr_slot = chunk.local_count;
    let idx_slot = arr_slot + 1;
    let count_slot = arr_slot + 2;
    let i_slot = arr_slot + 3;
    let target_slot = arr_slot + 4;
    let elem_slot = arr_slot + 5;
    chunk.local_count = arr_slot + 6;

    // Stash args (top of stack first → reverse order)
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);   chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);   chunk.emit_op(Op::DROP, line);

    // i = 0
    chunk.emit_op(Op::I32_CONST_0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_op(Op::DROP, line);

    // Loop: while i < count { arr[idx + i] = default(existing_value); i++ }
    let block_p = chunk.emit_block(line);
    let (loop_p, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    chunk.emit_op(Op::DYN_NOT, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunk.emit_op(Op::REF_IS_BOOL, line);
    let bool_default = chunk.emit_jump(Op::BR_IF_TRUE, line);

    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    chunk.emit_op(Op::REF_IS_NUMBER, line);
    let number_default = chunk.emit_jump(Op::BR_IF_TRUE, line);

    chunk.emit_op(Op::NULL, line);
    let done_default = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(bool_default);
    chunk.emit_op(Op::FALSE, line);
    let done_bool = chunk.emit_jump(Op::BR, line);

    chunk.patch_jump(number_default);
    chunk.emit_op(Op::I32_CONST_0, line);

    chunk.patch_jump(done_bool);
    chunk.patch_jump(done_default);
    chunk.emit_op(Op::ARRAY_SET, line);
    chunk.emit_op(Op::DROP, line); // ARRAY_SET pushes the value; drop it

    // i++
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    let one = chunk.add_constant(Value::I32(1));
    chunk.emit_op_u16(Op::CONST, one, line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_br(0, line);
    chunk.emit_end(line); // end loop
    chunk.patch_loop(loop_p);
    chunk.emit_end(line); // end block
    chunk.patch_block(block_p);

    chunk.emit_op(Op::NULL, line);
}

/// `Array.Copy(src, dst, count)` — copy first `count` elements from
/// `src` to `dst`. Lowers to `__vybe_array_copy` stdlib chunk
/// (already bundled).
///
/// Stack on entry: `[src, dst, count]` ; Stack on exit: `[null]`
pub fn emit_array_copy(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::collections::emit_stdlib_call(chunks, current, "__vybe_array_copy", 3, line);
    // Stdlib chunk returns null; leave it on stack.
}

/// `Array.Resize(arr, newSize)` — extend or truncate `arr` to
/// `newSize` elements. Lowers to `__vybe_redim` stdlib chunk.
///
/// Stack on entry: `[arr, newSize]` ; Stack on exit: `[arr]` (the
/// stdlib helper returns the resized array; .NET `Array.Resize`
/// signature is by-ref but the bytecode propagates the value).
pub fn emit_array_resize(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::collections::emit_stdlib_call(chunks, current, "__vybe_redim", 2, line);
}

/// `Array.Sort(arr)` — in-place sort. Lowers to `__vybe_sort_in_place`
/// stdlib chunk.
///
/// Stack on entry: `[arr]` ; Stack on exit: `[null]` (sort is void in
/// .NET; the stdlib chunk returns the array but we drop it).
pub fn emit_array_sort(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::collections::emit_stdlib_call(chunks, current, "__vybe_sort_in_place", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `Array.Reverse(arr)` — in-place reverse. Stack: `[arr]` → `[null]`.
/// Reverse mutates the array; .NET's signature returns void.
pub fn emit_array_reverse(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::collections::emit_reverse(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `Array.IndexOf(arr, value)` — search for `value`, return index or -1.
/// Stack: `[arr, value]` → `[index]`.
///
/// The C# profile previously routed this through `opcode:str_index_of`,
/// which only works for strings. Routing through this adapter uses the
/// shared array `indexOf` opcode that ECMA-262 §23.1.3.16 specifies.
pub fn emit_array_index_of(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::emitter::collections::emit_index_of(chunks, current, line);
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
    let base = chunk.local_count;
    chunk.local_count = base + count;
    base
}

/// `Array.Exists(arr, pred)` → `arr.some(pred)`. Stack: `[arr, pred]` → `[bool]`.
pub fn emit_array_exists(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 4);
    let fn_slot = arr_slot + 1;
    let idx_slot = arr_slot + 2;
    let _result_slot = arr_slot + 3;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);  chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
    crate::emitter::loops::emit_any_every(chunks, current, fn_slot, arr_slot, idx_slot, /* is_some= */ true, line);
}

/// `Array.TrueForAll(arr, pred)` → `arr.every(pred)`. Stack: `[arr, pred]` → `[bool]`.
pub fn emit_array_true_for_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 4);
    let fn_slot = arr_slot + 1;
    let idx_slot = arr_slot + 2;
    let _result_slot = arr_slot + 3;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);  chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
    crate::emitter::loops::emit_any_every(chunks, current, fn_slot, arr_slot, idx_slot, /* is_some= */ false, line);
}

/// `Array.Find(arr, pred)` → `arr.find(pred)`. Stack: `[arr, pred]` → `[elem | default]`.
pub fn emit_array_find(chunks: &mut [Chunk], current: usize, line: u32) {
    // Use existing find loop pattern: iterate, return first matching elem.
    // We compose: filter then index 0 (matches the LINQ adapter's
    // `Array.Find` walker rewrite).
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);  chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
    crate::emitter::loops::emit_filter(chunks, current, fn_slot, arr_slot, result_slot, idx_slot, elem_slot, line);
    // The result is on the stack as an array — take element 0.
    chunks[current].emit_op(Op::I32_CONST_0, line);
    crate::emitter::collections::emit_get(chunks, current, line);
}

/// `Array.FindAll(arr, pred)` → `arr.filter(pred)`. Stack: `[arr, pred]` → `[array]`.
pub fn emit_array_find_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 5);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let elem_slot = arr_slot + 4;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);  chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
    crate::emitter::loops::emit_filter(chunks, current, fn_slot, arr_slot, result_slot, idx_slot, elem_slot, line);
}

/// `Array.ConvertAll(arr, selector)` → `arr.map(selector)`. Stack: `[arr, fn]` → `[array]`.
pub fn emit_array_convert_all(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 4);
    let fn_slot = arr_slot + 1;
    let result_slot = arr_slot + 2;
    let idx_slot = arr_slot + 3;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);  chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
    crate::emitter::loops::emit_map(chunks, current, fn_slot, arr_slot, result_slot, idx_slot, line);
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
    chunk.emit_op_u16(Op::LOCAL_SET, other_slot, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, list_slot, line);  chunk.emit_op(Op::DROP, line);

    // for elem in other: list.push(elem)
    let state = crate::emitter::loops::emit_for_in_start(chunks, current, other_slot, idx_slot, line);
    let chunk = &mut chunks[current];
    let elem_slot = {
        let s = chunk.local_count;
        chunk.local_count = s + 1;
        s
    };
    chunk.emit_op_u16(Op::LOCAL_SET, elem_slot, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, list_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, elem_slot, line);
    crate::emitter::collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line); // discard new length
    crate::emitter::loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op(Op::NULL, line);
}

/// `Array.ForEach(arr, action)` → `arr.forEach(action)`. Stack: `[arr, fn]` → `[null]`.
pub fn emit_array_for_each(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 3);
    let fn_slot = arr_slot + 1;
    let idx_slot = arr_slot + 2;
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, fn_slot, line);  chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line); chunk.emit_op(Op::DROP, line);
    crate::emitter::loops::emit_foreach(chunks, current, fn_slot, arr_slot, idx_slot, line);
}
