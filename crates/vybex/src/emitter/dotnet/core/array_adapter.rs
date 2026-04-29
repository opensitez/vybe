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

/// `Array.Clear(arr, idx, count)` — set `count` elements starting at
/// `idx` to `Null`. Inline loop using `LOCAL_GET` / `ARRAY_SET`.
///
/// Stack on entry: `[arr, idx, count]` ; Stack on exit: `[null]`
pub fn emit_array_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    // Allocate scratch slots: arr, idx, count, i (counter)
    let arr_slot = chunk.local_count;
    let idx_slot = arr_slot + 1;
    let count_slot = arr_slot + 2;
    let i_slot = arr_slot + 3;
    chunk.local_count = arr_slot + 4;

    // Stash args (top of stack first → reverse order)
    chunk.emit_op_u16(Op::LOCAL_SET, count_slot, line); chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, idx_slot, line);   chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_SET, arr_slot, line);   chunk.emit_op(Op::DROP, line);

    // i = 0
    chunk.emit_op(Op::I32_CONST_0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i_slot, line);
    chunk.emit_op(Op::DROP, line);

    // Loop: while i < count { arr[idx + i] = Null; i++ }
    let _block_p = chunk.emit_block(line);
    let (_loop_p, _) = chunk.emit_loop_s(line);

    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunk.emit_op(Op::DYN_LT, line);
    chunk.emit_op(Op::DYN_NOT, line);
    chunk.emit_br_if(1, line);

    // arr[idx + i] = Null
    chunk.emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, idx_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunk.emit_op(Op::DYN_ADD, line);
    chunk.emit_op(Op::NULL, line);
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
    chunk.emit_end(line); // end block

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
