//! `System.Span<T>` members whose shape does NOT line up with a single ECMA
//! array method — bytecode-only.
//!
//! Everything that *does* line up is a plain `HostCall` in
//! `component_classes_span.rs` (`Slice(start)` → `slice(arr, start)`, and so
//! on). What lands here needs operands the call site doesn't supply
//! (`Clear` zeroes with `default(T)`), operands in the other order
//! (`CopyTo`'s receiver is the SOURCE, `set_range`'s is the destination), or a
//! loop with no ECMA analogue at all (`TrimStart`, `Mismatch`).
//!
//! Pure WASM over the shared `vybe_compiler::emitter` primitives — a Span is just an
//! array here, so these compose the same helpers `List<T>` uses.

use super::array_adapter;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};
use vybe_compiler::primitives::instructions::core_wasm;

fn alloc_locals(chunk: &mut Chunk, n: u16) -> u16 {
    chunk.alloc_scratch(n)
}

fn key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(Value::String(name.into()))
}

fn set_field(chunk: &mut Chunk, object_slot: u16, field: &str, value_slot: u16, line: u32) {
    let field_key = key(chunk, field);
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, field_key, line);
    chunk.emit_op(Op::DROP, line);
}

fn set_string_field(chunk: &mut Chunk, object_slot: u16, field: &str, value: &str, line: u32) {
    let value_slot = chunk.alloc_scratch(1);
    chunk.emit_string_const(value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    set_field(chunk, object_slot, field, value_slot, line);
}

fn get_field(chunk: &mut Chunk, object_slot: u16, field: &str, line: u32) {
    let field_key = key(chunk, field);
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, field_key, line);
}

fn emit_array_segment_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    array_slot: u16,
    offset_slot: u16,
    count_slot: u16,
    line: u32,
) {
    let object_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, object_slot, line);
    set_field(&mut chunks[current], object_slot, "Array", array_slot, line);
    set_field(&mut chunks[current], object_slot, "array", array_slot, line);
    set_field(&mut chunks[current], object_slot, "Offset", offset_slot, line);
    set_field(&mut chunks[current], object_slot, "offset", offset_slot, line);
    set_field(&mut chunks[current], object_slot, "Count", count_slot, line);
    set_field(&mut chunks[current], object_slot, "count", count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, object_slot, line);
}

fn emit_array_with_length(chunks: &mut [Chunk], current: usize, count_slot: u16, line: u32) {
    let idx = chunks[current].add_import("ecma:array", "newWithLength");
    chunks[current].emit_op_u16(Op::LOCAL_GET, count_slot, line);
    chunks[current].emit_op_u16(Op::CALL_IMPORT, idx, line);
    chunks[current].emit(1, line);
}

pub fn emit_array_pool_shared(chunks: &mut [Chunk], current: usize, line: u32) {
    let pool_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pool_slot, line);
    set_string_field(
        &mut chunks[current],
        pool_slot,
        "__type",
        "ArrayPool",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, pool_slot, line);
}

pub fn emit_array_pool_rent(chunks: &mut [Chunk], current: usize, line: u32) {
    let count_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_array_with_length(chunks, current, count_slot, line);
}

pub fn emit_array_pool_rent_static(chunks: &mut [Chunk], current: usize, line: u32) {
    let count_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    emit_array_with_length(chunks, current, count_slot, line);
}

pub fn emit_array_pool_return(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_memory_pool_shared(chunks: &mut [Chunk], current: usize, line: u32) {
    let pool_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pool_slot, line);
    set_string_field(
        &mut chunks[current],
        pool_slot,
        "__type",
        "MemoryPool",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, pool_slot, line);
}

pub fn emit_memory_pool_rent(chunks: &mut [Chunk], current: usize, line: u32) {
    let count_slot = chunks[current].alloc_scratch(3);
    let array_slot = count_slot + 1;
    let owner_slot = count_slot + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_array_with_length(chunks, current, count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, owner_slot, line);
    set_string_field(
        &mut chunks[current],
        owner_slot,
        "__type",
        "MemoryPoolOwner",
        line,
    );
    set_field(&mut chunks[current], owner_slot, "Memory", array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, owner_slot, line);
}

pub fn emit_memory_pool_rent_static(chunks: &mut [Chunk], current: usize, line: u32) {
    let count_slot = chunks[current].alloc_scratch(3);
    let array_slot = count_slot + 1;
    let owner_slot = count_slot + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    emit_array_with_length(chunks, current, count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, owner_slot, line);
    set_string_field(
        &mut chunks[current],
        owner_slot,
        "__type",
        "MemoryPoolOwner",
        line,
    );
    set_field(&mut chunks[current], owner_slot, "Memory", array_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, owner_slot, line);
}

/// `new Span<T>(array)` / `new Span<T>(array, start, length)` — and the same
/// for `Memory<T>`. A class has ONE constructor slot but these two arities
/// differ in meaning, so the branch is on the argument count already on the
/// stack:
///   1 arg  → identity: a span IS the array, so window the whole thing.
///   3 args → `[array, start, length]` → `collections.get_range`'s
///            `[array, index, count]`.
///
/// Stack on entry: `[array]` or `[array, start, length]` ; exit: `[array]`
pub fn emit_span_ctor(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 3 {
        vybe_compiler::primitives::collections::emit_get_range(chunks, current, line);
    }
    // argc == 1 (or anything else): the array is already the result — leave it.
}

/// `new ArraySegment<T>(array)` / `(array, offset, count)`.
pub fn emit_array_segment_ctor(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let array_slot = alloc_locals(&mut chunks[current], 3);
    let offset_slot = array_slot + 1;
    let count_slot = array_slot + 2;
    if argc >= 3 {
        chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, offset_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
        core_wasm::i32_const(&mut chunks[current], line, 0);
        chunks[current].emit_op_u16(Op::LOCAL_SET, offset_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, array_slot, line);
        vybe_compiler::primitives::collections::emit_len(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    }
    emit_array_segment_from_slots(chunks, current, array_slot, offset_slot, count_slot, line);
}

pub fn emit_array_segment_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    let array_slot = alloc_locals(&mut chunks[current], 3);
    let offset_slot = array_slot + 1;
    let count_slot = array_slot + 2;
    vybe_compiler::primitives::collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, array_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, offset_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    emit_array_segment_from_slots(chunks, current, array_slot, offset_slot, count_slot, line);
}

fn emit_segment_abs_index(chunk: &mut Chunk, segment_slot: u16, index_slot: u16, line: u32) {
    get_field(chunk, segment_slot, "Offset", line);
    chunk.emit_op_u16(Op::LOCAL_GET, index_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
}

/// `segment.Item(i)`.
pub fn emit_array_segment_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let segment_slot = alloc_locals(&mut chunks[current], 2);
    let index_slot = segment_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, segment_slot, line);
    get_field(&mut chunks[current], segment_slot, "Array", line);
    emit_segment_abs_index(&mut chunks[current], segment_slot, index_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
}

/// `segment.SetItem(i, value)`.
pub fn emit_array_segment_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let segment_slot = alloc_locals(&mut chunks[current], 3);
    let index_slot = segment_slot + 1;
    let value_slot = segment_slot + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, segment_slot, line);
    get_field(&mut chunks[current], segment_slot, "Array", line);
    emit_segment_abs_index(&mut chunks[current], segment_slot, index_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
}

/// `segment.Slice(start, count)`.
pub fn emit_array_segment_slice(chunks: &mut [Chunk], current: usize, line: u32) {
    let segment_slot = alloc_locals(&mut chunks[current], 4);
    let start_slot = segment_slot + 1;
    let count_slot = segment_slot + 2;
    let offset_slot = segment_slot + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, count_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, segment_slot, line);
    get_field(&mut chunks[current], segment_slot, "Offset", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, offset_slot, line);
    get_field(&mut chunks[current], segment_slot, "Array", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, segment_slot, line);
    emit_array_segment_from_slots(chunks, current, segment_slot, offset_slot, count_slot, line);
}

/// `segment.ToArray()`.
pub fn emit_array_segment_to_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let segment_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, segment_slot, line);
    get_field(&mut chunks[current], segment_slot, "Array", line);
    get_field(&mut chunks[current], segment_slot, "Offset", line);
    get_field(&mut chunks[current], segment_slot, "Count", line);
    vybe_compiler::primitives::collections::emit_get_range(chunks, current, line);
}

/// `segment.CopyTo(dest)`.
pub fn emit_array_segment_copy_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let segment_slot = alloc_locals(&mut chunks[current], 2);
    let dest_slot = segment_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dest_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, segment_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest_slot, line);
    core_wasm::i32_const(&mut chunks[current], line, 0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, segment_slot, line);
    emit_array_segment_to_array(chunks, current, line);
    vybe_compiler::primitives::collections::emit_set_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `segment.Equals(other)` — value equality over the visible range.
pub fn emit_array_segment_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    let left_slot = alloc_locals(&mut chunks[current], 2);
    let right_slot = left_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, right_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, left_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_slot, line);
    emit_array_segment_to_array(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_slot, line);
    emit_array_segment_to_array(chunks, current, line);
    vybe_compiler::primitives::collections::emit_sequence_equal(chunks, current, line);
}

pub fn emit_buffer_byte_length(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    core_wasm::f64_const(&mut chunks[current], line, 4.0);
    chunks[current].emit_op(Op::F64_MUL, line);
}

pub fn emit_buffer_get_byte(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
}

pub fn emit_buffer_set_byte(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `span.IsEmpty` → `span.Length == 0`.
///
/// The comparison is branched back into a real JS boolean rather than handed
/// out as-is: `emit_dyn_eq` leaves a raw `i32`, and `Console.WriteLine` needs
/// an actual boolean to render .NET's `True`/`False`.
///
/// Stack on entry: `[arr]` ; Stack on exit: `[bool]`
pub fn emit_span_is_empty(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    let chunk = &mut chunks[current];
    core_wasm::f64_const(chunk, line, 0.0);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    core_wasm::bool_const(chunk, line, true);
    chunk.emit_else(line);
    core_wasm::bool_const(chunk, line, false);
    chunk.emit_end(line);
}

/// `span.Clear()` — .NET sets every slot to `default(T)`, which is NOT the
/// same as emptying the array (`collections.clear`, i.e. `length = 0`). The
/// element-type default is inferred per slot from the runtime category, which
/// is exactly what `Array.Clear(arr, idx, count)` already does — so this
/// supplies the range it wants: the whole span.
///
/// Stack on entry: `[arr]` ; Stack on exit: `[null]`
pub fn emit_span_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // [arr, 0, arr.length] — the contract `emit_array_clear` expects.
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    array_adapter::emit_array_clear(chunks, current, line);
}

/// Emit `dst[0..] = src`, consuming nothing from the stack.
///
/// `collections.set_range` takes `[array, index, src_array]` with the array as
/// the DESTINATION — the reverse of `src.CopyTo(dst)`, hence the reorder.
fn emit_copy_into(chunks: &mut [Chunk], current: usize, src_slot: u16, dst_slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, dst_slot, line);
    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src_slot, line);
    vybe_compiler::primitives::collections::emit_set_range(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `src.CopyTo(dst)` — copy every element of the receiver into `dst`.
///
/// Stack on entry: `[src, dst]` ; Stack on exit: `[null]`
pub fn emit_span_copy_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let src_slot = alloc_locals(&mut chunks[current], 2);
    let dst_slot = src_slot + 1;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dst_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src_slot, line);

    emit_copy_into(chunks, current, src_slot, dst_slot, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `src.TryCopyTo(dst)` — copy only if `dst` has room, reporting whether it
/// did. .NET's contract is all-or-nothing: a too-small destination is left
/// untouched.
///
/// Stack on entry: `[src, dst]` ; Stack on exit: `[bool]`
pub fn emit_span_try_copy_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let src_slot = alloc_locals(&mut chunks[current], 3);
    let dst_slot = src_slot + 1;
    let result_slot = src_slot + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, dst_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src_slot, line);

    // dst.Length >= src.Length
    chunks[current].emit_op_u16(Op::LOCAL_GET, dst_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_ge(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);

    chunks[current].emit_if(line);
    emit_copy_into(chunks, current, src_slot, dst_slot, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    // A `() -> ()` block can't carry a value out; reload it.
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Shared body for `TrimStart(value)` / `TrimEnd(value)`.
///
/// Walks a cursor from one end past every element equal to `value`, then
/// slices the survivors out. `from_start` picks which end and which slice.
fn emit_span_trim(chunks: &mut [Chunk], current: usize, from_start: bool, line: u32) {
    let arr_slot = alloc_locals(&mut chunks[current], 4);
    let val_slot = arr_slot + 1;
    let cursor_slot = arr_slot + 2;
    let probe_slot = arr_slot + 3;
    chunks[current].emit_op_u16(Op::LOCAL_SET, val_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr_slot, line);

    // TrimStart walks up from 0; TrimEnd walks down from Length.
    if from_start {
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
        vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor_slot, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);

    // Bound: TrimStart stops at Length, TrimEnd stops at 0.
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor_slot, line);
    if from_start {
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
        vybe_compiler::primitives::collections::emit_len(chunks, current, line);
        vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    } else {
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
        vybe_compiler::primitives::ops::emit_dyn_gt(&mut chunks[current], line);
    }
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    // probe = arr[from_start ? cursor : cursor - 1]
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor_slot, line);
    if !from_start {
        core_wasm::f64_const(&mut chunks[current], line, -1.0);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    }
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, probe_slot, line);

    // Stop at the first element that isn't the trimmed value.
    chunks[current].emit_op_u16(Op::LOCAL_GET, probe_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, val_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_ne(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    // TrimStart walks the cursor up, TrimEnd walks it down.
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor_slot, line);
    core_wasm::f64_const(
        &mut chunks[current],
        line,
        if from_start { 1.0 } else { -1.0 },
    );
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor_slot, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    // TrimStart keeps [cursor..], TrimEnd keeps [..cursor].
    //
    // `collections.slice` wants a REAL end index — handing it `undefined` to
    // mean "to the end" yields an EMPTY array, not the tail, so the end is
    // materialized explicitly.
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
    if from_start {
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor_slot, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, arr_slot, line);
        vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    } else {
        core_wasm::f64_const(&mut chunks[current], line, 0.0);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor_slot, line);
    }
    vybe_compiler::primitives::collections::emit_slice(chunks, current, line);
}

/// `span.TrimStart(value)` — drop leading elements equal to `value`.
///
/// Stack on entry: `[arr, value]` ; Stack on exit: `[array]`
pub fn emit_span_trim_start(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_span_trim(chunks, current, true, line);
}

/// `span.TrimEnd(value)` — drop trailing elements equal to `value`.
///
/// Stack on entry: `[arr, value]` ; Stack on exit: `[array]`
pub fn emit_span_trim_end(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_span_trim(chunks, current, false, line);
}

/// `a.Mismatch(b)` — index of the first differing element, or `-1` when the
/// spans are equal. Where one span is a strict prefix of the other, .NET
/// reports the shorter length: that IS the first index at which they differ.
///
/// Stack on entry: `[a, b]` ; Stack on exit: `[number]`
pub fn emit_span_mismatch(chunks: &mut [Chunk], current: usize, line: u32) {
    let a_slot = alloc_locals(&mut chunks[current], 7);
    let b_slot = a_slot + 1;
    let i_slot = a_slot + 2;
    let a_len_slot = a_slot + 3;
    let b_len_slot = a_slot + 4;
    let result_slot = a_slot + 5;
    let found_slot = a_slot + 6;
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_slot, line);

    core_wasm::bool_const(&mut chunks[current], line, false);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b_len_slot, line);

    // Scanning stops at the shorter length; that index is also the answer
    // when everything before it matched but the lengths differ.
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    // result_slot now holds min(a.Length, b.Length) — the scan bound.

    core_wasm::f64_const(&mut chunks[current], line, 0.0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    let block_p = chunks[current].emit_block(line);
    let (loop_p, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    vybe_compiler::primitives::ops::emit_dyn_ne(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    core_wasm::bool_const(&mut chunks[current], line, true);
    chunks[current].emit_op_u16(Op::LOCAL_SET, found_slot, line);
    chunks[current].emit_br(2, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i_slot, line);
    core_wasm::f64_const(&mut chunks[current], line, 1.0);
    vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i_slot, line);

    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_p);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block_p);

    // Falling out of the scan WITHOUT a break means every common element
    // matched: equal lengths make the spans equal (-1), otherwise the bound
    // already in `result_slot` is the answer. The `found` guard matters —
    // a real mismatch at index i must not be overwritten by the -1 path just
    // because the two spans happen to be the same length.
    chunks[current].emit_op_u16(Op::LOCAL_GET, found_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, a_len_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, b_len_slot, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    core_wasm::f64_const(&mut chunks[current], line, -1.0);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}
