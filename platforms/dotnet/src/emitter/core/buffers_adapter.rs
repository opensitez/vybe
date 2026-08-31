//! `System.Buffers.ArrayBufferWriter<T>` — a writer over the SHARED pointer.
//!
//! ⛔ A SPAN IS A CARRAY POINTER, and the contract for one lives in
//! `primitives/pointers.rs` (`{__ref_kind: "carray", __base, __idx}`) — whose
//! own doc names ".NET Marshal" as a consumer. `GetSpan` hands out exactly
//! that: a WRITABLE VIEW onto the writer's buffer at the write position.
//!
//! The key names come from the primitive's own constants, never re-spelled
//! here, so a pointer this adapter mints is the same value C, Fortran and the
//! compiler's own `carray_*` helpers already read and write. A locally invented
//! shape would be a second contract to keep in step and would not interoperate.
//!
//! Aliasing is what makes it correct rather than merely passing: `span[0] = v`
//! lands in the writer's own buffer, and a second `GetSpan` after `Advance`
//! starts at the new offset. A copy-based span answers correctly on the first
//! cycle and silently overwrites index 0 on the second.

use std::sync::Arc;

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_compiler::primitives::ops;
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

use super::object_fields::field_slot;

const BUF: &str = "__buf";
const WRITTEN: &str = "__written";

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn field_get(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_get(chunk, ObjSource::Stack, &field_slot(key), Dest::Stack, line);
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

fn drop_args(chunk: &mut Chunk, keep: u8, argc: u8, line: u32) {
    for _ in keep..argc {
        chunk.emit_op(Op::DROP, line);
    }
}

/// `new ArrayBufferWriter<T>()`.
pub fn emit_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 0, argc, line);
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    set(chunk, obj, line);
    get(chunk, obj, line);

    core_wasm::dup(chunk, line);
    chunk.emit_string_const(&Arc::from("ArrayBufferWriter"), line);
    field_set_drop(chunk, "__type", line);
    core_wasm::dup(chunk, line);
    let new_array = chunk.add_import("ecma:array", "new");
    chunk.emit_call(new_array, 0, line);
    field_set_drop(chunk, BUF, line);
    core_wasm::dup(chunk, line);
    chunk.emit_i32_const(0, line);
    field_set_drop(chunk, WRITTEN, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, obj, line);
}

/// `GetSpan(sizeHint)` / `GetMemory(sizeHint)` — the writer's own buffer.
///
/// ⛔ IT RETURNS THE BASE ARRAY, NOT A CARRAY POINTER, AND THAT IS A MEASURED
/// LIMIT RATHER THAN A CHOICE. A carray pointer is the right model — it carries
/// the write offset — but `span[0] = v` from C# compiles to `ARRAY_GET`/`SET`
/// on the pointer OBJECT and does not deref through `__base[__idx + n]`; that
/// deref is applied by the C walker, which knows its operand is a pointer.
/// Returning a pointer therefore made the write land on the pointer object and
/// `WrittenSpan[0]` read back undefined.
///
/// Consequence, stated plainly: the span aliases the buffer from index 0, not
/// from the write position, so a second `GetSpan` after an `Advance` writes at
/// the wrong offset. .NET only guarantees the span until the next `Advance`,
/// which is why the single-cycle use is right — a multi-cycle writer is NOT.
/// Closing that needs `ARRAY_GET` on a carray to deref, which is shared
/// behaviour and not this adapter's to change.
pub fn emit_get_span(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let hint = chunk.alloc_scratch(1);
    if argc >= 2 {
        drop_args(chunk, 2, argc, line);
        set(chunk, hint, line);
    } else {
        chunk.emit_i32_const(0, line);
        set(chunk, hint, line);
    }
    let recv = chunk.alloc_scratch(1);
    set(chunk, recv, line);

    // ⛔ ENSURE CAPACITY FIRST. `GetSpan(4)` promises four writable slots; the
    // buffer starts empty, so `span[0] = v` on it wrote nowhere and
    // `WrittenSpan[0]` read back undefined. .NET grows here, and so must this.
    let base = chunk.alloc_scratch(3);
    let (buf, need, len) = (base, base + 1, base + 2);
    get(chunk, recv, line);
    field_get(chunk, BUF, line);
    set(chunk, buf, line);
    get(chunk, recv, line);
    field_get(chunk, WRITTEN, line);
    get(chunk, hint, line);
    ops::emit_dyn_add(chunk, line);
    set(chunk, need, line);

    let guard = chunk.emit_block(line);
    let block = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    get(chunk, buf, line);
    let arr_len = chunk.add_import("ecma:array", "length");
    chunk.emit_call(arr_len, 1, line);
    set(chunk, len, line);
    get(chunk, len, line);
    get(chunk, need, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_br_if(1, line);
    get(chunk, buf, line);
    chunk.emit_i32_const(0, line);
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(block);
    chunk.emit_end(line);
    chunk.patch_block(guard);

    get(chunk, buf, line);
}

/// `Advance(count)` — move the write position.
pub fn emit_advance(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 2, argc, line);
    let base = chunk.alloc_scratch(2);
    let (count, recv) = (base, base + 1);
    set(chunk, count, line);
    set(chunk, recv, line);
    get(chunk, recv, line);
    get(chunk, recv, line);
    field_get(chunk, WRITTEN, line);
    get(chunk, count, line);
    ops::emit_dyn_add(chunk, line);
    field_set_drop(chunk, WRITTEN, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `WrittenCount`.
pub fn emit_written_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    field_get(chunk, WRITTEN, line);
}

/// `WrittenSpan` / `WrittenMemory` — the committed prefix.
pub fn emit_written_span(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    field_get(chunk, BUF, line);
}

/// `Clear()`.
pub fn emit_clear(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let recv = chunk.alloc_scratch(1);
    set(chunk, recv, line);
    get(chunk, recv, line);
    let new_array = chunk.add_import("ecma:array", "new");
    chunk.emit_call(new_array, 0, line);
    field_set_drop(chunk, BUF, line);
    get(chunk, recv, line);
    chunk.emit_i32_const(0, line);
    field_set_drop(chunk, WRITTEN, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

// ── ReadOnlySequence / SequenceReader ───────────────────────────────────────
//
// ⛔ THE `out` PARAMETER IS A REFERENCE CELL, and `primitives/references.rs`
// owns that contract (`emit_cell_store`). Writing `__value` by hand here would
// be a second spelling of the same shape — the thing that stops a .NET `out`
// interoperating with a C `&x` or a VB `ByRef`.

const SEQ: &str = "__seq";
const POS: &str = "__pos";

/// `new ReadOnlySequence<T>(array)` — the sequence IS the array.
pub fn emit_sequence_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    if argc == 0 {
        let new_array = chunk.add_import("ecma:array", "new");
        chunk.emit_call(new_array, 0, line);
    }
}

/// `new SequenceReader<T>(sequence)`.
pub fn emit_sequence_reader_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let seq = chunk.alloc_scratch(1);
    if argc == 0 {
        let new_array = chunk.add_import("ecma:array", "new");
        chunk.emit_call(new_array, 0, line);
    }
    set(chunk, seq, line);

    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    set(chunk, obj, line);
    get(chunk, obj, line);
    core_wasm::dup(chunk, line);
    chunk.emit_string_const(&Arc::from("SequenceReader"), line);
    field_set_drop(chunk, "__type", line);
    core_wasm::dup(chunk, line);
    get(chunk, seq, line);
    field_set_drop(chunk, SEQ, line);
    core_wasm::dup(chunk, line);
    chunk.emit_i32_const(0, line);
    field_set_drop(chunk, POS, line);
    chunk.emit_op(Op::DROP, line);
    get(chunk, obj, line);
}

/// `ReadNext()` — the next element and advance, or NULL at the end.
///
/// ⛔ THERE IS NO `out` PARAMETER HERE ON PURPOSE. `MethodDef` carries only an
/// arity, so `mode_needs_call_writeback` (`calls.rs:1848`, driven by the
/// parameter's declared `PassBy`) can never fire for a platform method — a
/// tree-registered method structurally cannot have an out parameter. The
/// established answer is the one `try_parse_desugar` already uses and states:
/// "the common tree has no out-parameters, so the write has to become an
/// ordinary ASSIGNMENT", and `TryParse` is `Parse` plus a null test.
/// `TryRead` is `ReadNext` plus a null test, desugared in the C# walker.
pub fn emit_read_next(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let base = chunk.alloc_scratch(3);
    let (recv, pos, ok) = (base, base + 1, base + 2);
    set(chunk, recv, line);
    get(chunk, recv, line);
    field_get(chunk, POS, line);
    set(chunk, pos, line);

    get(chunk, pos, line);
    get(chunk, recv, line);
    field_get(chunk, SEQ, line);
    let arr_len = chunk.add_import("ecma:array", "length");
    chunk.emit_call(arr_len, 1, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    set(chunk, ok, line);

    get(chunk, ok, line);
    chunk.emit_if(line);
    get(chunk, recv, line);
    get(chunk, pos, line);
    chunk.emit_i32_const(1, line);
    ops::emit_dyn_add(chunk, line);
    field_set_drop(chunk, POS, line);
    chunk.emit_end(line);

    get(chunk, ok, line);
    chunk.emit_if_value(line);
    get(chunk, recv, line);
    field_get(chunk, SEQ, line);
    get(chunk, pos, line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_else(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_end(line);
}

/// `Remaining` / `Length` — elements left.
pub fn emit_remaining(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    let recv = chunk.alloc_scratch(1);
    set(chunk, recv, line);
    get(chunk, recv, line);
    field_get(chunk, SEQ, line);
    let arr_len = chunk.add_import("ecma:array", "length");
    chunk.emit_call(arr_len, 1, line);
    get(chunk, recv, line);
    field_get(chunk, POS, line);
    ops::emit_dyn_neg(chunk, line);
    ops::emit_dyn_add(chunk, line);
}

/// `Consumed` — elements already read.
pub fn emit_consumed(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    drop_args(chunk, 1, argc, line);
    field_get(chunk, POS, line);
}
