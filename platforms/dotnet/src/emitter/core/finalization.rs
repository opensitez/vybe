//! `System.GC`'s finalisation surface — the .NET SPELLING of a collection
//! point, and nothing else.
//!
//! ⛔ THE LIFECYCLE IS NOT HERE. Reading `ProtocolSlot::Destructor` and calling
//! it is class machinery whichever language asked for it, and it lives in
//! `primitives::classes` beside the queue and the drop site. This file briefly
//! owned that drain, which meant a .NET file read a protocol slot and invented
//! a spelling for the suppression bit — a platform deciding a shared fact,
//! which is the shape flexclassplan exists to remove.
//!
//! What a platform DOES own is what its API promises: that `GC.Collect` is a
//! collection point at all, and that `SuppressFinalize(Nothing)` throws rather
//! than being ignored. Both are .NET facts and both are stated here.

use vybe_compiler::primitives::class_slots::ValueSource;
use vybe_compiler::primitives::{classes, errors, globals, ops};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// `GC.Collect()` / `GC.WaitForPendingFinalizers()`.
///
/// Both spellings land here on purpose. .NET separates starting a collection
/// from waiting for its finalisers; with a queue rather than a collector there
/// is one operation, and doing it twice drains an already-empty queue.
///
/// Stack: `[]` → `[null]`.
pub fn emit_run_finalizers(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_count_collection(chunks, current, line);
    classes::emit_finalize_queue_drain(chunks, current, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// A drain IS this runtime's collection, so `GC.CollectionCount` reports a real
/// number rather than a constant.
fn emit_count_collection(chunks: &mut [Chunk], current: usize, line: u32) {
    globals::emit_read(&mut chunks[current], super::gc_adapter::COLLECTION_COUNT, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    globals::emit_read(&mut chunks[current], super::gc_adapter::COLLECTION_COUNT, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_end(line);
    globals::emit_write(&mut chunks[current], super::gc_adapter::COLLECTION_COUNT, line);
}

/// `GC.SuppressFinalize(o)` — stack `[o]` → `[null]`.
///
/// Idempotent, because a program may call it three times on the same object.
/// `Nothing` is an `ArgumentNullException` — MEASURED on .NET 10, not a guess:
/// the API rejects a null argument rather than ignoring it, and the corpus
/// catches it by type.
pub fn emit_suppress_finalize(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_suppressed(chunks, current, true, line);
}

/// `GC.ReRegisterForFinalize(o)` — the inverse, and null-rejecting for the same
/// reason.
pub fn emit_re_register_for_finalize(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_set_suppressed(chunks, current, false, line);
}

fn emit_set_suppressed(chunks: &mut [Chunk], current: usize, value: bool, line: u32) {
    let obj = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_TEE, obj, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    errors::emit_exception_new(
        &mut chunks[current],
        "ArgumentNullException",
        ValueSource::ConstStr("Value cannot be null.".into()),
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    classes::emit_set_finalize_suppressed(chunks, current, value, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
