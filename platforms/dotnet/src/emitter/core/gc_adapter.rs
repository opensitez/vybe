//! `System.GC` — the parts that are not the finalisation queue.
//!
//! ⚠ WHAT THESE MAY AND MAY NOT ANSWER. A `GC` member that reports a real
//! property of the collector is answered truthfully or not at all; nothing here
//! invents a measurement to make a test pass. `MaxGeneration` and `HeapCount`
//! are contract constants a client branches on and are answered; a byte total
//! we do not track is left as it was rather than given a plausible number.

use vybe_compiler::primitives::class_slots::{self, ValueSource};
use vybe_compiler::primitives::{collections, errors, globals};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

/// How many collections have run. Incremented by the finalisation drain, which
/// IS this runtime's collection, so `CollectionCount` reports a real number
/// rather than a constant.
pub const COLLECTION_COUNT: &str = "__vybe_gc_collection_count";

pub fn emit_gc_noop(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_gc_zero(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_i32_const(0, line);
}

/// `GC.MaxGeneration` — 2.
///
/// A CONTRACT CONSTANT, not a measurement: `.NET` has answered 2 on every
/// desktop CLR, and client code loops `For gen = 0 To GC.MaxGeneration`. It
/// answered 0, which makes that loop run once and makes
/// `GC.Collect(GC.MaxGeneration)` a gen-0 collection.
pub fn emit_gc_max_generation(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_i32_const(2, line);
}

/// `GC.GetTotalMemory` / `GetTotalAllocatedBytes` /
/// `GetAllocatedBytesForCurrentThread` — the module's linear memory, in bytes.
///
/// `memory.size` is the page count this module actually holds, so
/// `pages × 65536` is a fact about the running program rather than a plausible
/// CLR figure.
///
/// ⚠ IT IS CURRENTLY ZERO, AND THAT IS THE HONEST ANSWER. Our objects live in
/// GC structs and host values, not in linear memory, so this meters the wrong
/// heap — `memory.size` reports 0 pages for an ordinary program. .NET answers
/// all three of these greater than zero (measured on the SDK), and three
/// corpus tests assert `> 0`. Closing that needs the runtime to meter its own
/// heap; inventing a number here would make the tests pass and the answer
/// false.
pub fn emit_gc_get_total_memory(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::MEMORY_SIZE, 0, line);
    chunks[current].emit_op(Op::F64_CONVERT_I32_U, line);
    chunks[current].emit_f64_const(65536.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
}

/// `GC.GetGeneration(o)` — 0, or `ArgumentNullException` for `Nothing`.
///
/// ⛔ It answered −1 for null. `.NET` throws, and a caller that branches on a
/// generation number cannot tell a sentinel from a real answer.
pub fn emit_gc_get_generation(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    errors::emit_exception_new(
        chunk,
        "ArgumentNullException",
        ValueSource::ConstStr("Value cannot be null.".into()),
        line,
    );
    errors::emit_throw(chunk, line);
    chunk.emit_end(line);
    chunk.emit_i32_const(0, line);
}

/// `GC.AddMemoryPressure(n)` / `RemoveMemoryPressure(n)` — `n` must be
/// positive. We hold no unmanaged budget to adjust, so a valid call is a
/// no-op; an INVALID one is not, because rejecting it is the only observable
/// behaviour either method has.
pub fn emit_gc_memory_pressure(chunks: &mut [Chunk], current: usize, line: u32) {
    chunks[current].emit_f64_const(0.0, line);
    vybe_compiler::primitives::ops::emit_dyn_le(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    errors::emit_exception_new(
        &mut chunks[current],
        "ArgumentOutOfRangeException",
        ValueSource::ConstStr("Value must be positive.".into()),
        line,
    );
    errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// `GC.AllocateArray(Of T)(n)` / `AllocateUninitializedArray(Of T)(n)`, with an
/// optional `pinned` flag.
///
/// Pinning is what the two names differ by in .NET and it is not observable
/// here — an array of the requested length is the whole contract the corpus
/// checks, and it is one the runtime can actually keep.
pub fn emit_gc_allocate_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    // Trailing `pinned` is dropped; the length is the last value left.
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    collections::emit_new_with_length(chunks, current, line);
}

/// `GC.CollectionCount(gen)` — the real number of drains this program has run.
pub fn emit_gc_collection_count(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    globals::emit_read(&mut chunks[current], COLLECTION_COUNT, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_else(line);
    globals::emit_read(&mut chunks[current], COLLECTION_COUNT, line);
    chunks[current].emit_end(line);
}

/// `GC.GetGCMemoryInfo()` — an instance of the registered `GCMemoryInfo` type.
///
/// ⛔ NOT A BARE OBJECT WITH STAMPED KEYS. It briefly allocated one and wrote
/// `HeapCount` onto it, which meant inventing the spelling: VB folds a member
/// read to `heapcount`, so every field answered `undefined` until the key was
/// stamped TWICE. Two spellings for one member is the shape the plan exists to
/// remove — the fold belongs to the type, not to whoever writes the field.
///
/// `GCMemoryInfo` is now a registered `ClassType` whose members are argc-0
/// leaves, exactly like every other .NET type on the tree, and
/// `emit_class_construct` stamps the type so a read resolves through it.
pub fn emit_gc_memory_info(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    class_slots::emit_class_construct(&mut chunks[current], "GCMemoryInfo", &[], line);
}

/// A `GCMemoryInfo` member read: drop the receiver, answer the value.
///
/// Every one is a fact about THIS runtime rather than a plausible CLR number —
/// one heap, no collector pauses — and `TotalAvailableMemoryBytes` is the
/// ceiling we actually run under. `.NET 10` answers `PauseTimePercentage` 0 on
/// an idle process, measured on the SDK.
pub fn emit_gc_memory_info_member(
    chunks: &mut [Chunk],
    current: usize,
    value: f64,
    line: u32,
) {
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_f64_const(value, line);
}

/// `GC.GetTotalPauseDuration()` — a `TimeSpan` of zero.
///
/// Truthful rather than convenient: this runtime has no collector that pauses,
/// so the total pause IS zero, and the corpus asks only that it be `>= 0`.
pub fn emit_gc_total_pause_duration(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_f64_const(0.0, line);
    super::timespan_adapter::emit_timespan_from_milliseconds(chunks, current, line);
}
