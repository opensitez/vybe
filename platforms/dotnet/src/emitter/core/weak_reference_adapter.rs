//! `System.WeakReference` and `WeakReference(Of T)`, shared by .NET languages.
//!
//! The type had no `ClassType` anywhere, so every member answered "undefined is
//! not callable". Registered here as DATA the way `Lazy(Of T)` and
//! `TaskCompletionSource` are, so the common resolver answers it from the tree.
//!
//! Shape: `{__type:"WeakReference", __wr_target: <obj>}`.
//!
//! ## The reference is STRONG, and that is a real limitation
//!
//! Nothing here can observe collection: the target is held normally, so
//! `IsAlive` is "the target is not null" rather than "the GC has not taken it".
//! A test that sets the only reference to `Nothing`, calls `GC.Collect()` and
//! expects `IsAlive` to flip to False CANNOT pass, and is left failing rather
//! than special-cased into passing — a `GC.Collect()` that clears weak targets
//! would be a lie about every other program's lifetimes.
//!
//! `trackResurrection` is accepted and dropped for the same reason: it selects
//! whether the target survives finalisation, and no finaliser runs here.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ResolvedSlot, ValueSource};
use vybe_compiler::primitives::{globals, ops};

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const TYPE_KEY: &str = "__type";
const TARGET_KEY: &str = "__wr_target";
const EPOCH_KEY: &str = "__wr_epoch";
const TYPE_NAME: &str = "WeakReference";

fn weakref_slot(key: &str) -> ResolvedSlot {
    ResolvedSlot::Dynamic(ValueSource::ConstStr(key.to_string()))
}

fn emit_current_gc_epoch(chunk: &mut Chunk, line: u32) {
    globals::emit_read(chunk, super::gc_adapter::COLLECTION_COUNT, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_else(line);
    globals::emit_read(chunk, super::gc_adapter::COLLECTION_COUNT, line);
    chunk.emit_end(line);
}

/// `New WeakReference(obj)`, `New WeakReference(obj, trackResurrection)` and
/// `New WeakReference(Of T)(obj)`.
///
/// Stack on entry: `[obj]` or `[obj, trackResurrection]` ; on exit: `[wr]`.
pub fn emit_weakref_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(3);
    let (obj_slot, target_slot, epoch_slot) = (base, base + 1, base + 2);

    if argc >= 2 {
        // `trackResurrection` — see the module note.
        chunk.emit_op(Op::DROP, line);
    }
    if argc >= 1 {
        chunk.emit_op_u16(Op::LOCAL_SET, target_slot, line);
    } else {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunk.emit_op_u16(Op::LOCAL_SET, target_slot, line);
    }

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_string_const(TYPE_NAME, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &weakref_slot(TYPE_KEY),
        ValueSource::Stack,
        line,
    );
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &weakref_slot(TARGET_KEY),
        ValueSource::Stack,
        line,
    );
    chunk.emit_op(Op::DROP, line);

    emit_current_gc_epoch(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, epoch_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, epoch_slot, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &weakref_slot(EPOCH_KEY),
        ValueSource::Stack,
        line,
    );
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `wr.Target` — and the zero-arity core the `TryGetTarget` out-param desugar
/// calls. One emitter for both spellings on purpose.
///
/// Stack on entry: `[wr]` ; on exit: `[target]`.
pub fn emit_weakref_target(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(4);
    let (wr_slot, target_slot, epoch_slot, current_epoch_slot) = (base, base + 1, base + 2, base + 3);

    chunk.emit_op_u16(Op::LOCAL_SET, wr_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, wr_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &weakref_slot(TARGET_KEY),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, target_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, wr_slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &weakref_slot(EPOCH_KEY),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, epoch_slot, line);

    emit_current_gc_epoch(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, current_epoch_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, epoch_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, current_epoch_slot, line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    chunk.emit_else(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_end(line);
}

/// `wr.Target = v` / `wr.SetTarget(v)`.
///
/// Stack on entry: `[wr, value]` ; on exit: `[]`.
pub fn emit_weakref_set_target(chunks: &mut [Chunk], current: usize, line: u32) {
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Stack,
        &weakref_slot(TARGET_KEY),
        ValueSource::Stack,
        line,
    );
    chunks[current].emit_op(Op::DROP, line);
}

/// `wr.IsAlive` — the target is still present. See the module note on why this
/// is presence rather than reachability.
///
/// Stack on entry: `[wr]` ; on exit: `[bool]`.
pub fn emit_weakref_is_alive(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_weakref_target(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    // ⛔ A Bool VALUE, not the raw i32 — `emit_dyn_to_bool`-shaped results
    // render as `1` where the corpus wants `True`.
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}
