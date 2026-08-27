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

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

const TYPE_KEY: &str = "__type";
const TARGET_KEY: &str = "__wr_target";
const TYPE_NAME: &str = "WeakReference";

fn key(chunk: &mut Chunk, name: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(name)))
}

fn struct_set(chunk: &mut Chunk, name: &str, line: u32) {
    let k = key(chunk, name);
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
}

fn struct_get(chunk: &mut Chunk, name: &str, line: u32) {
    let k = key(chunk, name);
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
}

/// `New WeakReference(obj)`, `New WeakReference(obj, trackResurrection)` and
/// `New WeakReference(Of T)(obj)`.
///
/// Stack on entry: `[obj]` or `[obj, trackResurrection]` ; on exit: `[wr]`.
pub fn emit_weakref_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let base = chunk.alloc_scratch(2);
    let (obj_slot, target_slot) = (base, base + 1);

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

    chunk.emit_struct_new(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_string_const(TYPE_NAME, line);
    struct_set(chunk, TYPE_KEY, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target_slot, line);
    struct_set(chunk, TARGET_KEY, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `wr.Target` — and the zero-arity core the `TryGetTarget` out-param desugar
/// calls. One emitter for both spellings on purpose.
///
/// Stack on entry: `[wr]` ; on exit: `[target]`.
pub fn emit_weakref_target(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_get(&mut chunks[current], TARGET_KEY, line);
}

/// `wr.Target = v` / `wr.SetTarget(v)`.
///
/// Stack on entry: `[wr, value]` ; on exit: `[]`.
pub fn emit_weakref_set_target(chunks: &mut [Chunk], current: usize, line: u32) {
    struct_set(&mut chunks[current], TARGET_KEY, line);
}

/// `wr.IsAlive` — the target is still present. See the module note on why this
/// is presence rather than reachability.
///
/// Stack on entry: `[wr]` ; on exit: `[bool]`.
pub fn emit_weakref_is_alive(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    struct_get(chunk, TARGET_KEY, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    // ⛔ A Bool VALUE, not the raw i32 — `emit_dyn_to_bool`-shaped results
    // render as `1` where the corpus wants `True`.
    vybe_compiler::primitives::ops::emit_i32_to_bool(chunk, line);
}
