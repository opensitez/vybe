//! Python `weakref` and `gc`.
//!
//! ⚠ Both modules sit on top of a documented VM limitation: `ecma:weakref` is
//! a STRONG-reference stand-in (`platforms/ecma/src/builtin_types.rs` says so
//! — the WASM GC MVP exposes no weak references), and the VM runs no
//! collector a program can trigger. So a `weakref.ref` here never goes dead
//! and `gc.collect()` never frees anything. The SHAPES are real — `r()`
//! derefs, `finalize` runs its callback, `gc.collect()` returns a count — and
//! they become truthful the moment the VM grows real weak refs, because both
//! read through the same `ecma:weakref` surface a real implementation would.

use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use vybe_compiler::primitives::collections;

use super::adapter_util::{lget, new_tagged, set_call_slot, stash_exact, struct_get, struct_set};

/// `weakref.ref(obj[, callback])` — a callable that derefs to the referent.
pub fn emit_ref(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let call_idx = build_deref_helper(chunks, line);
    let base = stash_exact(chunks, current, argc, 2, line);
    let chunk = &mut chunks[current];

    new_tagged(chunk, "weakref", &[("__callback", base + 1)], line);
    chunk.emit_dup(line);
    lget(chunk, base, line);
    let idx = chunk.add_import("ecma:weakref", "new");
    chunk.emit_call(idx, 1, line);
    struct_set(chunk, "__weak", line);
    set_call_slot(chunk, call_idx, line);

    // Record the ref ON the referent, so `getweakrefs`/`getweakrefcount`
    // answer from real bookkeeping rather than a guess. CPython keeps the
    // same list; the difference is only that ours never loses entries,
    // because nothing collects.
    let this_ref = chunks[current].alloc_scratch(1);
    super::adapter_util::lset(&mut chunks[current], this_ref, line);
    let list = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], base, line);
    chunks[current].emit_string_const(WEAKREFS_KEY, line);
    let get = chunks[current].add_import("ecma:object", "get");
    chunks[current].emit_call(get, 2, line);
    super::adapter_util::lset(&mut chunks[current], list, line);
    lget(&mut chunks[current], list, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    collections::emit_array_new(chunks, current, 0, line);
    super::adapter_util::lset(&mut chunks[current], list, line);
    lget(&mut chunks[current], base, line);
    lget(&mut chunks[current], list, line);
    struct_set(&mut chunks[current], WEAKREFS_KEY, line);
    chunks[current].emit_end(line);
    lget(&mut chunks[current], list, line);
    lget(&mut chunks[current], this_ref, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    lget(&mut chunks[current], this_ref, line);
}

/// `(this) -> deref(this.__weak)`.
fn build_deref_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut helper = Chunk::new("__py_weakref_deref");
    helper.arity = 1;
    helper.local_count = helper.local_count.max(1);
    helper.emit_op_u16(Op::LOCAL_GET, 0, line);
    struct_get(&mut helper, "__weak", line);
    let idx = helper.add_import("ecma:weakref", "deref");
    helper.emit_call(idx, 1, line);
    helper.emit_op(Op::RETURN, line);
    chunks.push(helper);
    chunks.len() - 1
}

/// `weakref.proxy(obj)`.
///
/// A CPython proxy forwards every operation to a referent that can die. With
/// strong-only refs there is nothing to forward THROUGH: a proxy that can
/// never go dead is observationally the referent, so that is what this
/// returns rather than a wrapper that would only add a layer of divergence.
pub fn emit_proxy(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 2, line);
    lget(&mut chunks[current], base, line);
}

/// `weakref.getweakrefcount(obj)` — the length of the same list
/// `getweakrefs` returns.
pub fn emit_getweakrefcount(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_getweakrefs(chunks, current, argc, line);
    collections::emit_len(chunks, current, line);
}

/// `weakref.finalize(obj, func, *args)` — a callable that runs `func(*args)`.
///
/// CPython also runs it when the referent dies or at interpreter exit; with no
/// collector, calling it explicitly (which is a documented part of the API —
/// `f()` runs the callback and returns its result) is the path that works.
pub fn emit_finalize(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let call_idx = build_finalize_call_helper(chunks, line);
    let base = stash_exact(chunks, current, argc, argc.max(2) as u16, line);
    // Everything past `(obj, func)` is a bound argument for the callback.
    let extra = argc.saturating_sub(2) as u16;
    let args_slot = chunks[current].alloc_scratch(1);
    for offset in 0..extra {
        lget(&mut chunks[current], base + 2 + offset, line);
    }
    collections::emit_array_new(chunks, current, extra, line);
    super::adapter_util::lset(&mut chunks[current], args_slot, line);

    let chunk = &mut chunks[current];
    new_tagged(
        chunk,
        "finalize",
        &[
            ("__obj", base),
            ("__func", base + 1),
            ("__args", args_slot),
        ],
        line,
    );
    chunk.emit_dup(line);
    chunk.emit_bool_const(true, line);
    struct_set(chunk, "alive", line);
    set_call_slot(chunk, call_idx, line);
}

/// `(this) -> this.__func(*this.__args)`.
fn build_finalize_call_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let mut helper = Chunk::new("__py_finalize_call");
    helper.arity = 1;
    helper.local_count = helper.local_count.max(1);
    helper.emit_op_u16(Op::LOCAL_GET, 0, line);
    struct_get(&mut helper, "__func", line);
    // `apply(fn, thisArg, argsArray)` — a plain callback has no receiver.
    helper.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    helper.emit_op_u16(Op::LOCAL_GET, 0, line);
    struct_get(&mut helper, "__args", line);
    let apply = helper.add_import("ecma:function", "apply");
    helper.emit_call(apply, 3, line);
    helper.emit_op(Op::RETURN, line);
    chunks.push(helper);
    chunks.len() - 1
}

/// `gc.collect()` → 0 objects freed: the VM has no cycle collector to run.
pub fn emit_gc_collect(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_i32_const(0, line);
}

/// `gc.get_objects()` / `gc.get_referrers(x)` → an empty list. The heap is not
/// enumerable from bytecode; an empty list is the honest answer to "which
/// tracked objects can you name", not a placeholder for one.
pub fn emit_gc_empty_list(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    collections::emit_array_new(chunks, current, 0, line);
}

/// `gc.isenabled()` → True; `gc.enable()`/`gc.disable()` → None.
pub fn emit_gc_bool(chunks: &mut [Chunk], current: usize, argc: u8, value: bool, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_bool_const(value, line);
}

pub fn emit_gc_none(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    let _ = Value::Null;
}

/// `gc.get_count()` → `(0, 0, 0)` and `gc.get_threshold()` → `(700, 10, 10)`.
///
/// Three generations is not a detail a caller may vary: `get_count` and
/// `get_threshold` are documented to return a 3-tuple, and CPython's default
/// thresholds are what an untouched interpreter reports. With no collector the
/// counts are genuinely zero — nothing is pending collection.
pub fn emit_gc_triple(chunks: &mut [Chunk], current: usize, argc: u8, values: [f64; 3], line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    for value in values {
        chunks[current].emit_f64_const(value, line);
    }
    vybe_compiler::primitives::tuples::emit_tuple(chunks, current, 3, line);
}

/// `gc.get_stats()` → one record per generation, the shape CPython documents:
/// `{'collections': …, 'collected': …, 'uncollectable': …}`.
pub fn emit_gc_stats(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    for _ in 0..3 {
        vybe_compiler::primitives::dict::emit_new(chunks, current, line);
        for key in ["collections", "collected", "uncollectable"] {
            chunks[current].emit_f64_const(0.0, line);
            vybe_compiler::primitives::dict::emit_set_const_key(chunks, current, key, line);
        }
    }
    collections::emit_array_new(chunks, current, 3, line);
}

/// `gc.is_tracked(x)` — CPython tracks containers and never tracks an atomic
/// value, so this is "is it an object", which is exactly what `typeof` says.
pub fn emit_gc_is_tracked(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    lget(&mut chunks[current], base, line);
    vybe_compiler::primitives::reflection::emit_typeof(chunks, current, line);
    chunks[current].emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    let from_i32 = chunks[current].add_import("wasm:js-boolean", "fromI32");
    chunks[current].emit_call(from_i32, 1, line);
}

/// `gc.get_debug()` → 0: no debug flags are set, and none can be.
pub fn emit_gc_zero(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_f64_const(0.0, line);
}

/// The list of objects weakly referenced by `obj`, which `weakref.ref` keeps
/// on the referent itself.
const WEAKREFS_KEY: &str = "__weakrefs__";

/// `weakref.getweakrefs(obj)` → every `weakref.ref` created for `obj`.
pub fn emit_getweakrefs(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let base = stash_exact(chunks, current, argc, 1, line);
    let found = chunks[current].alloc_scratch(1);
    lget(&mut chunks[current], base, line);
    chunks[current].emit_string_const(WEAKREFS_KEY, line);
    let get = chunks[current].add_import("ecma:object", "get");
    chunks[current].emit_call(get, 2, line);
    super::adapter_util::lset(&mut chunks[current], found, line);
    lget(&mut chunks[current], found, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    lget(&mut chunks[current], found, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_end(line);
}
