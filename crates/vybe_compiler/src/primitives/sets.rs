//! Cross-language Set primitives.
//!
//! The backing store is the ECMA Set surface (`ecma:set.*`), because that is
//! already the portable substrate in `platforms/ecma`: unique values,
//! insertion-ordered iteration, and native set algebra. Language adapters layer
//! their quirks above this module instead of exposing `ecma:set` directly.

use crate::primitives::class_slots;
pub use vybe_ast::{
    SetAlgebraArity, SetMembership, SetMissingDelete, SetMutationResult, SetSemantics,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Sidecar property on a snapshot-keyed set: an `ecma:map` from the element's
/// INSERTION-TIME structural render to the element itself. The backing store
/// stays the ECMA Set (order, size, iteration are native); this map is what
/// makes membership follow the JDK `hashCode`/`equals` contract
/// ([`SetMembership::SnapshotKey`]): structurally equal values collide, and an
/// element mutated after insertion no longer answers to its current render.
pub const SNAPSHOT_KEYS_KEY: &str = "__snapshot_keys";

fn call(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("ecma:set", name);
    chunks[current].emit_call(idx, argc, line);
}

fn call_chunk(chunk: &mut Chunk, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import("ecma:set", name);
    chunk.emit_call(idx, argc, line);
}

fn stash_args(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(argc as u16);
    for offset in (0..argc as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }
    base
}

/// Stack: `[] -> [set]`.
pub fn emit_new(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "new", 0, line);
}

/// Stack: `[iterable] -> [set]`.
pub fn emit_from_iterable(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "fromIterable", 1, line);
}

/// Stack: `[elem0, elem1, ...] -> [set]`.
pub fn emit_literal(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        emit_new(chunks, current, line);
        return;
    }
    let base = chunks[current].alloc_scratch(argc as u16);
    crate::primitives::collections::emit_pack_n(chunks, current, argc as u16, base, line);
    emit_from_iterable(chunks, current, line);
}

/// Stack: `[set, value] -> [bool]`.
///
/// ECMA `Set.prototype.add` returns the receiver; Kotlin/.NET-style APIs need
/// "changed?" instead. Probe first, then add only when absent.
pub fn emit_add_changed(chunks: &mut [Chunk], current: usize, line: u32) {
    let set_slot = chunks[current].alloc_scratch(1);
    let value_slot = chunks[current].alloc_scratch(1);
    let existed_slot = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_has(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, existed_slot, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, existed_slot, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    emit_add(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

/// Stack: `[set, value] -> [set]`.
pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "add", 2, line);
}

pub fn emit_add_chunk(chunk: &mut Chunk, line: u32) {
    call_chunk(chunk, "add", 2, line);
}

/// Stack: `[set, value] -> [null]`.
pub fn emit_add_void(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_add(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

// ── SnapshotKey membership ([`SetMembership::SnapshotKey`]) ────────────────
//
// The JDK contract over the same ECMA-Set backing. Identity is the element's
// structural render (`ecma:json.stringify`) taken at INSERTION and kept in the
// [`SNAPSHOT_KEYS_KEY`] sidecar map; every membership question consults the
// sidecar, never the native `has` (which is SameValueZero).

fn call_host(chunks: &mut [Chunk], current: usize, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

/// Render the membership key for the value on the stack — the JDK
/// `hashCode`/`equals` proxy:
///
/// * primitives and null are their own key (an `ecma:map` keys primitives by
///   value, which IS their equality);
/// * an object whose class BINDS the `Eq` protocol slot (a Kotlin data class,
///   a record) keys by its structural render at this moment — so equivalent
///   values collide and a later mutation orphans the entry;
/// * a plain object keys by ITSELF — map keys hold objects by reference,
///   which is exactly Java's default identity `equals`.
///
/// The discriminator is the published Eq SLOT KEY (`protocol_slot_key` — an
/// unspellable role key the class emitter binds alongside the method), never
/// a method name: a user method merely spelled `equals` cannot set it
/// (flexclassplan §1e).
///
/// Stack: `[value] -> [key]`.
fn emit_snapshot_render(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_host(chunks, current, "ecma:value", "typeof", 1, line);
    chunks[current].emit_string_const("object", line);
    crate::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    // A Set ELEMENT keys by its sorted values — `AbstractSet.equals` is
    // order-independent, so `setOf(1,2)` and `setOf(2,1)` must collide.
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_host(chunks, current, "ecma:object", "toStringTag", 1, line);
    chunks[current].emit_string_const("[object Set]", line);
    call_host(chunks, current, "wasm:js-string", "equals", 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_values_array(chunks, current, line);
    crate::primitives::collections::emit_sorted(chunks, current, line);
    call_host(chunks, current, "ecma:json", "stringify", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_string_const(
        vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Eq).as_str(),
        line,
    );
    // `getMethodForCall` performs the FULL method walk (own props, prototype
    // chain, type-registry vtable) — bound slot methods live in the vtable
    // for typed instances, where a plain property read cannot see them.
    call_host(chunks, current, "ecma:value", "getMethodForCall", 2, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_host(chunks, current, "ecma:json", "stringify", 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// Load the sidecar map of the set in `set_slot` into `out_slot`, creating and
/// attaching it when absent (a set adopted from a plain construction).
fn emit_snapshot_keys_map(
    chunks: &mut [Chunk],
    current: usize,
    set_slot: u16,
    out_slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_string_const(SNAPSHOT_KEYS_KEY, line);
    call_host(chunks, current, "ecma:object", "get", 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    call_host(chunks, current, "ecma:map", "new", 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_string_const(SNAPSHOT_KEYS_KEY, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    call_host(chunks, current, "ecma:object", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

/// Stack: `[set, value] -> [changed-bool]`.
pub fn emit_add_snapshot(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);
    emit_snapshot_keys_map(chunks, current, set_slot, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_snapshot_render(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_host(chunks, current, "ecma:map", "has", 2, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call(chunks, current, "add", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    call_host(chunks, current, "ecma:map", "set", 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
}

/// Stack: `[set, value] -> [bool]`.
pub fn emit_has_snapshot(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);
    emit_snapshot_keys_map(chunks, current, set_slot, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_snapshot_render(chunks, current, line);
    call_host(chunks, current, "ecma:map", "has", 2, line);
}

/// Stack: `[set, value] -> [changed-bool]`.
pub fn emit_delete_snapshot(chunks: &mut [Chunk], current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    let set_slot = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);
    emit_snapshot_keys_map(chunks, current, set_slot, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_snapshot_render(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_host(chunks, current, "ecma:map", "has", 2, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_host(chunks, current, "ecma:map", "get", 2, line);
    call(chunks, current, "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    call_host(chunks, current, "ecma:map", "delete", 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
}

/// Stack: `[set] -> [set]` — native clear plus the sidecar's.
pub fn emit_clear_snapshot(chunks: &mut [Chunk], current: usize, line: u32) {
    let set_slot = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
    call(chunks, current, "clear", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    emit_snapshot_keys_map(chunks, current, set_slot, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    call_host(chunks, current, "ecma:map", "clear", 1, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
}

/// Stack: `[values-array] -> [set]` — a snapshot-keyed set of the array's
/// elements, deduplicated by their insertion-time render.
pub fn emit_from_iterable_snapshot(chunks: &mut [Chunk], current: usize, line: u32) {
    let values = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let index = chunks[current].alloc_scratch(1);
    let len = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, values, line);
    emit_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, values, line);
    crate::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    let block = chunks[current].emit_block(line);
    let (loop_id, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    crate::primitives::ops::emit_dyn_lt(&mut chunks[current], line);
    crate::primitives::ops::emit_dyn_not(&mut chunks[current], line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, values, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    crate::primitives::collections::emit_get(chunks, current, line);
    emit_add_snapshot(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, index, line);
    chunks[current].emit_i32_const(1, line);
    crate::primitives::ops::emit_dyn_add(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, index, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_id);
    chunks[current].emit_end(line);
    chunks[current].patch_block(block);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Stack: `[set, value] -> [bool]` — membership under the declared identity.
pub fn emit_has_mode(chunks: &mut [Chunk], current: usize, semantics: SetSemantics, line: u32) {
    match semantics.membership {
        SetMembership::SnapshotKey => emit_has_snapshot(chunks, current, line),
        SetMembership::SameValueZero => emit_has(chunks, current, line),
    }
}

/// Stack: `[set, value] -> [mode result]`.
pub fn emit_add_mode(chunks: &mut [Chunk], current: usize, semantics: SetSemantics, line: u32) {
    if semantics.membership == SetMembership::SnapshotKey {
        match semantics.mutation_result {
            SetMutationResult::ChangedBool => emit_add_snapshot(chunks, current, line),
            SetMutationResult::Receiver => {
                let value = chunks[current].alloc_scratch(1);
                let set_slot = chunks[current].alloc_scratch(1);
                chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
                chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
                emit_add_snapshot(chunks, current, line);
                chunks[current].emit_op(Op::DROP, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
            }
            SetMutationResult::Void => {
                emit_add_snapshot(chunks, current, line);
                chunks[current].emit_op(Op::DROP, line);
                chunks[current]
                    .emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            }
        }
        return;
    }
    match semantics.mutation_result {
        SetMutationResult::Receiver => emit_add(chunks, current, line),
        SetMutationResult::ChangedBool => emit_add_changed(chunks, current, line),
        SetMutationResult::Void => emit_add_void(chunks, current, line),
    }
}

/// Stack: `[set, value] -> [bool]`.
pub fn emit_delete(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "delete", 2, line);
}

pub fn emit_delete_chunk(chunk: &mut Chunk, line: u32) {
    call_chunk(chunk, "delete", 2, line);
}

/// Stack: `[set, value] -> [null]`.
pub fn emit_delete_void(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_delete(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[set, value] -> [null]`; throws `KeyError` if the value is missing.
pub fn emit_delete_or_key_error(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_delete(chunks, current, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    class_slots::emit_class_alloc(&mut chunks[current], line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    crate::primitives::errors::emit_exception_new_finalize(&mut chunks[current], "KeyError", line);
    crate::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[set, value] -> [mode result]`.
pub fn emit_delete_mode(chunks: &mut [Chunk], current: usize, semantics: SetSemantics, line: u32) {
    if semantics.membership == SetMembership::SnapshotKey {
        if semantics.missing_delete == SetMissingDelete::ThrowKeyError {
            emit_delete_snapshot(chunks, current, line);
            crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if(line);
            class_slots::emit_class_alloc(&mut chunks[current], line);
            chunks[current].emit_dup(line);
            chunks[current].emit_string_const("", line);
            crate::primitives::errors::emit_exception_new_finalize(
                &mut chunks[current],
                "KeyError",
                line,
            );
            crate::primitives::errors::emit_throw(&mut chunks[current], line);
            chunks[current].emit_end(line);
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            return;
        }
        match semantics.delete_result {
            SetMutationResult::ChangedBool => emit_delete_snapshot(chunks, current, line),
            SetMutationResult::Receiver => {
                let value = chunks[current].alloc_scratch(1);
                let set_slot = chunks[current].alloc_scratch(1);
                chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
                chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
                emit_delete_snapshot(chunks, current, line);
                chunks[current].emit_op(Op::DROP, line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
            }
            SetMutationResult::Void => {
                emit_delete_snapshot(chunks, current, line);
                chunks[current].emit_op(Op::DROP, line);
                chunks[current]
                    .emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            }
        }
        return;
    }
    if semantics.missing_delete == SetMissingDelete::ThrowKeyError {
        emit_delete_or_key_error(chunks, current, line);
        return;
    }
    match semantics.delete_result {
        SetMutationResult::Receiver => {
            let set_slot = chunks[current].alloc_scratch(1);
            let value_slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, value_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, set_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
            emit_delete(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, set_slot, line);
        }
        SetMutationResult::ChangedBool => emit_delete(chunks, current, line),
        SetMutationResult::Void => emit_delete_void(chunks, current, line),
    }
}

/// Stack: `[set, value] -> [mode delete result]`; missing is always ignored.
pub fn emit_discard_mode(
    chunks: &mut [Chunk],
    current: usize,
    mut semantics: SetSemantics,
    line: u32,
) {
    semantics.missing_delete = SetMissingDelete::Ignore;
    emit_delete_mode(chunks, current, semantics, line);
}

/// Stack: `[set, value] -> [bool]`.
pub fn emit_has(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "has", 2, line);
}

pub fn emit_has_chunk(chunk: &mut Chunk, line: u32) {
    call_chunk(chunk, "has", 2, line);
}

/// Stack: `[set] -> [i32]`.
pub fn emit_size(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "size", 1, line);
}

pub fn emit_size_chunk(chunk: &mut Chunk, line: u32) {
    call_chunk(chunk, "size", 1, line);
}

/// Stack: `[set] -> [null]`.
pub fn emit_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "clear", 1, line);
}

/// Stack: `[set] -> [null]`.
pub fn emit_clear_void(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_clear(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[set] -> [mode result]`.
pub fn emit_clear_mode(chunks: &mut [Chunk], current: usize, semantics: SetSemantics, line: u32) {
    if semantics.membership == SetMembership::SnapshotKey {
        emit_clear_snapshot(chunks, current, line);
        if semantics.mutation_result == SetMutationResult::Void {
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
        return;
    }
    match semantics.mutation_result {
        SetMutationResult::Void => emit_clear_void(chunks, current, line),
        _ => emit_clear(chunks, current, line),
    }
}

/// Stack: `[left, right] -> [set]`.
pub fn emit_union(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "union", 2, line);
}

pub fn emit_union_chunk(chunk: &mut Chunk, line: u32) {
    call_chunk(chunk, "union", 2, line);
}

/// Stack: `[set0, set1, ...] -> [set]`.
pub fn emit_union_variadic(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        emit_new(chunks, current, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let out = chunks[current].alloc_scratch(1);
    emit_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    for offset in 0..argc as u16 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + offset, line);
        emit_union_with(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Stack: binary mode `[left, right]`, variadic mode `[set0, set1, ...]`.
pub fn emit_union_mode(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    semantics: SetSemantics,
    line: u32,
) {
    match semantics.algebra_arity {
        SetAlgebraArity::Binary => emit_union(chunks, current, line),
        SetAlgebraArity::Variadic => emit_union_variadic(chunks, current, argc, line),
    }
}

/// Stack: `[target, values] -> [target]`.
pub fn emit_union_with(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "unionWith", 2, line);
}

/// Stack: `[target, values] -> [null]`.
pub fn emit_union_with_void(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_union_with(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[target, values] -> [mode mutation result]`.
pub fn emit_union_with_mode(
    chunks: &mut [Chunk],
    current: usize,
    semantics: SetSemantics,
    line: u32,
) {
    match semantics.mutation_result {
        SetMutationResult::Receiver => emit_union_with(chunks, current, line),
        SetMutationResult::ChangedBool => {
            emit_union_with(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_bool_const(true, line);
        }
        SetMutationResult::Void => emit_union_with_void(chunks, current, line),
    }
}

/// Stack: `[left, right] -> [set]`.
pub fn emit_intersection(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "intersection", 2, line);
}

pub fn emit_intersection_chunk(chunk: &mut Chunk, line: u32) {
    call_chunk(chunk, "intersection", 2, line);
}

/// Stack: `[set0, set1, ...] -> [set]`.
pub fn emit_intersection_variadic(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        emit_new(chunks, current, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let out = chunks[current].alloc_scratch(1);
    emit_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    emit_union(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    for offset in 1..argc as u16 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + offset, line);
        emit_intersection(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Stack: binary mode `[left, right]`, variadic mode `[set0, set1, ...]`.
pub fn emit_intersection_mode(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    semantics: SetSemantics,
    line: u32,
) {
    match semantics.algebra_arity {
        SetAlgebraArity::Binary => emit_intersection(chunks, current, line),
        SetAlgebraArity::Variadic => emit_intersection_variadic(chunks, current, argc, line),
    }
}

/// Stack: `[target, values] -> [target]`.
pub fn emit_intersect_with(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "intersectWith", 2, line);
}

/// Stack: `[target, values] -> [null]`.
pub fn emit_intersect_with_void(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_intersect_with(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[target, values] -> [mode mutation result]`.
pub fn emit_intersect_with_mode(
    chunks: &mut [Chunk],
    current: usize,
    semantics: SetSemantics,
    line: u32,
) {
    match semantics.mutation_result {
        SetMutationResult::Receiver => emit_intersect_with(chunks, current, line),
        SetMutationResult::ChangedBool => {
            emit_intersect_with(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_bool_const(true, line);
        }
        SetMutationResult::Void => emit_intersect_with_void(chunks, current, line),
    }
}

/// Stack: `[left, right] -> [set]`.
pub fn emit_difference(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "difference", 2, line);
}

pub fn emit_difference_chunk(chunk: &mut Chunk, line: u32) {
    call_chunk(chunk, "difference", 2, line);
}

/// Stack: `[set0, set1, ...] -> [set]`.
pub fn emit_difference_variadic(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        emit_new(chunks, current, line);
        return;
    }
    let base = stash_args(chunks, current, argc, line);
    let out = chunks[current].alloc_scratch(1);
    emit_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    emit_union(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    for offset in 1..argc as u16 {
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, base + offset, line);
        emit_difference(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// Stack: binary mode `[left, right]`, variadic mode `[set0, set1, ...]`.
pub fn emit_difference_mode(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    semantics: SetSemantics,
    line: u32,
) {
    match semantics.algebra_arity {
        SetAlgebraArity::Binary => emit_difference(chunks, current, line),
        SetAlgebraArity::Variadic => emit_difference_variadic(chunks, current, argc, line),
    }
}

/// Stack: `[target, values] -> [target]`.
pub fn emit_except_with(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "exceptWith", 2, line);
}

/// Stack: `[target, values] -> [null]`.
pub fn emit_except_with_void(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_except_with(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[target, values] -> [mode mutation result]`.
pub fn emit_except_with_mode(
    chunks: &mut [Chunk],
    current: usize,
    semantics: SetSemantics,
    line: u32,
) {
    match semantics.mutation_result {
        SetMutationResult::Receiver => emit_except_with(chunks, current, line),
        SetMutationResult::ChangedBool => {
            emit_except_with(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_bool_const(true, line);
        }
        SetMutationResult::Void => emit_except_with_void(chunks, current, line),
    }
}

/// Stack: `[left, right] -> [set]`.
pub fn emit_symmetric_difference(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "symmetricDifference", 2, line);
}

/// Stack: `[target, values] -> [target]`.
pub fn emit_symmetric_except_with(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "symmetricExceptWith", 2, line);
}

/// Stack: `[target, values] -> [null]`.
pub fn emit_symmetric_except_with_void(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_symmetric_except_with(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[target, values] -> [mode mutation result]`.
pub fn emit_symmetric_except_with_mode(
    chunks: &mut [Chunk],
    current: usize,
    semantics: SetSemantics,
    line: u32,
) {
    match semantics.mutation_result {
        SetMutationResult::Receiver => emit_symmetric_except_with(chunks, current, line),
        SetMutationResult::ChangedBool => {
            emit_symmetric_except_with(chunks, current, line);
            chunks[current].emit_op(Op::DROP, line);
            chunks[current].emit_bool_const(true, line);
        }
        SetMutationResult::Void => emit_symmetric_except_with_void(chunks, current, line),
    }
}

/// Stack: `[left, right] -> [bool]`.
pub fn emit_is_subset_of(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "isSubsetOf", 2, line);
}

pub fn emit_is_subset_of_chunk(chunk: &mut Chunk, line: u32) {
    call_chunk(chunk, "isSubsetOf", 2, line);
}

/// Stack: `[left, right] -> [bool]`.
pub fn emit_is_subset_of_bool(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_is_subset_of(chunks, current, line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Stack: `[left, right] -> [predicate result]`.
pub fn emit_subset_mode(chunks: &mut [Chunk], current: usize, semantics: SetSemantics, line: u32) {
    if semantics.predicate_bool_object {
        emit_is_subset_of_bool(chunks, current, line);
    } else {
        emit_is_subset_of(chunks, current, line);
    }
}

/// Stack: `[left, right] -> [bool]`.
pub fn emit_is_superset_of(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "isSupersetOf", 2, line);
}

pub fn emit_is_superset_of_chunk(chunk: &mut Chunk, line: u32) {
    call_chunk(chunk, "isSupersetOf", 2, line);
}

/// Stack: `[left, right] -> [bool]`.
pub fn emit_is_superset_of_bool(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_is_superset_of(chunks, current, line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Stack: `[left, right] -> [predicate result]`.
pub fn emit_superset_mode(
    chunks: &mut [Chunk],
    current: usize,
    semantics: SetSemantics,
    line: u32,
) {
    if semantics.predicate_bool_object {
        emit_is_superset_of_bool(chunks, current, line);
    } else {
        emit_is_superset_of(chunks, current, line);
    }
}

/// Stack: `[left, right] -> [bool]`.
pub fn emit_is_disjoint_from(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "isDisjointFrom", 2, line);
}

/// Stack: `[left, right] -> [bool]`.
pub fn emit_is_disjoint_from_bool(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_is_disjoint_from(chunks, current, line);
    crate::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}

/// Stack: `[left, right] -> [predicate result]`.
pub fn emit_disjoint_mode(
    chunks: &mut [Chunk],
    current: usize,
    semantics: SetSemantics,
    line: u32,
) {
    if semantics.predicate_bool_object {
        emit_is_disjoint_from_bool(chunks, current, line);
    } else {
        emit_is_disjoint_from(chunks, current, line);
    }
}

/// Stack: `[set] -> [array]`.
pub fn emit_values_array(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::primitives::collections::emit_iter_for_of(chunks, current, line);
}


// ── Linkable chunk builders ──────────────────────────────────────────────────
//
// The `emit_*_mode` functions above already carry every convention a language
// needs — which value comes back from a mutation, whether a missing delete
// raises, whether a predicate yields a bool OBJECT — driven by
// `vybe_ast::SetSemantics`. These builders only wrap them as standalone chunks
// so compiled code can reach one through a `__vybe_*` global.
//
// The `pascal_` prefix records which language first needed a linkable chunk,
// not anything Pascal-specific: set union is set union. A second language
// wanting the same should reuse these rather than clone the pattern under its
// own prefix.

pub fn build_pascal_set_include(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_include");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    emit_add_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

pub fn build_pascal_set_exclude(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_exclude");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    emit_delete_chunk(&mut c, 0);
    c.emit_op(Op::DROP, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

pub fn build_pascal_set_union(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_union");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    emit_union_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

pub fn build_pascal_set_intersection(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_intersection");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    emit_intersection_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

pub fn build_pascal_set_difference(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_difference");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    emit_difference_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

pub fn build_pascal_set_contains(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_pascal_set_contains");
    c.arity = 2;
    c.local_count = 2;
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    emit_has_chunk(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    c
}
