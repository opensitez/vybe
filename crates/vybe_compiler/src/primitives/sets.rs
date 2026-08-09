//! Cross-language Set primitives.
//!
//! The backing store is the ECMA Set surface (`ecma:set.*`), because that is
//! already the portable substrate in `platforms/ecma`: unique values,
//! insertion-ordered iteration, and native set algebra. Language adapters layer
//! their quirks above this module instead of exposing `ecma:set` directly.

use vybe_ast::{SetAlgebraArity, SetMissingDelete, SetMutationResult, SetSemantics};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

fn call(chunks: &mut [Chunk], current: usize, name: &str, argc: u8, line: u32) {
    let idx = chunks[current].add_import("ecma:set", name);
    chunks[current].emit_call(idx, argc, line);
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

/// Stack: `[set, value] -> [null]`.
pub fn emit_add_void(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_add(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[set, value] -> [mode result]`.
pub fn emit_add_mode(chunks: &mut [Chunk], current: usize, semantics: SetSemantics, line: u32) {
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
    chunks[current].emit_struct_new(0, 0, line);
    chunks[current].emit_dup(line);
    chunks[current].emit_string_const("", line);
    crate::primitives::errors::emit_exception_new_finalize(&mut chunks[current], "KeyError", line);
    crate::primitives::errors::emit_throw(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// Stack: `[set, value] -> [mode result]`.
pub fn emit_delete_mode(chunks: &mut [Chunk], current: usize, semantics: SetSemantics, line: u32) {
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

/// Stack: `[set] -> [i32]`.
pub fn emit_size(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "size", 1, line);
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
    match semantics.mutation_result {
        SetMutationResult::Void => emit_clear_void(chunks, current, line),
        _ => emit_clear(chunks, current, line),
    }
}

/// Stack: `[left, right] -> [set]`.
pub fn emit_union(chunks: &mut [Chunk], current: usize, line: u32) {
    call(chunks, current, "union", 2, line);
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
