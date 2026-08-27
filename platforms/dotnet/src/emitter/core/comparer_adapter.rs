//! `Comparer` / `EqualityComparer` / `StringComparer` — the comparison
//! behaviour, owned by the platform.
//!
//! These used to be a compile-time STRING marker
//! (`__dotnet_stringcomparer_ordinalignorecase`) that each .NET frontend then
//! pattern-matched and rewrote for itself: five sites in the C# walker, three
//! more in the VB walker, and nothing at all for PowerShell. The marker was
//! produced here but the semantics acting on it were not, so the same rule was
//! written twice and shipped once.
//!
//! The marker survives — it is a fine internal sentinel, and
//! `array_adapter`/`linq_adapter` already hand it to comparer-taking APIs. What
//! changes is that the sentinel is now the VALUE of a typed tree member, so
//! `.Compare(a, b)` resolves through the ordinary member walk in every
//! frontend instead of through a frontend's own rewrite.
//!
//! Case sensitivity is decided at RUN time by looking at the receiver, not at
//! compile time by looking at which spelling produced it. That is what lets one
//! emit serve `Ordinal` and `OrdinalIgnoreCase` — and it is also what lets a
//! comparer held in a variable behave correctly, which a compile-time branch on
//! the literal never could.

use vybe_compiler::primitives::{object, ops};
use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

pub const ORDINAL: &str = "__dotnet_stringcomparer_ordinal";
pub const ORDINAL_IGNORE_CASE: &str = "__dotnet_stringcomparer_ordinalignorecase";
pub const COMPARER_DEFAULT: &str = "__dotnet_comparer_default";
pub const EQUALITY_COMPARER_DEFAULT: &str = "__dotnet_equalitycomparer_default";

/// `Comparer.Default` / `StringComparer.Ordinal` / … — the sentinel value.
pub fn emit_marker(chunks: &mut [Chunk], current: usize, marker: &str, line: u32) {
    chunks[current].emit_string_const(marker, line);
}

/// Push `slot` folded to lower case when `recv` is the ignore-case comparer,
/// and unfolded otherwise.
fn emit_operand_folded_for(
    chunks: &mut [Chunk],
    current: usize,
    recv: u16,
    slot: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_string_const(ORDINAL_IGNORE_CASE, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_if(line);

    // `toLowerCase` traps on a non-string, and `Comparer<int>.Default` reaches
    // the same emit — guard on the operand's type, not on the comparer alone.
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_string = chunks[current].add_import("wasm:js-string", "test");
    chunks[current].emit_call(is_string, 1, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    let lower = chunks[current].add_import("ecma:string", "toLowerCase");
    chunks[current].emit_call(lower, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_end(line);

    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_end(line);
}

/// Stack `[recv, a, b]` → the two operands, case-folded per the receiver.
fn stash_and_fold(chunks: &mut [Chunk], current: usize, line: u32) -> u16 {
    let base = chunks[current].alloc_scratch(3);
    let (recv, a, b) = (base, base + 1, base + 2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, b, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, a, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    emit_operand_folded_for(chunks, current, recv, a, line);
    emit_operand_folded_for(chunks, current, recv, b, line);
    base
}

/// `comparer.Compare(a, b)` → -1 / 0 / 1.
///
/// Routed to the SHARED `object.compare`, which the JVM tree already reaches
/// under the same name. A .NET-only spaceship here would be a second answer to
/// a question that already has one.
pub fn emit_compare(chunks: &mut [Chunk], current: usize, line: u32) {
    stash_and_fold(chunks, current, line);
    // `object.compare` takes [a, b, comparator]; the comparator slot is the
    // user-supplied one, and these built-ins are exactly the absence of it.
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    object::emit_compare(&mut chunks[current], line);
}

/// `comparer.Equals(a, b)`.
pub fn emit_equals(chunks: &mut [Chunk], current: usize, line: u32) {
    stash_and_fold(chunks, current, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
}

/// `comparer.GetHashCode(x)` — stack `[recv, x]`.
///
/// Case-folded first for the ignore-case comparer, because .NET's contract is
/// that equal values hash equally: `Equals("A", "a")` is true there, so their
/// hashes may not differ.
pub fn emit_get_hash_code(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(2);
    let (recv, value) = (base, base + 1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    emit_operand_folded_for(chunks, current, recv, value, line);
    object::emit_hash_code(&mut chunks[current], line);
}

/// `Array.Sort(arr, comparer)` — in-place, stack `[arr, comparer]` → `[null]`.
///
/// .NET declares `Sort` at arity 1 in this platform and nowhere at arity 2, so
/// BOTH the C# and VB walkers carried their own two-argument rewrite. This is
/// that overload, once.
///
/// An insertion sort mirroring `collections::emit_sort_func`, with the
/// comparator call replaced by the folded comparison above — `emit_sort_func`
/// invokes a comparator VALUE, and these built-ins are precisely the case where
/// there is no such value to invoke.
pub fn emit_array_sort_with_comparer(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].alloc_scratch(7);
    let (cmp, arr, len, i, j, tmp, probe) =
        (base, base + 1, base + 2, base + 3, base + 4, base + 5, base + 6);

    chunks[current].emit_op_u16(Op::LOCAL_SET, cmp, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, arr, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    vybe_compiler::primitives::collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let outer_block = chunks[current].emit_block(line);
    let (outer_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);

    let inner_block = chunks[current].emit_block(line);
    let (inner_loop, _) = chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    // arr[j] and arr[j-1], each folded per the comparer, then compared.
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tmp, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, probe, line);

    emit_operand_folded_for(chunks, current, cmp, tmp, line);
    emit_operand_folded_for(chunks, current, cmp, probe, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    object::emit_compare(&mut chunks[current], line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_lt(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(1, line);

    // swap arr[j] and arr[j-1] — `tmp` already holds arr[j].
    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, probe, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, arr, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp, line);
    vybe_compiler::primitives::collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(inner_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(inner_block);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(outer_loop);
    chunks[current].emit_end(line);
    chunks[current].patch_block(outer_block);

    // .NET's `Array.Sort` is void.
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}
