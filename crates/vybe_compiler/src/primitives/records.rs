//! Record semantics — the behaviour a language DECLARES rather than implements.
//!
//! A language states `ValueSemantics { storage, equality, layout, variant }` on
//! its declaration and this module owns what that means. See
//! `recordprimitiveplan.md`.
//!
//! **Why it cannot live in the walkers.** A Pascal record can be passed to C#
//! or PHP. The receiver sees a runtime object and cannot know a `record`
//! declared it, so a per-language pass can never reach a foreign value —
//! Pascal's `lower_struct_copy_assignments` keys on Pascal's own declarations,
//! and a COBOL group simply is not in that map, so the assignment aliases with
//! no diagnostic. The semantics have to travel on the INSTANCE.
//!
//! The channel already exists: `__value_eq`, stamped at construction by
//! `classes::emit_value_equality_stamp` when the declaration says
//! `ValueEquality::Structural`.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;


use crate::primitives::{collections, loops, ops};

/// Field-wise equality for two objects held in locals. Stack: `[] → [bool]`.
///
/// Generalised from Dart's private `emit_dart_fields_equal`, which was the only
/// implementation — the same reason `tuples.rs` exists for the structural
/// flavour. Every language whose declaration says `Structural` reaches this,
/// whoever allocated the object, which is what makes a record crossing a
/// language boundary keep its equality.
///
/// Fields are compared with the PRIMITIVE equality, not recursively: this
/// emitter inlines its body, so a nested full comparison would expand forever
/// at compile time. A value type's fields are scalars — numbers, strings,
/// bools, enum spellings — which is what the primitive form handles.
pub fn emit_value_fields_equal(
    chunks: &mut [Chunk],
    current: usize,
    left_slot: u16,
    right_slot: u16,
    line: u32,
) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(5);
    let (left_keys, right_keys, idx_slot, key_slot, result_slot) =
        (base, base + 1, base + 2, base + 3, base + 4);

    let keys_of = |chunks: &mut [Chunk], current: usize, src: u16, dst: u16| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
        let idx = chunks[current].add_import("ecma:object", "keys");
        chunks[current].emit_call(idx, 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, dst, line);
    };
    keys_of(chunks, current, left_slot, left_keys);
    keys_of(chunks, current, right_slot, right_keys);

    // Start equal, then let any difference falsify it.
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);

    // A differing field COUNT is already a mismatch.
    chunks[current].emit_op_u16(Op::LOCAL_GET, left_keys, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, right_keys, line);
    collections::emit_len(chunks, current, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);

    let state = loops::emit_for_in_start(chunks, current, left_keys, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_slot, line);
    let get_field = |chunks: &mut [Chunk], current: usize, obj: u16| {
        chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_slot, line);
        let idx = chunks[current].add_import("ecma:object", "get");
        chunks[current].emit_call(idx, 2, line);
    };
    get_field(chunks, current, left_slot);
    get_field(chunks, current, right_slot);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result_slot, line);
    chunks[current].emit_end(line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result_slot, line);
}

/// Is the value in `slot` an instance whose declaration said its equality is
/// by VALUE? Stack: `[] → [i32]`, reading the `__value_eq` instance stamp.
pub fn emit_is_value_eq(chunk: &mut Chunk, slot: u16, line: u32) {
    emit_reads_stamp(chunk, slot, "__value_eq", line)
}

/// Is the value in `slot` an instance whose declaration said assignment COPIES?
/// Stack: `[] → [i32]`, reading the `__value_copy` instance stamp.
pub fn emit_is_value_copy(chunk: &mut Chunk, slot: u16, line: u32) {
    emit_reads_stamp(chunk, slot, "__value_copy", line)
}

/// A stamp is either ABSENT or the literal `true` — never a number, a string,
/// or anything else. So a null test answers it.
///
/// This was `ops::emit_dyn_to_bool`, the general dynamic-truthiness conversion,
/// which expands into `js-boolean:test` → `js-boolean:cast` → `js-number:test`
/// → `js-number:toF64` → `f64.ne` … Measured: 62 extra opcodes per assignment
/// for a check that needs three. Paying for a general case that cannot occur.
fn emit_reads_stamp(chunk: &mut Chunk, slot: u16, key: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let k = chunk.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(key)));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
}

/// Copy the value in `slot`, if its declaration said assignment copies.
/// Stack: `[] → [value]` — the copy for a value type, the original otherwise.
///
/// This is the SHARED replacement for three per-language implementations:
/// Pascal's `lower_struct_copy_assignments` walker pass, PHP's injected
/// `__php_copy_on_assign` prelude, and — for C#, Go, VB, Java and C — nothing
/// at all. None of them can act on a value that arrived from another language,
/// because each keys on its own declarations; the stamp is what makes this work
/// whoever allocated the object.
///
/// DEEP, because a value type's fields can themselves be value types. A
/// Pascal `record` holding a `record`, a C `struct` holding a `struct` and a
/// Go struct holding a struct all copy the whole tree; only a field holding a
/// genuine REFERENCE keeps sharing its referent.
///
/// This was shallow until 2026-08-06 and the bug was demonstrable in three
/// lines of Pascal — `b := a; b.I.V := 99` mutated `a.I.V`, because
/// `Object.assign` copied the inner record's REFERENCE.
///
/// The machinery lives in `primitives/clone.rs`: copying is not a record
/// concept, it is a capability records SHARE with collections, classes and
/// argument passing. This module owns only the record POLICY — which values
/// copy — and `clone.rs` owns what copying means.
///
/// Copy unconditionally — the caller already knows this is a value type from
/// its STATIC type, so no stamp read is needed. Stack: `[] → [copy]`.
///
/// The static path exists because Pascal's `var a, b: TR` DEFAULT-INITIALISES
/// and never runs a constructor, so the instance carries no stamp at all.
pub fn emit_value_copy(chunks: &mut Vec<Chunk>, current: usize, slot: u16, line: u32) {
    crate::primitives::clone::emit_deep_copy(chunks, current, slot, true, line);
}

/// Copy only if the instance says its declaration asked for it — the
/// cross-language half, for a value whose type this compiler cannot see.
pub fn emit_value_copy_if_needed(chunks: &mut Vec<Chunk>, current: usize, slot: u16, line: u32) {
    crate::primitives::clone::emit_deep_copy(chunks, current, slot, false, line);
}

/// Can this expression possibly evaluate to a value-type INSTANCE?
///
/// The runtime stamp check is what makes value semantics survive a language
/// boundary, but it is only needed when the compiler cannot tell. A literal,
/// an arithmetic result or a comparison is never a record, so the check — and
/// the copy branch around it — is pure cost there.
///
/// Measured before this: 62 extra opcodes on EVERY assignment in every
/// language, `x = 1` included. Conservative by construction — anything not
/// listed here keeps the runtime check, so a miss costs speed and never
/// correctness.
pub fn may_be_value_instance(expr: &vybe_ast::Expression) -> bool {
    use vybe_ast::{BinOp, ExprKind};
    match &expr.kind {
        // A literal is a number, string, bool or null — never an instance.
        ExprKind::Lit(_) => false,
        // Arithmetic and comparison yield primitives whatever the operands
        // were. `+` is excluded: it is also string concatenation, and some
        // languages route it through a user `__add__` that can return anything.
        ExprKind::Binary { op, .. } => !matches!(
            op,
            BinOp::Sub
                | BinOp::Mul
                | BinOp::Div
                | BinOp::IDiv
                | BinOp::Mod
                | BinOp::Pow
                | BinOp::Eq
                | BinOp::NotEq
                | BinOp::Lt
                | BinOp::Gt
                | BinOp::LtEq
                | BinOp::GtEq
                | BinOp::And
                | BinOp::Or
        ),
        _ => true }
}
