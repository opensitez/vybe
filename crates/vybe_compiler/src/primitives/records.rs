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

// ── Variant part — arms SHARE one region ────────────────────────────────
//
// A Pascal `case` part, a C `union` and a COBOL `REDEFINES` are ONE concept:
// several field sets naming the same storage. Three implementations existed and
// each flattened the arms into independent fields, so `u.I := 65` left `u.B`
// reading 0 where every one of those languages says 65.
//
// The region is a `DataView` over an `ArrayBuffer` sized to the widest arm, and
// each arm's fields are VIEWS onto it at their offsets. That is the only model
// that answers punning, because punning is a question about BYTES — no amount
// of field aliasing expresses "the low byte of this Integer".

/// How one variant field reads and writes the shared region.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ViewKind {
    Int8,
    Uint8,
    Int16,
    Int32,
    Uint32,
    Float64 }

impl ViewKind {
    pub fn width(self) -> u32 {
        match self {
            ViewKind::Int8 | ViewKind::Uint8 => 1,
            ViewKind::Int16 => 2,
            ViewKind::Int32 | ViewKind::Uint32 => 4,
            ViewKind::Float64 => 8 }
    }

    /// The `ecma:dataview` accessor pair. Named here rather than spelled at the
    /// emit site so the width and the accessor cannot disagree.
    pub fn accessors(self) -> (&'static str, &'static str) {
        match self {
            ViewKind::Int8 => ("getInt8", "setInt8"),
            ViewKind::Uint8 => ("getUint8", "setUint8"),
            ViewKind::Int16 => ("getInt16", "setInt16"),
            ViewKind::Int32 => ("getInt32", "setInt32"),
            ViewKind::Uint32 => ("getUint32", "setUint32"),
            ViewKind::Float64 => ("getFloat64", "setFloat64") }
    }
}

/// The view a DECLARED type takes onto the region.
///
/// `None` when the type has no byte image the platform narrows — a string, a
/// nested record, a pointer. Those keep being ordinary fields: overlapping them
/// would mean inventing a representation the language never declared, and a
/// wrong answer is worse than the honest non-overlap.
///
/// Widths come from `builtin_types::int_width_of`, the table the narrowing
/// emitter already uses, so a spelling narrows and overlaps at the same size.
pub fn view_for_hint(hint: &str) -> Option<ViewKind> {
    use vybe_ast::builtin_slots::BuiltinType;
    use vybe_ast::builtin_types::IntWidth;
    if let Some(width) = vybe_ast::builtin_types::int_width_of(hint) {
        return Some(match width {
            IntWidth::I8 => ViewKind::Int8,
            IntWidth::U8 => ViewKind::Uint8,
            IntWidth::I16 => ViewKind::Int16,
            IntWidth::I32 => ViewKind::Int32,
            IntWidth::U32 => ViewKind::Uint32 });
    }
    match vybe_ast::builtin_types::classify(hint) {
        // A boolean occupies one byte in every language that has a variant
        // part, and reading it back through the integer arm is exactly the
        // punning those languages define.
        Some(BuiltinType::Bool) => Some(ViewKind::Uint8),
        // An integer spelling the width table does not narrow (Pascal
        // `Integer`, C `int`) still has a byte image; f64 is the platform's
        // number, so a non-narrowed numeric reads and writes as one.
        Some(BuiltinType::Int) | Some(BuiltinType::Double) => Some(ViewKind::Float64),
        _ => None }
}

#[derive(Debug, Clone)]
pub struct VariantView {
    pub name: String,
    pub offset: u32,
    pub kind: ViewKind }

#[derive(Debug, Clone, Default)]
pub struct VariantLayout {
    pub views: Vec<VariantView>,
    /// Bytes in the shared region — the WIDEST arm, since all arms start at 0.
    pub size: u32 }

/// Place every arm's fields at offsets within one shared region.
///
/// Arms all start at offset 0 — that IS the overlap — and each arm lays its own
/// fields out in declaration order. `packed` drops alignment padding, which is
/// what a Pascal `packed record` and a COBOL group declare.
pub fn variant_layout(variant: &vybe_ast::VariantPart, packed: bool) -> VariantLayout {
    let mut views = Vec::new();
    let mut size = 0u32;
    for arm in &variant.arms {
        let mut offset = 0u32;
        for member in &arm.members {
            let vybe_ast::ClassMember::Field {
                name, type_hint, ..
            } = member
            else {
                continue;
            };
            let Some(kind) = type_hint.as_deref().and_then(view_for_hint) else {
                continue;
            };
            let width = kind.width();
            if !packed && width > 1 {
                // Natural alignment: a 4-byte field starts on a 4-byte
                // boundary. `packed` is the declaration that says otherwise.
                offset = offset.div_ceil(width) * width;
            }
            views.push(VariantView {
                name: name.clone(),
                offset,
                kind });
            offset += width;
        }
        size = size.max(offset);
    }
    VariantLayout { views, size }
}

/// The instance field holding the shared region.
pub const VARIANT_REGION_FIELD: &str = "__variant";

/// Every arm reads and writes the region LITTLE-ENDIAN.
///
/// Not a default worth leaving to the accessor: a `DataView` is big-endian
/// unless told otherwise, while every platform these languages target is
/// little-endian, so an unstated endianness would make `u.I := 65` read back as
/// `B = 0` — the exact bug this lowering exists to fix, reintroduced one layer
/// down.
const LITTLE_ENDIAN: bool = true;

/// Allocate the shared region on a freshly constructed instance.
/// Stack: unchanged.
pub fn emit_variant_region_init(chunk: &mut Chunk, this_slot: u16, size: u32, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_f64_const(size as f64, line);
    let buffer = chunk.add_import("ecma:arraybuffer", "newWithLength");
    chunk.emit_call(buffer, 1, line);
    let view = chunk.add_import("ecma:dataview", "new");
    chunk.emit_call(view, 1, line);
    let key = chunk.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        VARIANT_REGION_FIELD,
    )));
    chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
}

/// Push `this.__variant` for a getter/setter body whose `this` is local 0.
fn emit_region_get(chunk: &mut Chunk, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, 0, line);
    let key = chunk.add_constant(vybe_runtime::Value::String(std::sync::Arc::from(
        VARIANT_REGION_FIELD,
    )));
    chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line);
}

/// `function (this) { return dv.getX(this.__variant, offset, LE) }`
///
/// Returns the chunk index to bind with `object::emit_bind_getter`.
pub fn emit_variant_getter_chunk(chunks: &mut Vec<Chunk>, view: &VariantView) -> usize {
    const LINE: u32 = 0;
    let mut chunk = Chunk::new(&format!("__variant_get_{}", view.name));
    chunk.arity = 1; // this
    chunk.local_count = 1;
    chunks.push(chunk);
    let idx = chunks.len() - 1;

    let (getter, _) = view.kind.accessors();
    emit_region_get(&mut chunks[idx], LINE);
    chunks[idx].emit_f64_const(view.offset as f64, LINE);
    chunks[idx].emit_bool_const(LITTLE_ENDIAN, LINE);
    let f = chunks[idx].add_import("ecma:dataview", getter);
    chunks[idx].emit_call(f, 3, LINE);
    chunks[idx].emit_op(Op::RETURN, LINE);
    idx
}

/// `function (this, value) { dv.setX(this.__variant, offset, value, LE) }`
pub fn emit_variant_setter_chunk(chunks: &mut Vec<Chunk>, view: &VariantView) -> usize {
    const LINE: u32 = 0;
    let mut chunk = Chunk::new(&format!("__variant_set_{}", view.name));
    chunk.arity = 2; // this, value
    chunk.local_count = 2;
    chunks.push(chunk);
    let idx = chunks.len() - 1;

    let (_, setter) = view.kind.accessors();
    emit_region_get(&mut chunks[idx], LINE);
    chunks[idx].emit_f64_const(view.offset as f64, LINE);
    chunks[idx].emit_op_u16(Op::LOCAL_GET, 1, LINE);
    chunks[idx].emit_bool_const(LITTLE_ENDIAN, LINE);
    let f = chunks[idx].add_import("ecma:dataview", setter);
    chunks[idx].emit_call(f, 4, LINE);
    // The host call leaves its own result; drop it before pushing the return
    // value, so the RETURN sees exactly one operand.
    chunks[idx].emit_op(Op::DROP, LINE);
    // Hand back the value that was stored: a Pascal assignment yields nothing,
    // but a language whose assignment IS an expression needs one.
    chunks[idx].emit_op_u16(Op::LOCAL_GET, 1, LINE);
    chunks[idx].emit_op(Op::RETURN, LINE);
    idx
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
