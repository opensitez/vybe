//! Named-tuple normalisation — the shared, language-agnostic lowering of a
//! named tuple onto one canonical runtime shape, so a named tuple built by any
//! source language is the *same value* as one built by another.
//!
//!   walker (language-specific)          C# `(X: 1, Y: 2)`
//!       ↓  calls                        Python `namedtuple` / `NamedTuple`
//!   build_named_tuple  ← THIS MODULE    (…future languages…)
//!       ↓  produces
//!   ExprKind::NamedTuple  ← canonical node, lowered by the shared compiler
//!
//! The canonical runtime shape is the *same tagged array* a plain tuple lowers
//! to (so it indexes / iterates / `len`s / slices / unpacks for free), plus:
//!   - a by-name property per named field (`arr.X == arr[0]`),
//!   - `__fields`: the ordered field-name list (for `_asdict`/`_replace`/repr),
//!   - `__typename`: the type name (Python `namedtuple`) that selects the
//!     `Name(f=v)` repr; absent for anonymous C# named tuples, whose repr stays
//!     the positional `(a, b)` form.
//!
//! One shape means a named tuple built by any front-end is the same value:
//! C#'s `.Item1` → `t[0]` and Python's `p[0]`/`list(p)`/`len(p)`/unpack all read
//! the array backing, while `.Field` / `.x` read the by-name key. See
//! [`emit_named_tuple`].

use std::sync::Arc;
use vybe_ast::{ExprKind, Expression};
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use crate::primitives::instructions::core_wasm;

// ── Normal (positional) tuples ──────────────────────────────────────────
//
// A normal tuple keeps the same *underlying* value as a list — an
// `ObjectKind::Array`, so it slices / indexes / iterates / `len`s for free —
// but carries a hidden [`TUPLE_TAG`] property so `repr`, `type()`, and slicing
// can tell it apart from a list. The tag is the cross-language contract: a
// tuple built by any front-end is the same tagged array, and it survives
// JSON-based serialization because JSON ignores non-index array properties.

/// Hidden marker property stamped on a tuple's backing array. Language
/// front-ends opt in via the `tuple_literals_tagged` profile property.
pub const TUPLE_TAG: &str = "__tuple";

/// Stamp the tuple tag onto the array on top of stack. Stack: `[arr] -> [arr]`.
/// (`STRUCT_SET` pushes the value back, so dup the array, set, then drop it.)
pub fn emit_tag(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    c.emit_dup(line); // [arr, arr]
    core_wasm::bool_const(c, line, true); // [arr, arr, true]
    let k = c.add_constant(Value::String(Arc::from(TUPLE_TAG)));
    c.emit_struct_field_op(Op::STRUCT_SET, 0, k, line); // [arr, true]
}

/// Build a tuple from the top `n` stack values: pack them into a growable
/// `ecma:array` array, then stamp the tuple tag. This is THE way to construct
/// a tuple from computed values — the same sequence the `ExprKind::Tuple`
/// literal path uses. (`ARRAY_NEW_FIXED` must NOT be used: a fixed array
/// carries no tag and reprs as a plain list.)
/// Stack: `[v0, …, v_{n-1}] -> [tuple]`.
pub fn emit_tuple(chunks: &mut [Chunk], current: usize, n: u16, line: u32) {
    let base = chunks[current].alloc_scratch(n.max(1));
    crate::primitives::collections::emit_pack_n(chunks, current, n, base, line);
    emit_tag(chunks, current, line);
}

/// Push an i32 truthiness flag for "is the value on TOS a tagged tuple".
/// Stack: `[value] -> [i32]`. A plain list (no tag) reads null → falsy.
pub fn emit_is_tuple(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let k = c.add_constant(Value::String(Arc::from(TUPLE_TAG)));
    c.emit_struct_field_op(Op::STRUCT_GET, 0, k, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line); // 1 when tag present (non-null)
}

/// Tuple-aware value equality, for a language whose tuples compare by value but
/// whose *other* collections compare by reference (Dart: `(1,2) == (1,2)` is
/// true, `[1,2] == [1,2]` is false). Stack: `[a, b] -> [i32]`. When both sides
/// are tagged tuples it compares them structurally (positional values via
/// `JSON.stringify`, which ignores the hidden tag / by-name keys); otherwise it
/// falls back to plain reference/primitive equality. Register as a language's
/// `value_eq` hook.
pub fn emit_tuple_value_eq(chunk: &mut Chunk, line: u32) {
    let json = chunk.add_import("ecma:json", "stringify");
    let str_eq = chunk.add_import("wasm:js-string", "equals");
    let s = chunk.alloc_scratch(2);
    let (a, b) = (s, s + 1);
    chunk.emit_op_u16(Op::LOCAL_SET, b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a, line);

    // isTuple(a) && isTuple(b)
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    emit_is_tuple(std::slice::from_mut(chunk), 0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    emit_is_tuple(std::slice::from_mut(chunk), 0, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    // structural: JSON.stringify(a) == JSON.stringify(b)
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_call(json, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_call(json, 1, line);
    chunk.emit_call(str_eq, 2, line);
    chunk.emit_else(line);
    // reference / primitive equality
    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    crate::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_end(line);
}

/// Given `[result, source]`, stamp `result` as a tuple **iff** `source` was a
/// tagged tuple **and** `result` is an array; then leave `[result]`. Used by
/// slicing so `tuple[i:j]` stays a tuple while `list[i:j]` stays a list. Safe
/// for any `source`/`result` (string operands skip via the `isArray` guards),
/// so callers need not know the operand types.
pub fn emit_propagate_tag(chunks: &mut [Chunk], current: usize, line: u32) {
    let src = chunks[current].alloc_scratch(1);
    set(chunks, current, src, line); // [result]; source stashed

    // if isArray(source) && isTuple(source) && isArray(result) { tag(result) }
    get(chunks, current, src, line);
    call_is_array(chunks, current, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(chunks, current, src, line);
    emit_is_tuple(chunks, current, line);
    chunks[current].emit_if(line);
    core_wasm::dup(&mut chunks[current], line); // [result, result]
    call_is_array(chunks, current, line);
    crate::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    emit_tag(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn get(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunks: &mut [Chunk], current: usize, slot: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn call_is_array(chunks: &mut [Chunk], current: usize, line: u32) {
    let idx = chunks[current].add_import("ecma:array", "isArray");
    chunks[current].emit_call(idx, 1, line);
}

/// Structural `repr` transform shared across languages: rewrite an
/// already-formatted list string `"[a, b]"` into the tuple form `"(a, b)"`,
/// adding the single-element trailing comma `"(x,)"`. Element formatting stays
/// the front-end's concern (it produced the list string); only the brackets and
/// the lone-element comma are universal. Stack: `[list_string] -> [tuple_string]`.
pub fn emit_list_string_to_tuple(chunk: &mut Chunk, line: u32) {
    let s = chunk.alloc_scratch(4);
    let len = s + 1;
    let inner = s + 2;
    let res = s + 3;
    chunk.emit_op_u16(Op::LOCAL_SET, s, line);

    // len = s.length
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    let strlen = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(strlen, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, len, line);

    // inner = slice(s, 1, len - 1)  (ecma:array.slice dispatches string→substring)
    chunk.emit_op_u16(Op::LOCAL_GET, s, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    core_wasm::i32_const(chunk, line, 1);
    chunk.emit_op(Op::I32_SUB, line);
    let slice = chunk.add_import("ecma:array", "slice");
    chunk.emit_call(slice, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, inner, line);

    // res = "(" + inner
    core_wasm::string_const(chunk, line, "(");
    chunk.emit_op_u16(Op::LOCAL_GET, inner, line);
    crate::primitives::strings::emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, res, line);

    // single element? inner.length > 0 && inner.indexOf(", ") < 0
    chunk.emit_op_u16(Op::LOCAL_GET, inner, line);
    let ilen = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(ilen, 1, line);
    core_wasm::i32_const(chunk, line, 0);
    crate::primitives::ops::emit_dyn_gt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, inner, line);
    core_wasm::string_const(chunk, line, ", ");
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    core_wasm::i32_const(chunk, line, 0);
    crate::primitives::ops::emit_dyn_lt(chunk, line);
    crate::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, res, line);
    core_wasm::string_const(chunk, line, ",");
    crate::primitives::strings::emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, res, line);
    chunk.emit_end(line);

    // res + ")"
    chunk.emit_op_u16(Op::LOCAL_GET, res, line);
    core_wasm::string_const(chunk, line, ")");
    crate::primitives::strings::emit_str_concat(chunk, line);
}

/// Build the canonical named tuple from ordered `(name, value)` fields.
/// `name` is `None` for a positional-only element. `type_name` is the tuple's
/// type (Python `namedtuple`), or `None` for an anonymous named tuple (C#
/// `ValueTuple`). The shared compiler lowers this to a tagged array carrying
/// by-name field keys and hidden `__fields`/`__typename` (see
/// [`emit_named_tuple`]).
pub fn build_named_tuple(fields: Vec<(Option<String>, Expression)>) -> ExprKind {
    ExprKind::NamedTuple {
        fields,
        type_name: None }
}

/// Positional arity of a named-tuple literal (the field count), or `None` if
/// `expr` is not a named-tuple node.
pub fn named_tuple_arity(expr: &Expression) -> Option<usize> {
    match &expr.kind {
        ExprKind::NamedTuple { fields, .. } => Some(fields.len()),
        _ => None }
}

/// The positional read off a named-tuple value: `object[index]`. The canonical
/// shape is array-backed, so deconstruction rejoins the shared array-destructure
/// path directly.
pub fn positional_read(object: Expression, index: usize) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(Expression::int(index as i64)),
        null_safe: false })
}

// ── Named-tuple runtime metadata ────────────────────────────────────────
//
// The canonical named tuple is the positional tuple's tagged array plus:
//   - a by-name property per named field (`arr.x == arr[0]`),
//   - `__fields`: the ordered field-name list, for `_asdict`/`_replace`/repr,
//   - `__typename`: the type name (Python `namedtuple`) driving the
//     `Name(f=v)` repr; absent for anonymous (C#) named tuples, whose repr
//     stays the positional `(a, b)` form.

/// Hidden ordered field-name list stamped on a named tuple's array.
pub const FIELDS_TAG: &str = "__fields";
/// Hidden type name stamped on a named tuple's array (Python `namedtuple`).
pub const TYPENAME_TAG: &str = "__typename";

/// Stamp named-tuple metadata onto the array on TOS — already the packed
/// positional values (as for a plain tuple). Adds the `__tuple` tag, a by-name
/// key for each named field (`arr.name = arr[i]`), the ordered `__fields` name
/// list, and `__typename` when `type_name` is set. Stack: `[arr] -> [arr]`.
pub fn emit_named_tuple(
    chunks: &mut [Chunk],
    current: usize,
    field_names: &[Option<String>],
    type_name: Option<&str>,
    line: u32,
) {
    // 1. Tuple tag — the value behaves / repr's / slices as a tuple.
    emit_tag(chunks, current, line);

    // 2. By-name key per named field: `arr.<name> = arr[i]` (re-read the value
    //    from the array so field-value expressions are never re-evaluated).
    for (i, name) in field_names.iter().enumerate() {
        let Some(name) = name else { continue };
        let c = &mut chunks[current];
        c.emit_dup(line); // [arr, arr]
        c.emit_dup(line); // [arr, arr, arr]
        core_wasm::i32_const(c, line, i as i32); // [arr, arr, arr, i]
        c.emit_op(Op::ARRAY_GET, line); // [arr, arr, arr[i]]
        let k = c.add_constant(Value::String(Arc::from(name.as_str())));
        c.emit_struct_field_op(Op::STRUCT_SET, 0, k, line); // [arr, arr[i]]
    }

    // 3. Ordered field-name list.
    {
        let c = &mut chunks[current];
        c.emit_dup(line); // [arr, arr]
        for name in field_names {
            core_wasm::string_const(c, line, name.as_deref().unwrap_or(""));
        }
        c.emit_array_new_fixed(0, field_names.len() as u16, line); // [arr, fields]
        let k = c.add_constant(Value::String(Arc::from(FIELDS_TAG)));
        c.emit_struct_field_op(Op::STRUCT_SET, 0, k, line); // [arr, fields]
    }

    // 4. Type name (Python `namedtuple`) → drives the `Name(f=v)` repr.
    if let Some(tn) = type_name {
        let c = &mut chunks[current];
        c.emit_dup(line);
        core_wasm::string_const(c, line, tn);
        let k = c.add_constant(Value::String(Arc::from(TYPENAME_TAG)));
        c.emit_struct_field_op(Op::STRUCT_SET, 0, k, line);
    }
}
