//! Named-tuple normalisation — the shared, language-agnostic lowering of a
//! named tuple onto one canonical runtime shape, so a named tuple built by any
//! source language is the *same value* as one built by another.
//!
//!   walker (language-specific)          C# `(X: 1, Y: 2)`
//!       ↓  calls                        Python `namedtuple` / `NamedTuple`
//!   build_named_tuple  ← THIS MODULE    (…future languages…)
//!       ↓  produces
//!   ExprKind::Object  ← canonical shape, reuses the existing object runtime
//!
//! The canonical shape is a plain object (`ExprKind::Object`) carrying:
//!   - positional keys `Item1..ItemN` (1-based, matching .NET `ValueTuple`),
//!   - each field's by-name key when the field is named.
//!
//! This reuses the object runtime with no new `ObjectKind`, bytecode op, or
//! host support — exactly like anonymous types reuse `ExprKind::Object`.
//! Positional access / deconstruction reads `Item1..ItemN`; by-name access
//! reads the named key. Both work through LINQ / comprehension lambdas without
//! any element-type inference, because the names live on the value itself.

use std::sync::Arc;
use vybe_ast::{ExprKind, Expression, Literal, ObjectProperty};
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

use crate::instructions::core_wasm;

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
    c.emit_op_u16(Op::STRUCT_SET, k, line); // [arr, true]
    c.emit_op(Op::DROP, line); // [arr]
}

/// Push an i32 truthiness flag for "is the value on TOS a tagged tuple".
/// Stack: `[value] -> [i32]`. A plain list (no tag) reads null → falsy.
pub fn emit_is_tuple(chunks: &mut [Chunk], current: usize, line: u32) {
    let c = &mut chunks[current];
    let k = c.add_constant(Value::String(Arc::from(TUPLE_TAG)));
    c.emit_op_u16(Op::STRUCT_GET, k, line);
    c.emit_op(Op::REF_IS_NULL, line);
    c.emit_op(Op::I32_EQZ, line); // 1 when tag present (non-null)
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
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    get(chunks, current, src, line);
    emit_is_tuple(chunks, current, line);
    chunks[current].emit_if(line);
    core_wasm::dup(&mut chunks[current], line); // [result, result]
    call_is_array(chunks, current, line);
    crate::ops::emit_dyn_to_bool(&mut chunks[current], line);
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
    crate::strings::emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, res, line);

    // single element? inner.length > 0 && inner.indexOf(", ") < 0
    chunk.emit_op_u16(Op::LOCAL_GET, inner, line);
    let ilen = chunk.add_import("wasm:js-string", "length");
    chunk.emit_call(ilen, 1, line);
    core_wasm::i32_const(chunk, line, 0);
    crate::ops::emit_dyn_gt(chunk, line);
    crate::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, inner, line);
    core_wasm::string_const(chunk, line, ", ");
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_call(index_of, 2, line);
    core_wasm::i32_const(chunk, line, 0);
    crate::ops::emit_dyn_lt(chunk, line);
    crate::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, res, line);
    core_wasm::string_const(chunk, line, ",");
    crate::strings::emit_str_concat(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, res, line);
    chunk.emit_end(line);

    // res + ")"
    chunk.emit_op_u16(Op::LOCAL_GET, res, line);
    core_wasm::string_const(chunk, line, ")");
    crate::strings::emit_str_concat(chunk, line);
}

/// Build the canonical named-tuple object from ordered `(name, value)` fields.
/// `name` is `None` for a positional-only element. The result always carries
/// `Item1..ItemN`; named elements additionally get their by-name key.
pub fn build_named_tuple(fields: Vec<(Option<String>, Expression)>) -> ExprKind {
    let mut props = Vec::with_capacity(fields.len() * 2);
    for (i, (name, value)) in fields.into_iter().enumerate() {
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&format!("Item{}", i + 1)),
            value: value.clone(),
        });
        if let Some(n) = name {
            props.push(ObjectProperty::KeyValue {
                key: Expression::string(&n),
                value,
            });
        }
    }
    ExprKind::Object(props)
}

/// Positional arity of a canonical named-tuple object (the count of contiguous
/// `Item1..ItemN` keys), or `None` if `expr` is not such an object.
pub fn named_tuple_arity(expr: &Expression) -> Option<usize> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    let has_item = |k: usize| {
        let want = format!("Item{}", k);
        props.iter().any(|p| {
            matches!(p, ObjectProperty::KeyValue { key, .. }
                if matches!(&key.kind, ExprKind::Lit(Literal::Str(s)) if *s == want))
        })
    };
    if !has_item(1) {
        return None;
    }
    let mut n = 1;
    while has_item(n + 1) {
        n += 1;
    }
    Some(n)
}

/// The `Item{index+1}` positional read off a named-tuple value, used to lower
/// positional deconstruction back onto the shared array-destructure path.
pub fn positional_read(object: Expression, index: usize) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: format!("Item{}", index + 1),
        null_safe: false,
    })
}
