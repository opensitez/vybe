//! Addressable storage constructors for languages with pointer/index semantics.
//!
//! These helpers build AST values for flat, mutable, index-addressable backing
//! storage. They are intentionally separate from:
//! - `primitives/arrays.rs`, which owns compiler-side array metadata/index rules.
//! - `primitives/collections.rs`, which emits bytecode/imports for array/list/map
//!   runtime operations.
//!
//! Use this module for storage that can be addressed through `primitives/pointers`
//! or equivalent language pointer/span/reference lowering.

use vybe_ast::{ArrayElement, ExprKind, Expression, Literal};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn array_elem(value: Expression) -> ArrayElement {
    ArrayElement {
        value,
        spread: false,
        key: None,
        by_ref: false,
    }
}

/// Create a zero-filled flat storage array of `count` elements.
pub fn make_zero_storage(count: usize) -> Expression {
    let elems = (0..count)
        .map(|_| array_elem(e(ExprKind::Lit(Literal::Int(0)))))
        .collect();
    e(ExprKind::Array(elems))
}

/// Convert a string literal into a NUL-terminated codepoint storage array.
pub fn codepoints_from_string_literal(s: &str) -> Expression {
    let mut elems: Vec<ArrayElement> = s
        .chars()
        .map(|c| array_elem(e(ExprKind::Lit(Literal::Int(c as i64)))))
        .collect();
    elems.push(array_elem(e(ExprKind::Lit(Literal::Int(0)))));
    e(ExprKind::Array(elems))
}

/// Create flat storage from a list of pre-built element expressions.
pub fn make_storage(elems: Vec<Expression>) -> Expression {
    e(ExprKind::Array(elems.into_iter().map(array_elem).collect()))
}

/// Zero-pad an existing flat storage literal to `count` elements.
pub fn zero_pad_storage(mut elems: Vec<ArrayElement>, count: usize) -> Vec<ArrayElement> {
    while elems.len() < count {
        elems.push(array_elem(e(ExprKind::Lit(Literal::Int(0)))));
    }
    elems
}
