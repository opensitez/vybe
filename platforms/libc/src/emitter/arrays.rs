//! C array creation helpers — flat WASM-style arrays used by C and Go.
//!
//! These produce the backing storage for C arrays. Unlike PHP arrays (dicts)
//! or Python lists (dynamic), these are flat index-based arrays — the WASM
//! memory model applied to a high-level representation.

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

/// Create a zero-filled flat array of `count` elements: `[0, 0, ..., 0]`.
/// Used for `int arr[n]`, `calloc(n, size)`, uninitialised buffers.
pub fn make_zero_array(count: usize) -> Expression {
    let elems = (0..count)
        .map(|_| array_elem(e(ExprKind::Lit(Literal::Int(0)))))
        .collect();
    e(ExprKind::Array(elems))
}

/// Convert a C string literal into a char-code array including null terminator.
/// `"hello"` → `[104, 101, 108, 108, 111, 0]`
///
/// This is the correct C model: char arrays are flat int arrays, not JS strings.
/// Enables true in-place mutation (`text[1] = 'o'`) and pointer arithmetic
/// over the same backing store.
pub fn carray_from_string_literal(s: &str) -> Expression {
    let mut elems: Vec<ArrayElement> = s
        .chars()
        .map(|c| array_elem(e(ExprKind::Lit(Literal::Int(c as i64)))))
        .collect();
    // null terminator
    elems.push(array_elem(e(ExprKind::Lit(Literal::Int(0)))));
    e(ExprKind::Array(elems))
}

/// Create a flat array from a list of pre-built element expressions.
pub fn make_array(elems: Vec<Expression>) -> Expression {
    e(ExprKind::Array(elems.into_iter().map(array_elem).collect()))
}

/// Zero-pad an existing array literal to `count` elements.
/// Used for `int arr[4] = {1, 2}` → `[1, 2, 0, 0]`.
pub fn zero_pad_array(mut elems: Vec<ArrayElement>, count: usize) -> Vec<ArrayElement> {
    while elems.len() < count {
        elems.push(array_elem(e(ExprKind::Lit(Literal::Int(0)))));
    }
    elems
}
