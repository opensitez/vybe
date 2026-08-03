//! Shared memory-shaped AST constructors.
//!
//! These are not raw host allocations. They normalize language-level
//! allocation/free surfaces onto the runtime values the compiler already knows
//! how to dereference and eventually gives GC/accounting one common place to
//! grow from.

use vybe_ast::{ArrayElement, ExprKind, Expression, Literal, ObjectProperty};

use crate::primitives::pointers::{CELL_KIND, REF_KIND_KEY, REF_VALUE_KEY};

fn prop(key: &str, value: Expression) -> ObjectProperty {
    ObjectProperty::KeyValue {
        key: Expression::string(key),
        value }
}

fn elem(value: Expression) -> ArrayElement {
    ArrayElement {
        value,
        spread: false,
        key: None,
        by_ref: false }
}

pub fn heap_cell(value: Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![
        prop(REF_KIND_KEY, Expression::string(CELL_KIND)),
        prop(REF_VALUE_KEY, value),
    ]))
}

pub fn heap_cell_null() -> Expression {
    heap_cell(Expression::new(ExprKind::Lit(Literal::Null)))
}

pub fn heap_array(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(values.into_iter().map(elem).collect()))
}

pub fn heap_zeroed_array(count: usize) -> Expression {
    heap_array((0..count).map(|_| Expression::int(0)).collect())
}

pub fn free_value() -> Expression {
    Expression::new(ExprKind::Lit(Literal::Null))
}
