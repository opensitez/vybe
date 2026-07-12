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

use vybe_ast::{ExprKind, Expression, Literal, ObjectProperty};

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
