//! C string.h — expression-level adapters for string operations.
//!
//! These helpers work with the JS-string surface (simple read-only char*).
//! The mutable carray path is handled in `pointers.rs`.

use crate::ast::{Argument, BinOp, ExprKind, Expression, Literal};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn lit_int(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}

fn lit_null() -> Expression {
    e(ExprKind::Lit(Literal::Null))
}

fn member(object: Expression, field: &str) -> Expression {
    e(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

/// `String.fromCharCode(code)` — int char code → 1-char string.
pub fn char_code_to_string(code: Expression) -> Expression {
    call(
        e(ExprKind::Member {
            object: Box::new(e(ExprKind::Ident("String".to_string()))),
            field: "fromCharCode".to_string(),
            null_safe: false,
        }),
        vec![code],
    )
}

/// `s.charCodeAt(0)` — first char of string → int char code.
pub fn string_to_char_code(s: Expression) -> Expression {
    call(member(s, "charCodeAt"), vec![lit_int(0)])
}

/// `strchr(s, c_code)` — find first occurrence, return suffix or null.
/// `indexOf >= 0 ? s.slice(indexOf) : null`
pub fn strchr(s: Expression, c_code: Expression) -> Expression {
    let ch = char_code_to_string(c_code);
    let idx1 = call(member(s.clone(), "indexOf"), vec![ch.clone()]);
    let idx2 = call(member(s.clone(), "indexOf"), vec![ch]);
    let cond = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(idx1),
        right: Box::new(lit_int(0)),
    });
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(call(member(s, "slice"), vec![idx2])),
        else_: Box::new(lit_int(0)),
    })
}

/// `strrchr(s, c_code)` — find last occurrence, return suffix or null.
pub fn strrchr(s: Expression, c_code: Expression) -> Expression {
    let ch = char_code_to_string(c_code);
    let idx1 = call(member(s.clone(), "lastIndexOf"), vec![ch.clone()]);
    let idx2 = call(member(s.clone(), "lastIndexOf"), vec![ch]);
    let cond = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(idx1),
        right: Box::new(lit_int(0)),
    });
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(call(member(s, "slice"), vec![idx2])),
        else_: Box::new(lit_int(0)),
    })
}

/// `strstr(haystack, needle)` — find needle, return suffix or null.
pub fn strstr(haystack: Expression, needle: Expression) -> Expression {
    let idx1 = call(member(haystack.clone(), "indexOf"), vec![needle.clone()]);
    let idx2 = call(member(haystack.clone(), "indexOf"), vec![needle]);
    let cond = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(idx1),
        right: Box::new(lit_int(0)),
    });
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(call(member(haystack, "slice"), vec![idx2])),
        else_: Box::new(lit_int(0)),
    })
}

/// `s + n` for char pointer arithmetic on a JS string → `s.substring(n)`.
pub fn string_advance(s: Expression, n: Expression) -> Expression {
    call(member(s, "substring"), vec![n])
}
