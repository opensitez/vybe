//! C `uchar.h` adapters.
//!
//! The current C runtime uses the wasm32-wasi representation: `char8_t` is one
//! byte, `char16_t` is a 16-bit code unit, and `char32_t` is a 32-bit code
//! point. These helpers cover the single-code-unit ASCII/UTF-8 path directly;
//! stateful multibyte sequences can be layered on top without changing callers.

use super::pointers;
use vybe_ast::{Argument, BinOp, ExprKind, Expression, Literal};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn lit_int(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}

fn ident(name: &str) -> Expression {
    e(ExprKind::Ident(name.to_string()))
}

fn member(object: Expression, field: &str) -> Expression {
    e(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    e(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn byte_at(src: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(bin(BinOp::Eq, member(src.clone(), "length"), lit_int(0))),
        then: Box::new(lit_int(0)),
        else_: Box::new(call(ident("__c_char_code_at"), vec![src, lit_int(0)])),
    })
}

/// `mbrtoc16(pc16, s, n, ps)` / `mbrtoc32(pc32, s, n, ps)` for single-byte input.
/// Stores the first byte/code point when a destination is supplied and returns
/// the number of consumed bytes, or 0 for an empty input/NUL.
pub fn mbrtoc(dst: Expression, src: Expression, n: Expression) -> Expression {
    let value = byte_at(src.clone());
    let write = pointers::carray_deref_write(dst, value.clone());
    e(ExprKind::Sequence(vec![
        write,
        e(ExprKind::Ternary {
            cond: Box::new(bin(
                BinOp::Or,
                bin(BinOp::Eq, n, lit_int(0)),
                bin(BinOp::Eq, value, lit_int(0)),
            )),
            then: Box::new(lit_int(0)),
            else_: Box::new(lit_int(1)),
        }),
    ]))
}

/// `c16rtomb(s, c16, ps)` / `c32rtomb(s, c32, ps)` for ASCII code units.
/// Stores a one-character string in the char buffer and returns one byte.
pub fn crtomb(dst: Expression, ch: Expression) -> Expression {
    e(ExprKind::Sequence(vec![
        pointers::carray_deref_write(dst, ch),
        lit_int(1),
    ]))
}
