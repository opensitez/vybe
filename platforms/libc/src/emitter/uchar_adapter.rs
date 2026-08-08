//! C `uchar.h` adapters.
//!
//! The current C runtime uses the wasm32-wasi representation: `char8_t` is one
//! byte, `char16_t` is a 16-bit code unit, and `char32_t` is a 32-bit code
//! point. These helpers cover the single-code-unit ASCII/UTF-8 path directly;
//! stateful multibyte sequences can be layered on top without changing callers.

use vybe_ast::{Argument, BinOp, ExprKind, Expression, Literal};
use vybe_compiler::primitives::pointers;

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

/// `mbrtoc16(pc16, s, n, ps)` / `mbrtoc32(pc32, s, n, ps)` with an array
/// destination. Stores the next UTF-16 unit of the source and returns the
/// consumed count (1), or 0 for an empty input/NUL.
///
/// The C `char*` surface here IS an ECMA string, so "one byte" is one UTF-16
/// unit: a caller looping `s += mbrtoc16(...)` receives an astral character
/// as its surrogate pair across two calls — exactly the unit sequence a
/// `char16_t` buffer holds.
pub fn mbrtoc(dst: Expression, src: Expression, n: Expression) -> Expression {
    let value = byte_at(src.clone());
    let write = pointers::carray_deref_write(dst, value.clone());
    e(ExprKind::Sequence(vec![
        write,
        mbrtoc_consumed(n, value),
    ]))
}

/// `mbrtoc16(&c16, s, n, ps)` — the C11 SCALAR-destination idiom. The write
/// is a plain assignment to the addressed local.
pub fn mbrtoc_scalar(dst_lvalue: Expression, src: Expression, n: Expression) -> Expression {
    let value = byte_at(src.clone());
    e(ExprKind::Sequence(vec![
        e(ExprKind::Assign {
            target: Box::new(dst_lvalue),
            value: Box::new(value.clone()),
        }),
        mbrtoc_consumed(n, value),
    ]))
}

fn mbrtoc_consumed(n: Expression, value: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(bin(
            BinOp::Or,
            bin(BinOp::Eq, n, lit_int(0)),
            bin(BinOp::Eq, value, lit_int(0)),
        )),
        then: Box::new(lit_int(0)),
        else_: Box::new(lit_int(1)),
    })
}

/// `c16rtomb(s, c16, ps)` / `c32rtomb(s, c32, ps)` for ASCII code units.
/// Stores a one-character string in the char buffer and returns one byte.
pub fn crtomb(dst: Expression, ch: Expression) -> Expression {
    e(ExprKind::Sequence(vec![
        pointers::carray_deref_write(dst, ch),
        lit_int(1),
    ]))
}
