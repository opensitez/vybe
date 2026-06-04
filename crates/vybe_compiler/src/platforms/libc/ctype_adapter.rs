//! C ctype.h — inline arithmetic classification functions.
//!
//! All functions take an integer char code (not a string) and return an
//! integer 0/1 value (not a boolean). No host function calls — pure WASM
//! integer comparisons.
//!
//! Shared by C and any other language whose character classification maps
//! to ASCII/ISO-8859 semantics.

use crate::ast::{BinOp, ExprKind, Expression, Literal, UnaryOp};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn lit(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}

/// `c >= lo && c <= hi` — inclusive integer range check.
pub fn int_range(c: Expression, lo: i64, hi: i64) -> Expression {
    let ge = e(ExprKind::Binary { op: BinOp::GtEq, left: Box::new(c.clone()), right: Box::new(lit(lo)) });
    let le = e(ExprKind::Binary { op: BinOp::LtEq, left: Box::new(c), right: Box::new(lit(hi)) });
    e(ExprKind::Binary { op: BinOp::And, left: Box::new(ge), right: Box::new(le) })
}

/// Normalise a boolean expression to C int semantics: `expr ? 1 : 0`.
pub fn bool_to_int(b: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(b),
        then: Box::new(lit(1)),
        else_: Box::new(lit(0)),
    })
}

/// `isalpha(c)`: `(c >= 65 && c <= 90) || (c >= 97 && c <= 122)`
pub fn c_isalpha(c: Expression) -> Expression {
    let upper = int_range(c.clone(), 65, 90);
    let lower = int_range(c, 97, 122);
    bool_to_int(e(ExprKind::Binary { op: BinOp::Or, left: Box::new(upper), right: Box::new(lower) }))
}

/// `isdigit(c)`: `c >= 48 && c <= 57`
pub fn c_isdigit(c: Expression) -> Expression {
    bool_to_int(int_range(c, 48, 57))
}

/// `isalnum(c)`: `isalpha || isdigit`
pub fn c_isalnum(c: Expression) -> Expression {
    let alpha = c_isalpha(c.clone());
    let digit = c_isdigit(c);
    // Both already return 0/1; bitwise-or combines correctly.
    e(ExprKind::Binary { op: BinOp::BitOr, left: Box::new(alpha), right: Box::new(digit) })
}

/// `isspace(c)`: space (32) or control whitespace (9–13: \t \n \v \f \r)
pub fn c_isspace(c: Expression) -> Expression {
    let sp = e(ExprKind::Binary { op: BinOp::Eq, left: Box::new(c.clone()), right: Box::new(lit(32)) });
    let ctrl = int_range(c, 9, 13);
    bool_to_int(e(ExprKind::Binary { op: BinOp::Or, left: Box::new(sp), right: Box::new(ctrl) }))
}

/// `isupper(c)`: `c >= 65 && c <= 90`
pub fn c_isupper(c: Expression) -> Expression {
    bool_to_int(int_range(c, 65, 90))
}

/// `islower(c)`: `c >= 97 && c <= 122`
pub fn c_islower(c: Expression) -> Expression {
    bool_to_int(int_range(c, 97, 122))
}

/// `isxdigit(c)`: digit or A–F or a–f
pub fn c_isxdigit(c: Expression) -> Expression {
    let dig = int_range(c.clone(), 48, 57);
    let uf  = int_range(c.clone(), 65, 70);
    let lf  = int_range(c, 97, 102);
    bool_to_int(e(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(dig),
        right: Box::new(e(ExprKind::Binary { op: BinOp::Or, left: Box::new(uf), right: Box::new(lf) })),
    }))
}

/// `iscntrl(c)`: `c < 32 || c == 127`
pub fn c_iscntrl(c: Expression) -> Expression {
    let lt32  = e(ExprKind::Binary { op: BinOp::Lt, left: Box::new(c.clone()), right: Box::new(lit(32)) });
    let eq127 = e(ExprKind::Binary { op: BinOp::Eq, left: Box::new(c), right: Box::new(lit(127)) });
    e(ExprKind::Binary { op: BinOp::Or, left: Box::new(lt32), right: Box::new(eq127) })
}

/// `isprint(c)`: `c >= 32 && c < 127`
pub fn c_isprint(c: Expression) -> Expression {
    let ge32  = e(ExprKind::Binary { op: BinOp::GtEq, left: Box::new(c.clone()), right: Box::new(lit(32)) });
    let lt127 = e(ExprKind::Binary { op: BinOp::Lt,   left: Box::new(c),         right: Box::new(lit(127)) });
    e(ExprKind::Binary { op: BinOp::And, left: Box::new(ge32), right: Box::new(lt127) })
}

/// `ispunct(c)`: printable non-space non-alnum: `c >= 33 && c <= 126 && !isalnum`
pub fn c_ispunct(c: Expression) -> Expression {
    let ge33   = e(ExprKind::Binary { op: BinOp::GtEq, left: Box::new(c.clone()), right: Box::new(lit(33)) });
    let le126  = e(ExprKind::Binary { op: BinOp::LtEq, left: Box::new(c.clone()), right: Box::new(lit(126)) });
    let not_an = e(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(c_isalnum(c)) });
    e(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(e(ExprKind::Binary { op: BinOp::And, left: Box::new(ge33), right: Box::new(le126) })),
        right: Box::new(not_an),
    })
}

/// `toupper(c)`: if lowercase add 32 offset difference (65-97 = -32)
pub fn c_toupper(c: Expression) -> Expression {
    // c - 32 when islower, else c
    e(ExprKind::Ternary {
        cond: Box::new(int_range(c.clone(), 97, 122)),
        then: Box::new(e(ExprKind::Binary { op: BinOp::Sub, left: Box::new(c.clone()), right: Box::new(lit(32)) })),
        else_: Box::new(c),
    })
}

/// `tolower(c)`: if uppercase add 32
pub fn c_tolower(c: Expression) -> Expression {
    e(ExprKind::Ternary {
        cond: Box::new(int_range(c.clone(), 65, 90)),
        then: Box::new(e(ExprKind::Binary { op: BinOp::Add, left: Box::new(c.clone()), right: Box::new(lit(32)) })),
        else_: Box::new(c),
    })
}
