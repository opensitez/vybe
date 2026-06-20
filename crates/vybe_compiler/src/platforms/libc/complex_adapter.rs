//! C complex.h — the value-level arithmetic over a `{real, imag}` complex value.
//!
//! A complex number is modeled as an object `{real, imag}`. Deciding whether an
//! expression *is* complex (and pulling its parts) needs the front-end's type
//! knowledge, so that stays in the walker; these builders take the already
//! resolved real/imag parts and compose the standard operations. Reusable by
//! any front-end with the same `{real, imag}` representation.

use crate::ast::{BinOp, ExprKind, Expression, ObjectProperty, UnaryOp};
use crate::platforms::libc::build::{expr, str_lit};
use crate::platforms::libc::math_runtime::{ecma_math_call, ecma_math_call2};

fn bin(op: BinOp, l: Expression, r: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(l),
        right: Box::new(r),
    })
}

/// `{real, imag}` — the complex value representation.
pub fn complex_object(real: Expression, imag: Expression) -> Expression {
    expr(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: str_lit("real"),
            value: real,
        },
        ObjectProperty::KeyValue {
            key: str_lit("imag"),
            value: imag,
        },
    ]))
}

/// `conj(z)` → `{real, -imag}` (§7.3.9.4).
pub fn conj(real: Expression, imag: Expression) -> Expression {
    let neg_imag = expr(ExprKind::Unary {
        op: UnaryOp::Neg,
        expr: Box::new(imag),
    });
    complex_object(real, neg_imag)
}

/// `cabs(z)` → `sqrt(real^2 + imag^2)` — the modulus (§7.3.8.1). Always a
/// non-negative argument, so no `errno`/EDOM path is needed.
pub fn cabs(real: Expression, imag: Expression) -> Expression {
    let re2 = bin(BinOp::Mul, real.clone(), real);
    let im2 = bin(BinOp::Mul, imag.clone(), imag);
    ecma_math_call("sqrt", bin(BinOp::Add, re2, im2))
}

/// `carg(z)` → `atan2(imag, real)` — the phase angle (§7.3.9.1).
pub fn carg(real: Expression, imag: Expression) -> Expression {
    ecma_math_call2("atan2", imag, real)
}
