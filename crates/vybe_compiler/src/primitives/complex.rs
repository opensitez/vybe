//! Complex-number AST helpers.
//!
//! The shared representation is an object `{real, imag}`. Front-ends keep the
//! source-language type knowledge, then lower complex operations through these
//! builders so compatible languages agree on the runtime shape.

use vybe_ast::{Argument, BinOp, ExprKind, Expression, Literal, ObjectProperty, UnaryOp};

fn expr(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn ident(name: &str) -> Expression {
    expr(ExprKind::Ident(name.to_string()))
}

fn str_lit(value: &str) -> Expression {
    expr(ExprKind::Lit(Literal::Str(value.to_string())))
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn bin(op: BinOp, l: Expression, r: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(l),
        right: Box::new(r),
    })
}

/// Emit a bare math function call; profiles map the callee to WASM/ecma math.
pub fn ecma_math_call(method: &str, arg: Expression) -> Expression {
    call(ident(method), vec![arg])
}

/// Emit a bare two-arg math function call.
pub fn ecma_math_call2(method: &str, a: Expression, b: Expression) -> Expression {
    call(ident(method), vec![a, b])
}

/// `{real, imag}` — the shared complex value representation.
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

/// `conj(z)` -> `{real, -imag}`.
pub fn conj(real: Expression, imag: Expression) -> Expression {
    let neg_imag = expr(ExprKind::Unary {
        op: UnaryOp::Neg,
        expr: Box::new(imag),
    });
    complex_object(real, neg_imag)
}

/// `cabs(z)` -> `sqrt(real^2 + imag^2)`.
pub fn cabs(real: Expression, imag: Expression) -> Expression {
    let re2 = bin(BinOp::Mul, real.clone(), real);
    let im2 = bin(BinOp::Mul, imag.clone(), imag);
    ecma_math_call("sqrt", bin(BinOp::Add, re2, im2))
}

/// `carg(z)` -> `atan2(imag, real)`.
pub fn carg(real: Expression, imag: Expression) -> Expression {
    ecma_math_call2("atan2", imag, real)
}
