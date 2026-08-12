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

fn unary(op: UnaryOp, value: Expression) -> Expression {
    expr(ExprKind::Unary {
        op,
        expr: Box::new(value),
    })
}

fn int_lit(value: i64) -> Expression {
    expr(ExprKind::Lit(Literal::Int(value)))
}

fn float_lit(value: f64) -> Expression {
    expr(ExprKind::Lit(Literal::Float(value)))
}

fn member(object: Expression, field: &str) -> Expression {
    expr(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
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

pub fn real_part(value: Expression) -> Expression {
    member(value, "real")
}

pub fn imag_part(value: Expression) -> Expression {
    member(value, "imag")
}

pub fn add(a_re: Expression, a_im: Expression, b_re: Expression, b_im: Expression) -> Expression {
    complex_object(bin(BinOp::Add, a_re, b_re), bin(BinOp::Add, a_im, b_im))
}

pub fn sub(a_re: Expression, a_im: Expression, b_re: Expression, b_im: Expression) -> Expression {
    complex_object(bin(BinOp::Sub, a_re, b_re), bin(BinOp::Sub, a_im, b_im))
}

pub fn mul(a_re: Expression, a_im: Expression, b_re: Expression, b_im: Expression) -> Expression {
    complex_object(
        bin(
            BinOp::Sub,
            bin(BinOp::Mul, a_re.clone(), b_re.clone()),
            bin(BinOp::Mul, a_im.clone(), b_im.clone()),
        ),
        bin(BinOp::Add, bin(BinOp::Mul, a_re, b_im), bin(BinOp::Mul, a_im, b_re)),
    )
}

pub fn div(a_re: Expression, a_im: Expression, b_re: Expression, b_im: Expression) -> Expression {
    let denom = bin(
        BinOp::Add,
        bin(BinOp::Mul, b_re.clone(), b_re.clone()),
        bin(BinOp::Mul, b_im.clone(), b_im.clone()),
    );
    complex_object(
        bin(
            BinOp::Div,
            bin(
                BinOp::Add,
                bin(BinOp::Mul, a_re.clone(), b_re.clone()),
                bin(BinOp::Mul, a_im.clone(), b_im.clone()),
            ),
            denom.clone(),
        ),
        bin(
            BinOp::Div,
            bin(BinOp::Sub, bin(BinOp::Mul, a_im, b_re), bin(BinOp::Mul, a_re, b_im)),
            denom,
        ),
    )
}

/// `conj(z)` -> `{real, -imag}`.
pub fn conj(real: Expression, imag: Expression) -> Expression {
    let neg_imag = unary(UnaryOp::Neg, imag);
    complex_object(real, neg_imag)
}

/// `cabs(z)` -> `sqrt(real^2 + imag^2)`.
pub fn cabs(real: Expression, imag: Expression) -> Expression {
    let re2 = bin(BinOp::Mul, real.clone(), real);
    let im2 = bin(BinOp::Mul, imag.clone(), imag);
    ecma_math_call("sqrt", bin(BinOp::Add, re2, im2))
}

/// `abs2(z)` -> `real^2 + imag^2`.
pub fn abs2(real: Expression, imag: Expression) -> Expression {
    let re2 = bin(BinOp::Mul, real.clone(), real);
    let im2 = bin(BinOp::Mul, imag.clone(), imag);
    bin(BinOp::Add, re2, im2)
}

/// `carg(z)` -> `atan2(imag, real)`.
pub fn carg(real: Expression, imag: Expression) -> Expression {
    ecma_math_call2("atan2", imag, real)
}

pub fn exp(real: Expression, imag: Expression) -> Expression {
    let mag = ecma_math_call("exp", real);
    complex_object(
        bin(BinOp::Mul, mag.clone(), ecma_math_call("cos", imag.clone())),
        bin(BinOp::Mul, mag, ecma_math_call("sin", imag)),
    )
}

pub fn log(real: Expression, imag: Expression) -> Expression {
    complex_object(ecma_math_call("log", cabs(real.clone(), imag.clone())), carg(real, imag))
}

pub fn sqrt(real: Expression, imag: Expression) -> Expression {
    let r = cabs(real.clone(), imag.clone());
    let real_part = ecma_math_call(
        "sqrt",
        bin(BinOp::Div, bin(BinOp::Add, r.clone(), real.clone()), int_lit(2)),
    );
    let imag_mag = ecma_math_call(
        "sqrt",
        bin(BinOp::Div, bin(BinOp::Sub, r, real.clone()), int_lit(2)),
    );
    let sign = expr(ExprKind::Ternary {
        cond: Box::new(bin(BinOp::Lt, imag.clone(), int_lit(0))),
        then: Box::new(int_lit(-1)),
        else_: Box::new(int_lit(1)),
    });
    let imag_part = expr(ExprKind::Ternary {
        cond: Box::new(bin(BinOp::Eq, imag, int_lit(0))),
        then: Box::new(expr(ExprKind::Ternary {
            cond: Box::new(bin(BinOp::Lt, real, int_lit(0))),
            then: Box::new(imag_mag.clone()),
            else_: Box::new(int_lit(0)),
        })),
        else_: Box::new(bin(BinOp::Mul, sign, imag_mag)),
    });
    complex_object(real_part, imag_part)
}

pub fn sin(real: Expression, imag: Expression) -> Expression {
    complex_object(
        bin(
            BinOp::Mul,
            ecma_math_call("sin", real.clone()),
            ecma_math_call("cosh", imag.clone()),
        ),
        bin(BinOp::Mul, ecma_math_call("cos", real), ecma_math_call("sinh", imag)),
    )
}

pub fn cos(real: Expression, imag: Expression) -> Expression {
    complex_object(
        bin(
            BinOp::Mul,
            ecma_math_call("cos", real.clone()),
            ecma_math_call("cosh", imag.clone()),
        ),
        unary(
            UnaryOp::Neg,
            bin(BinOp::Mul, ecma_math_call("sin", real), ecma_math_call("sinh", imag)),
        ),
    )
}

pub fn asin(real: Expression, _imag: Expression) -> Expression {
    complex_object(
        expr(ExprKind::Ternary {
            cond: Box::new(bin(BinOp::Gt, real.clone(), int_lit(1))),
            then: Box::new(float_lit(std::f64::consts::FRAC_PI_2)),
            else_: Box::new(ecma_math_call("asin", real)),
        }),
        int_lit(0),
    )
}

pub fn acos(real: Expression, _imag: Expression) -> Expression {
    let acosh = ecma_math_call(
        "log",
        bin(
            BinOp::Add,
            real.clone(),
            ecma_math_call(
                "sqrt",
                bin(BinOp::Sub, bin(BinOp::Mul, real.clone(), real), int_lit(1)),
            ),
        ),
    );
    complex_object(int_lit(0), unary(UnaryOp::Neg, acosh))
}

pub fn atan(_real: Expression, imag: Expression) -> Expression {
    let imag_part = bin(
        BinOp::Mul,
        float_lit(0.5),
        ecma_math_call(
            "log",
            bin(
                BinOp::Div,
                bin(BinOp::Add, imag.clone(), int_lit(1)),
                bin(BinOp::Sub, imag, int_lit(1)),
            ),
        ),
    );
    complex_object(int_lit(0), imag_part)
}

pub fn tan(real: Expression, imag: Expression) -> Expression {
    let two_real = bin(BinOp::Mul, int_lit(2), real);
    let two_imag = bin(BinOp::Mul, int_lit(2), imag);
    let denom = bin(
        BinOp::Add,
        ecma_math_call("cos", two_real.clone()),
        ecma_math_call("cosh", two_imag.clone()),
    );
    complex_object(
        bin(BinOp::Div, ecma_math_call("sin", two_real), denom.clone()),
        bin(BinOp::Div, ecma_math_call("sinh", two_imag), denom),
    )
}

pub fn pow(
    base_re: Expression,
    base_im: Expression,
    exp_re: Expression,
    exp_im: Expression,
) -> Expression {
    let log_base = log(base_re, base_im);
    let product = mul(exp_re, exp_im, real_part(log_base.clone()), imag_part(log_base));
    exp(real_part(product.clone()), imag_part(product))
}

pub fn sinh(real: Expression, imag: Expression) -> Expression {
    complex_object(
        bin(
            BinOp::Mul,
            ecma_math_call("sinh", real.clone()),
            ecma_math_call("cos", imag.clone()),
        ),
        bin(BinOp::Mul, ecma_math_call("cosh", real), ecma_math_call("sin", imag)),
    )
}

pub fn cosh(real: Expression, imag: Expression) -> Expression {
    complex_object(
        bin(
            BinOp::Mul,
            ecma_math_call("cosh", real.clone()),
            ecma_math_call("cos", imag.clone()),
        ),
        bin(BinOp::Mul, ecma_math_call("sinh", real), ecma_math_call("sin", imag)),
    )
}

pub fn tanh(real: Expression, imag: Expression) -> Expression {
    let two_real = bin(BinOp::Mul, int_lit(2), real);
    let two_imag = bin(BinOp::Mul, int_lit(2), imag);
    let denom = bin(
        BinOp::Add,
        ecma_math_call("cosh", two_real.clone()),
        ecma_math_call("cos", two_imag.clone()),
    );
    complex_object(
        bin(BinOp::Div, ecma_math_call("sinh", two_real), denom.clone()),
        bin(BinOp::Div, ecma_math_call("sin", two_imag), denom),
    )
}

pub fn polar(real: Expression, imag: Expression) -> Expression {
    expr(ExprKind::Tuple(vec![cabs(real.clone(), imag.clone()), carg(real, imag)]))
}

pub fn rect(radius: Expression, phase: Expression) -> Expression {
    complex_object(
        bin(BinOp::Mul, radius.clone(), ecma_math_call("cos", phase.clone())),
        bin(BinOp::Mul, radius, ecma_math_call("sin", phase)),
    )
}

pub fn proj(real: Expression, imag: Expression) -> Expression {
    let inf = bin(BinOp::Div, float_lit(1.0), float_lit(0.0));
    expr(ExprKind::Ternary {
        cond: Box::new(bin(
            BinOp::Or,
            bin(BinOp::Eq, imag.clone(), inf.clone()),
            bin(BinOp::Eq, imag, unary(UnaryOp::Neg, inf.clone())),
        )),
        then: Box::new(complex_object(inf, int_lit(0))),
        else_: Box::new(complex_object(real, int_lit(0))),
    })
}
