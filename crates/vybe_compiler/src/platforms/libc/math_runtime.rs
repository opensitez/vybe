//! math.h runtime helpers (libc surface) — builders for the numeric functions
//! that aren't WASM opcodes or `ecma:math` entries and so are emitted as common
//! AST (Stirling-series `tgamma`, the `erf` polynomial, the `__c_*` math helper
//! function shells, and the bare math-call constructors). Shared by the C
//! walker (which references them from several call sites) and the libc runtime
//! prelude, so they live in one place.

use crate::ast::{Argument, BinOp, ExprKind, Expression, Literal, Statement, StmtKind};

use crate::platforms::libc::build::{expr, ident, stmt};

/// Build a private helper function (`__tgamma`, `__j0`, …) used by the math
/// prelude. The body is supplied by the caller.
pub fn build_math_helper_fn(name: &str, params: &[&str], body: Vec<Statement>) -> Statement {
    use crate::ast::{Modifiers, Param, PassBy, Visibility};
    stmt(StmtKind::FunctionDecl {
        name: name.to_string(),
        params: params
            .iter()
            .map(|p| Param {
                name: p.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect(),
        body,
        return_type: None,
        is_async: false,
        is_generator: false,
        is_sub: false,
        modifiers: Modifiers {
            visibility: Visibility::Private,
            is_static: false,
            is_abstract: false,
            is_virtual: false,
            is_override: false,
            is_readonly: false,
            is_shared: false,
            is_extension: false,
            is_overloads: false,
            is_not_overridable: false,
            decorators: vec![],
        },
        handles: vec![],
    })
}

/// Stirling approximation for tgamma used in the prelude.
/// gamma(x) ≈ sqrt(2*pi/x) * (x/e)^x * S(x), where S is the Stirling series
/// correction `1 + 1/(12x) + 1/(288 x^2) - 139/(51840 x^3)`. The bare leading
/// term underestimates by ~1.6% for small x (Γ(6) ≈ 118); the series brings it
/// to ~120.0.
pub fn stirling_approx() -> Expression {
    let two_pi = expr(ExprKind::Lit(Literal::Float(2.0 * std::f64::consts::PI)));
    let e = expr(ExprKind::Lit(Literal::Float(std::f64::consts::E)));
    let x = ident("x");
    let sqrt_2pi_over_x = ecma_math_call(
        "sqrt",
        expr(ExprKind::Binary {
            op: BinOp::Div,
            left: Box::new(two_pi),
            right: Box::new(x.clone()),
        }),
    );
    let x_over_e_pow_x = ecma_math_call2(
        "pow",
        expr(ExprKind::Binary {
            op: BinOp::Div,
            left: Box::new(x.clone()),
            right: Box::new(e),
        }),
        x.clone(),
    );
    let leading = expr(ExprKind::Binary {
        op: BinOp::Mul,
        left: Box::new(sqrt_2pi_over_x),
        right: Box::new(x_over_e_pow_x),
    });
    // x, x^2, x^3
    let xn = |n: i32| {
        if n == 1 {
            x.clone()
        } else {
            ecma_math_call2("pow", x.clone(), expr(ExprKind::Lit(Literal::Float(n as f64))))
        }
    };
    let term = |num: f64, den: f64, pow: i32| {
        expr(ExprKind::Binary {
            op: BinOp::Div,
            left: Box::new(expr(ExprKind::Lit(Literal::Float(num)))),
            right: Box::new(expr(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(expr(ExprKind::Lit(Literal::Float(den)))),
                right: Box::new(xn(pow)),
            })),
        })
    };
    // 1 + 1/(12x) + 1/(288 x^2) - 139/(51840 x^3)
    let correction = expr(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                right: Box::new(term(1.0, 12.0, 1)),
            })),
            right: Box::new(term(1.0, 288.0, 2)),
        })),
        right: Box::new(term(139.0, 51840.0, 3)),
    });
    expr(ExprKind::Binary {
        op: BinOp::Mul,
        left: Box::new(leading),
        right: Box::new(correction),
    })
}

/// Polynomial part of the erf approximation:
/// `0.254829592*t - 0.284496736*t^2 + 1.421413741*t^3 - 1.453152027*t^4 + 1.061405429*t^5`
pub fn poly_erf(t: Expression) -> Expression {
    let t2 = || expr(ExprKind::Binary { op: BinOp::Mul, left: Box::new(t.clone()), right: Box::new(t.clone()) });
    let t3 = || expr(ExprKind::Binary { op: BinOp::Mul, left: Box::new(t2()), right: Box::new(t.clone()) });
    let t4 = || expr(ExprKind::Binary { op: BinOp::Mul, left: Box::new(t3()), right: Box::new(t.clone()) });
    let t5 = || expr(ExprKind::Binary { op: BinOp::Mul, left: Box::new(t4()), right: Box::new(t.clone()) });
    let a1 = expr(ExprKind::Lit(Literal::Float(0.254829592)));
    let a2 = expr(ExprKind::Lit(Literal::Float(0.284496736)));
    let a3 = expr(ExprKind::Lit(Literal::Float(1.421413741)));
    let a4 = expr(ExprKind::Lit(Literal::Float(1.453152027)));
    let a5 = expr(ExprKind::Lit(Literal::Float(1.061405429)));
    let term1 = expr(ExprKind::Binary { op: BinOp::Mul, left: Box::new(a1), right: Box::new(t.clone()) });
    let term2 = expr(ExprKind::Binary { op: BinOp::Mul, left: Box::new(a2), right: Box::new(t2()) });
    let term3 = expr(ExprKind::Binary { op: BinOp::Mul, left: Box::new(a3), right: Box::new(t3()) });
    let term4 = expr(ExprKind::Binary { op: BinOp::Mul, left: Box::new(a4), right: Box::new(t4()) });
    let term5 = expr(ExprKind::Binary { op: BinOp::Mul, left: Box::new(a5), right: Box::new(t5()) });
    let sum12 = expr(ExprKind::Binary { op: BinOp::Sub, left: Box::new(term1), right: Box::new(term2) });
    let sum123 = expr(ExprKind::Binary { op: BinOp::Add, left: Box::new(sum12), right: Box::new(term3) });
    let sum1234 = expr(ExprKind::Binary { op: BinOp::Sub, left: Box::new(sum123), right: Box::new(term4) });
    expr(ExprKind::Binary { op: BinOp::Add, left: Box::new(sum1234), right: Box::new(term5) })
}

/// Emit a bare math function call — the C profile routes these to WASM/ecma:math.
pub fn ecma_math_call(method: &str, arg: Expression) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(ident(method)),
        args: vec![Argument::positional(arg)],
        optional: false,
    })
}

/// Emit a bare two-arg math function call.
pub fn ecma_math_call2(method: &str, a: Expression, b: Expression) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(ident(method)),
        args: vec![Argument::positional(a), Argument::positional(b)],
        optional: false,
    })
}
