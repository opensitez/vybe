//! C math.h — documents the opcode/host mappings; no new emission needed.
//!
//! Most math.h functions map directly to WASM opcodes or `ecma:math` host
//! functions via the language profile. This module provides helpers for the
//! cases that need AST-level rewriting (e.g. `round` semantics differ from
//! `f64.nearest`).

use vybe_ast::{
    Argument, BinOp, BindingPattern, ExprKind, Expression, Literal, Modifiers, Param, PassBy,
    Statement, StmtKind, VarDeclKind, VarDeclarator,
};

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn s(kind: StmtKind) -> Statement {
    Statement::new(kind)
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn ident(name: &str) -> Expression {
    e(ExprKind::Ident(name.to_string()))
}

fn lit_int(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}

fn lit_float(n: f64) -> Expression {
    e(ExprKind::Lit(Literal::Float(n)))
}

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    e(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn assign(target: Expression, value: Expression) -> Expression {
    e(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
    })
}

fn var_decl(name: &str, init: Expression) -> Statement {
    s(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Var,
    })
}

fn expr_stmt(value: Expression) -> Statement {
    s(StmtKind::Expr(value))
}

fn if_stmt(
    cond: Expression,
    then_body: Vec<Statement>,
    else_body: Option<Vec<Statement>>,
) -> Statement {
    s(StmtKind::If {
        cond,
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

fn function(name: &str, params: Vec<&str>, body: Vec<Statement>) -> Statement {
    s(StmtKind::FunctionDecl {
        name: name.to_string(),
        params: params
            .into_iter()
            .map(|param| Param {
                name: param.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect(),
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    })
}

/// math.h domain-error runtime helpers (libc surface, shared across libc-targeting
/// languages). `__c_sqrt` adds the EDOM side effect (§7.12.1) over the raw
/// `f64_sqrt` opcode (`__libc_sqrt_raw`, mapped in the profile); the walker
/// rewrites source `sqrt(x)` → `__c_sqrt(x)`. The raw path keeps results
/// bit-identical — only `errno` is touched.
///
/// ```text
/// function __c_sqrt(x) {
///   if (x < 0) errno = 33;          // EDOM
///   return __libc_sqrt_raw(x);
/// }
/// ```
pub fn domain_error_helpers() -> Vec<Statement> {
    vec![function(
        "__c_sqrt",
        vec!["x"],
        vec![
            s(StmtKind::If {
                cond: bin(BinOp::Lt, ident("x"), lit_int(0)),
                then_body: vec![
                    s(StmtKind::Expr(e(ExprKind::Assign {
                        target: Box::new(ident("errno")),
                        value: Box::new(lit_int(33)),
                    }))),
                    s(StmtKind::Expr(e(ExprKind::Assign {
                        target: Box::new(ident("__c_fenv_excepts")),
                        value: Box::new(bin(BinOp::BitOr, ident("__c_fenv_excepts"), lit_int(1))),
                    }))),
                ],
                elifs: Vec::new(),
                else_body: None,
            }),
            s(StmtKind::Return(Some(call(
                ident("__libc_sqrt_raw"),
                vec![ident("x")],
            )))),
        ],
    )]
}

/// Floating-point exception register helpers for fenv.h. This is a lightweight
/// libc-model bitset sufficient for observable C fenv APIs in the test suite.
pub fn fenv_runtime_helpers() -> Vec<Statement> {
    let inf = bin(BinOp::Div, lit_float(1.0), lit_float(0.0));
    let neg_inf = bin(BinOp::Sub, lit_float(0.0), inf.clone());
    let set_invalid_or_divzero = assign(
        ident("__c_fenv_excepts"),
        bin(
            BinOp::BitOr,
            ident("__c_fenv_excepts"),
            e(ExprKind::Ternary {
                cond: Box::new(bin(BinOp::Eq, ident("x"), lit_float(0.0))),
                then: Box::new(lit_int(1)),
                else_: Box::new(lit_int(4)),
            }),
        ),
    );
    let set_underflow = assign(
        ident("__c_fenv_excepts"),
        bin(BinOp::BitOr, ident("__c_fenv_excepts"), lit_int(16)),
    );
    let set_inexact = assign(
        ident("__c_fenv_excepts"),
        bin(BinOp::BitOr, ident("__c_fenv_excepts"), lit_int(32)),
    );
    let set_overflow = assign(
        ident("__c_fenv_excepts"),
        bin(BinOp::BitOr, ident("__c_fenv_excepts"), lit_int(8)),
    );

    vec![function(
        "__c_fenv_binary",
        vec!["op", "x", "y"],
        vec![
            var_decl(
                "r",
                e(ExprKind::Ternary {
                    cond: Box::new(bin(BinOp::Eq, ident("op"), lit_int(1))),
                    then: Box::new(bin(BinOp::Div, ident("x"), ident("y"))),
                    else_: Box::new(bin(BinOp::Mul, ident("x"), ident("y"))),
                }),
            ),
            if_stmt(
                bin(BinOp::Eq, ident("op"), lit_int(1)),
                vec![if_stmt(
                    bin(BinOp::Eq, ident("y"), lit_float(0.0)),
                    vec![expr_stmt(set_invalid_or_divzero)],
                    Some(vec![if_stmt(
                        bin(
                            BinOp::And,
                            bin(BinOp::Eq, ident("r"), lit_float(0.0)),
                            bin(BinOp::NotEq, ident("x"), lit_float(0.0)),
                        ),
                        vec![expr_stmt(set_underflow)],
                        Some(vec![if_stmt(
                            bin(
                                BinOp::NotEq,
                                bin(BinOp::Mod, ident("x"), ident("y")),
                                lit_float(0.0),
                            ),
                            vec![expr_stmt(set_inexact)],
                            None,
                        )]),
                    )]),
                )],
                Some(vec![if_stmt(
                    bin(
                        BinOp::Or,
                        bin(BinOp::Eq, ident("r"), inf),
                        bin(BinOp::Eq, ident("r"), neg_inf),
                    ),
                    vec![expr_stmt(set_overflow)],
                    None,
                )]),
            ),
            s(StmtKind::Return(Some(ident("r")))),
        ],
    )]
}

pub fn fenv_binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    let op_code = match op {
        BinOp::Div => 1,
        BinOp::Mul => 2,
        _ => 0,
    };
    call(
        ident("__c_fenv_binary"),
        vec![lit_int(op_code), left, right],
    )
}

/// C `round(x)` uses half-away-from-zero, not banker's rounding.
/// `f64.nearest` (WASM) uses banker's rounding — wrong for C.
/// Emit: `x >= 0 ? floor(x + 0.5) : ceil(x - 0.5)`
/// Uses bare `floor`/`ceil` idents so the language profile maps them to
/// `opcode:f64_floor` / `opcode:f64_ceil` (or ecma:math equivalents).
pub fn c_round(x: Expression) -> Expression {
    let half = e(ExprKind::Lit(Literal::Float(0.5)));
    let neg_half = e(ExprKind::Lit(Literal::Float(0.5)));
    let cond = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(x.clone()),
        right: Box::new(e(ExprKind::Lit(Literal::Float(0.0)))),
    });
    let pos = call(
        ident("floor"),
        vec![e(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(x.clone()),
            right: Box::new(half),
        })],
    );
    let neg = call(
        ident("ceil"),
        vec![e(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(x),
            right: Box::new(neg_half),
        })],
    );
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(pos),
        else_: Box::new(neg),
    })
}

// Profile mappings (documented, not reimplemented here):
//
//  sqrt   → opcode:f64_sqrt
//  fabs   → opcode:f64_abs
//  floor  → opcode:f64_floor
//  ceil   → opcode:f64_ceil
//  pow    → host:ecma:math:pow
//  sin    → host:ecma:math:sin
//  cos    → host:ecma:math:cos
//  tan    → host:ecma:math:tan
//  log    → host:ecma:math:log
//  log10  → host:ecma:math:log10
//  exp    → host:ecma:math:exp
//  atan2  → host:ecma:math:atan2
//  fmod   → host:ecma:math:fmod
