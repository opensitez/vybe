//! C math.h — documents the opcode/host mappings; no new emission needed.
//!
//! Most math.h functions map directly to WASM opcodes or `ecma:math` host
//! functions via the language profile. This module provides helpers for the
//! cases that need AST-level rewriting (e.g. `round` semantics differ from
//! `f64.nearest`).

use crate::ast::{
    Argument, BinOp, ExprKind, Expression, Literal, Modifiers, Param, PassBy, Statement, StmtKind,
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

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    e(ExprKind::Binary { op, left: Box::new(left), right: Box::new(right) })
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
                then_body: vec![s(StmtKind::Expr(e(ExprKind::Assign {
                    target: Box::new(ident("errno")),
                    value: Box::new(lit_int(33)),
                })))],
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
