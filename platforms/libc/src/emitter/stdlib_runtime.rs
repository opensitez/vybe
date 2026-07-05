//! C stdlib.h runtime helpers (libc surface): `qsort` / `bsearch`.
//!
//! Previously these rode the cross-language `__stdlib_*` helper bundle; per the
//! libc-platform policy the C runtime lives here instead. The walker lowers
//! `qsort(...)`/`bsearch(...)` to calls of these helpers (adapting the C
//! function-pointer comparator to a 2-arg callable first); the bodies are
//! injected once into the program prelude.
//!
//! `__c_qsort` is a stable insertion sort over the first `count` elements,
//! sorting in place and returning the array (matches the prior bundled helper).
//! `__c_bsearch` here is a linear scan returning the matching index or -1.

use vybe_ast::{Argument, BinOp, BreakTarget, ExprKind, Expression, Statement, StmtKind};
use crate::emitter::build::*;

fn bin(op: BinOp, l: Expression, r: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(l),
        right: Box::new(r),
    })
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

pub fn runtime_helpers() -> Vec<Statement> {
    vec![c_qsort(), c_bsearch_index()]
}

/// `__c_qsort(arr, count, cmp)` — stable in-place insertion sort over the first
/// `count` elements; `cmp(a, b) > 0` means `a` sorts after `b`. Returns `arr`.
fn c_qsort() -> Statement {
    // arr[j + 1] = arr[j]
    let shift = stmt(StmtKind::Expr(assign_expr(
        index_expr(ident("arr"), bin(BinOp::Add, ident("j"), int_lit(1))),
        index_expr(ident("arr"), ident("j")),
    )));
    // while (j >= 0) { if (cmp(arr[j], key) <= 0) break; arr[j+1] = arr[j]; j -= 1; }
    let inner = while_stmt(
        bin(BinOp::GtEq, ident("j"), int_lit(0)),
        vec![
            if_stmt(
                bin(
                    BinOp::LtEq,
                    call(
                        ident("cmp"),
                        vec![index_expr(ident("arr"), ident("j")), ident("key")],
                    ),
                    int_lit(0),
                ),
                vec![stmt(StmtKind::Break(BreakTarget::Implicit))],
                None,
            ),
            shift,
            stmt(StmtKind::Expr(assign_expr(
                ident("j"),
                bin(BinOp::Sub, ident("j"), int_lit(1)),
            ))),
        ],
    );
    // while (i < count) { key = arr[i]; j = i - 1; <inner>; arr[j+1] = key; i += 1; }
    let outer = while_stmt(
        bin(BinOp::Lt, ident("i"), ident("count")),
        vec![
            var_decl_stmt("key", index_expr(ident("arr"), ident("i"))),
            var_decl_stmt("j", bin(BinOp::Sub, ident("i"), int_lit(1))),
            inner,
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("arr"), bin(BinOp::Add, ident("j"), int_lit(1))),
                ident("key"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                ident("i"),
                bin(BinOp::Add, ident("i"), int_lit(1)),
            ))),
        ],
    );
    function_stmt(
        "__c_qsort",
        vec!["arr", "count", "cmp"],
        vec![
            var_decl_stmt("i", int_lit(1)),
            outer,
            stmt(StmtKind::Return(Some(ident("arr")))),
        ],
    )
}

/// `__c_bsearch_index(arr, count, key, cmp)` — linear scan; returns the first
/// index where `cmp(key, arr[i]) == 0`, or -1.
fn c_bsearch_index() -> Statement {
    let loop_body = while_stmt(
        bin(BinOp::Lt, ident("i"), ident("count")),
        vec![
            if_stmt(
                bin(
                    BinOp::Eq,
                    call(
                        ident("cmp"),
                        vec![ident("key"), index_expr(ident("arr"), ident("i"))],
                    ),
                    int_lit(0),
                ),
                vec![stmt(StmtKind::Return(Some(ident("i"))))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                ident("i"),
                bin(BinOp::Add, ident("i"), int_lit(1)),
            ))),
        ],
    );
    function_stmt(
        "__c_bsearch_index",
        vec!["arr", "count", "key", "cmp"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            loop_body,
            stmt(StmtKind::Return(Some(int_lit(-1)))),
        ],
    )
}
