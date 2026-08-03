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

use crate::emitter::build::*;
use vybe_ast::{Argument, BinOp, ExprKind, Expression, Statement, StmtKind};

fn bin(op: BinOp, l: Expression, r: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(l),
        right: Box::new(r) })
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false })
}

fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    stmt(StmtKind::While {
        cond,
        body,
        else_body: None })
}

pub fn runtime_helpers() -> Vec<Statement> {
    vec![
        c_qsort(),
        c_bsearch_index(),
        c_qsort_key(),
        c_qsort_auto(),
        c_bsearch_index_auto(),
    ]
}

/// `__c_qsort(arr, count, cmp)` — bounded in-place bubble sort over the first
/// `count` elements; `cmp(a, b) > 0` means `a` sorts after `b`. Returns `arr`.
fn c_qsort() -> Statement {
    let swap = vec![
        var_decl_stmt("tmp", index_expr(ident("arr"), ident("j"))),
        stmt(StmtKind::Expr(assign_expr(
            index_expr(ident("arr"), ident("j")),
            index_expr(ident("arr"), bin(BinOp::Add, ident("j"), int_lit(1))),
        ))),
        stmt(StmtKind::Expr(assign_expr(
            index_expr(ident("arr"), bin(BinOp::Add, ident("j"), int_lit(1))),
            ident("tmp"),
        ))),
    ];
    let inner = while_stmt(
        bin(
            BinOp::Lt,
            ident("j"),
            bin(BinOp::Sub, ident("count"), int_lit(1)),
        ),
        vec![
            if_stmt(
                bin(
                    BinOp::Gt,
                    call(
                        ident("__c_cmp_fn"),
                        vec![
                            index_expr(ident("arr"), ident("j")),
                            index_expr(ident("arr"), bin(BinOp::Add, ident("j"), int_lit(1))),
                        ],
                    ),
                    int_lit(0),
                ),
                swap,
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                ident("j"),
                bin(BinOp::Add, ident("j"), int_lit(1)),
            ))),
        ],
    );
    let outer = while_stmt(
        bin(BinOp::Lt, ident("i"), ident("count")),
        vec![
            stmt(StmtKind::Expr(assign_expr(ident("j"), int_lit(0)))),
            inner,
            stmt(StmtKind::Expr(assign_expr(
                ident("i"),
                bin(BinOp::Add, ident("i"), int_lit(1)),
            ))),
        ],
    );
    function_stmt(
        "__c_qsort",
        vec!["arr", "count", "__c_cmp_fn"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            var_decl_stmt("j", int_lit(0)),
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
                        ident("__c_cmp_fn"),
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
        vec!["arr", "count", "key", "__c_cmp_fn"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            loop_body,
            stmt(StmtKind::Return(Some(int_lit(-1)))),
        ],
    )
}

/// `mode`: 0 = scalar/string value, 1 = `.k`, 2 = `.id`, 3 = referenced value.
fn c_qsort_key() -> Statement {
    function_stmt(
        "__c_qsort_key",
        vec!["v", "mode"],
        vec![
            if_stmt(
                bin(BinOp::Eq, ident("mode"), int_lit(1)),
                vec![stmt(StmtKind::Return(Some(member(ident("v"), "k"))))],
                None,
            ),
            if_stmt(
                bin(BinOp::Eq, ident("mode"), int_lit(2)),
                vec![stmt(StmtKind::Return(Some(member(ident("v"), "id"))))],
                None,
            ),
            if_stmt(
                bin(BinOp::Eq, ident("mode"), int_lit(3)),
                vec![stmt(StmtKind::Return(Some(expr(ExprKind::RefLoad(
                    Box::new(ident("v")),
                )))))],
                None,
            ),
            stmt(StmtKind::Return(Some(ident("v")))),
        ],
    )
}

/// `__c_qsort_auto(arr, count, mode, order)` — bounded in-place bubble sort using a
/// walker-selected key mode, avoiding fragile C comparator function callbacks.
/// Negative order sorts descending; non-negative order sorts ascending.
fn c_qsort_auto() -> Statement {
    let next_j = bin(BinOp::Add, ident("j"), int_lit(1));
    let left = index_expr(ident("arr"), ident("j"));
    let right = index_expr(ident("arr"), next_j.clone());
    let left_key = call(ident("__c_qsort_key"), vec![left.clone(), ident("mode")]);
    let right_key = call(ident("__c_qsort_key"), vec![right.clone(), ident("mode")]);
    let should_swap = expr(ExprKind::Ternary {
        cond: Box::new(bin(BinOp::Lt, ident("order"), int_lit(0))),
        then: Box::new(bin(BinOp::Lt, left_key.clone(), right_key.clone())),
        else_: Box::new(bin(BinOp::Gt, left_key, right_key)) });
    let swap = vec![
        stmt(StmtKind::Expr(assign_expr(ident("tmp"), left.clone()))),
        stmt(StmtKind::Expr(assign_expr(left, right.clone()))),
        stmt(StmtKind::Expr(assign_expr(right, ident("tmp")))),
    ];
    let inner = while_stmt(
        bin(
            BinOp::Lt,
            ident("j"),
            bin(BinOp::Sub, ident("count"), int_lit(1)),
        ),
        vec![
            if_stmt(should_swap, swap, None),
            stmt(StmtKind::Expr(assign_expr(
                ident("j"),
                bin(BinOp::Add, ident("j"), int_lit(1)),
            ))),
        ],
    );
    let outer = while_stmt(
        bin(BinOp::Lt, ident("i"), ident("count")),
        vec![
            stmt(StmtKind::Expr(assign_expr(ident("j"), int_lit(0)))),
            inner,
            stmt(StmtKind::Expr(assign_expr(
                ident("i"),
                bin(BinOp::Add, ident("i"), int_lit(1)),
            ))),
        ],
    );
    function_stmt(
        "__c_qsort_auto",
        vec!["arr", "count", "mode", "order"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            var_decl_stmt("j", int_lit(0)),
            var_decl_stmt("tmp", null_lit()),
            outer,
            stmt(StmtKind::Return(Some(ident("arr")))),
        ],
    )
}

/// `__c_bsearch_index_auto(arr, count, key, mode)` — linear scan over the
/// same key modes used by `__c_qsort_auto`.
fn c_bsearch_index_auto() -> Statement {
    let loop_body = while_stmt(
        bin(BinOp::Lt, ident("i"), ident("count")),
        vec![
            if_stmt(
                bin(
                    BinOp::Eq,
                    call(ident("__c_qsort_key"), vec![ident("key"), ident("mode")]),
                    call(
                        ident("__c_qsort_key"),
                        vec![index_expr(ident("arr"), ident("i")), ident("mode")],
                    ),
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
        "__c_bsearch_index_auto",
        vec!["arr", "count", "key", "mode"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            loop_body,
            stmt(StmtKind::Return(Some(int_lit(-1)))),
        ],
    )
}
