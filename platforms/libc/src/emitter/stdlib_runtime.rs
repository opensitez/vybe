//! C stdlib.h runtime helpers (libc surface): `qsort`, `qsort_r` and `bsearch`.
//!
//! The walker lowers each call to one of these helpers, passing the C
//! comparator unchanged. The comparator receives element POINTERS, as C
//! requires: a linear address `base + i * size` when the array lives in linear
//! memory, otherwise a carray pointer `{__base, __idx}` over the managed
//! backing. Swaps move whole elements in the same backing.

use crate::emitter::build::*;
use vybe_ast::{Argument, BinOp, ExprKind, Expression, Statement, StmtKind};
use vybe_compiler::primitives::pointers;

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

fn expr_stmt(value: Expression) -> Statement {
    stmt(StmtKind::Expr(value))
}

fn set(name: &str, value: Expression) -> Statement {
    expr_stmt(assign_expr(ident(name), value))
}

pub fn runtime_helpers() -> Vec<Statement> {
    vec![
        c_sort_backing(),
        c_sort_element(),
        c_sort_compare(),
        c_sort_swap(),
        c_qsort(),
        c_bsearch(),
        c_lfind(),
    ]
}

/// The element-pointer coordinates every sort helper shares:
/// `[linear, base, offset]`. A linear array is its own address; a carray
/// pointer contributes its base and element offset; a bare managed array
/// starts at element 0.
fn c_sort_backing() -> Statement {
    let linear = expr(ExprKind::IsType {
        expr: Box::new(ident("arr")),
        type_name: "number".into(),
    });
    function_stmt(
        "__c_sort_backing",
        vec!["arr"],
        vec![
            if_stmt(
                linear,
                vec![stmt(StmtKind::Return(Some(expr(ExprKind::Array(vec![
                    array_item(int_lit(1)),
                    array_item(ident("arr")),
                    array_item(int_lit(0)),
                ])))))],
                None,
            ),
            if_stmt(
                pointers::is_carray_ptr_kind(ident("arr")),
                vec![stmt(StmtKind::Return(Some(expr(ExprKind::Array(vec![
                    array_item(int_lit(0)),
                    array_item(member(ident("arr"), "__base")),
                    array_item(member(ident("arr"), "__idx")),
                ])))))],
                None,
            ),
            stmt(StmtKind::Return(Some(expr(ExprKind::Array(vec![
                array_item(int_lit(0)),
                array_item(ident("arr")),
                array_item(int_lit(0)),
            ]))))),
        ],
    )
}

fn array_item(value: Expression) -> vybe_ast::ArrayElement {
    vybe_ast::ArrayElement {
        key: None,
        value,
        spread: false,
        by_ref: false,
    }
}

fn backing_field(index: i64) -> Expression {
    index_expr(ident("backing"), int_lit(index))
}

/// `__c_sort_element(backing, size, i)` — a C pointer to element `i`.
fn c_sort_element() -> Statement {
    function_stmt(
        "__c_sort_element",
        vec!["backing", "size", "i"],
        vec![
            if_stmt(
                bin(BinOp::Eq, backing_field(0), int_lit(1)),
                vec![stmt(StmtKind::Return(Some(bin(
                    BinOp::Add,
                    backing_field(1),
                    bin(BinOp::Mul, ident("i"), ident("size")),
                ))))],
                None,
            ),
            stmt(StmtKind::Return(Some(pointers::make_carray_ptr(
                backing_field(1),
                bin(BinOp::Add, backing_field(2), ident("i")),
            )))),
        ],
    )
}

/// `__c_sort_compare(cmp, ctx, has_ctx, a, b)` — `cmp(a, b)` or, for
/// `qsort_r`, `cmp(a, b, ctx)`.
fn c_sort_compare() -> Statement {
    function_stmt(
        "__c_sort_compare",
        vec!["cmp", "ctx", "has_ctx", "a", "b"],
        vec![
            if_stmt(
                ident("has_ctx"),
                vec![stmt(StmtKind::Return(Some(call(
                    ident("cmp"),
                    vec![ident("a"), ident("b"), ident("ctx")],
                ))))],
                None,
            ),
            stmt(StmtKind::Return(Some(call(ident("cmp"), vec![ident("a"), ident("b")])))),
        ],
    )
}

/// `__c_sort_swap(backing, size, i, j)` — exchange elements `i` and `j`:
/// `size` bytes in linear memory, or two managed slots.
fn c_sort_swap() -> Statement {
    let byte_at = |index: &str| {
        bin(
            BinOp::Add,
            bin(BinOp::Add, backing_field(1), bin(BinOp::Mul, ident(index), ident("size"))),
            ident("k"),
        )
    };
    let load = |index: &str| call(ident("__c_ptr_i32_load8_u"), vec![byte_at(index)]);
    let store = |index: &str, value: Expression| {
        expr_stmt(call(ident("__c_ptr_i32_store8"), vec![byte_at(index), value]))
    };
    let slot = |index: &str| {
        index_expr(backing_field(1), bin(BinOp::Add, backing_field(2), ident(index)))
    };
    function_stmt(
        "__c_sort_swap",
        vec!["backing", "size", "i", "j"],
        vec![
            var_decl_stmt("k", int_lit(0)),
            var_decl_stmt("t", int_lit(0)),
            if_stmt(
                bin(BinOp::Eq, backing_field(0), int_lit(1)),
                vec![
                    while_stmt(
                        bin(BinOp::Lt, ident("k"), ident("size")),
                        vec![
                            set("t", load("i")),
                            store("i", load("j")),
                            store("j", ident("t")),
                            set("k", bin(BinOp::Add, ident("k"), int_lit(1))),
                        ],
                    ),
                    stmt(StmtKind::Return(None)),
                ],
                None,
            ),
            set("t", slot("i")),
            expr_stmt(assign_expr(slot("i"), slot("j"))),
            expr_stmt(assign_expr(slot("j"), ident("t"))),
        ],
    )
}

/// `__c_qsort(arr, count, size, cmp, ctx, has_ctx)` — in-place shell sort;
/// `cmp(a, b) > 0` means element `a` sorts after element `b`.
fn c_qsort() -> Statement {
    let element = |index: Expression| {
        call(ident("__c_sort_element"), vec![ident("backing"), ident("size"), index])
    };
    let lower = bin(BinOp::Sub, ident("j"), ident("gap"));
    let out_of_order = bin(
        BinOp::Gt,
        call(
            ident("__c_sort_compare"),
            vec![
                ident("cmp"),
                ident("ctx"),
                ident("has_ctx"),
                element(lower.clone()),
                element(ident("j")),
            ],
        ),
        int_lit(0),
    );
    let sift = while_stmt(
        bin(BinOp::And, bin(BinOp::GtEq, ident("j"), ident("gap")), out_of_order),
        vec![
            expr_stmt(call(
                ident("__c_sort_swap"),
                vec![ident("backing"), ident("size"), lower.clone(), ident("j")],
            )),
            set("j", lower),
        ],
    );
    let pass = while_stmt(
        bin(BinOp::Lt, ident("i"), ident("count")),
        vec![
            set("j", ident("i")),
            sift,
            set("i", bin(BinOp::Add, ident("i"), int_lit(1))),
        ],
    );
    function_stmt(
        "__c_qsort",
        vec!["arr", "count", "size", "cmp", "ctx", "has_ctx"],
        vec![
            var_decl_stmt("backing", call(ident("__c_sort_backing"), vec![ident("arr")])),
            var_decl_stmt("gap", bin(BinOp::Shr, ident("count"), int_lit(1))),
            var_decl_stmt("i", int_lit(0)),
            var_decl_stmt("j", int_lit(0)),
            while_stmt(
                bin(BinOp::Gt, ident("gap"), int_lit(0)),
                vec![
                    set("i", ident("gap")),
                    pass,
                    set("gap", bin(BinOp::Shr, ident("gap"), int_lit(1))),
                ],
            ),
        ],
    )
}

/// `__c_bsearch(key, arr, count, size, cmp)` — binary search over a sorted
/// array; returns a pointer to a matching element, or null.
fn c_bsearch() -> Statement {
    let probe = call(
        ident("__c_sort_element"),
        vec![ident("backing"), ident("size"), ident("mid")],
    );
    function_stmt(
        "__c_bsearch",
        vec!["key", "arr", "count", "size", "cmp"],
        vec![
            var_decl_stmt("backing", call(ident("__c_sort_backing"), vec![ident("arr")])),
            var_decl_stmt("low", int_lit(0)),
            var_decl_stmt("high", ident("count")),
            var_decl_stmt("mid", int_lit(0)),
            var_decl_stmt("order", int_lit(0)),
            while_stmt(
                bin(BinOp::Lt, ident("low"), ident("high")),
                vec![
                    set("mid", bin(BinOp::Shr, bin(BinOp::Add, ident("low"), ident("high")), int_lit(1))),
                    set("order", call(ident("cmp"), vec![ident("key"), probe.clone()])),
                    if_stmt(
                        bin(BinOp::Eq, ident("order"), int_lit(0)),
                        vec![stmt(StmtKind::Return(Some(probe)))],
                        None,
                    ),
                    if_stmt(
                        bin(BinOp::Lt, ident("order"), int_lit(0)),
                        vec![set("high", ident("mid"))],
                        Some(vec![set("low", bin(BinOp::Add, ident("mid"), int_lit(1)))]),
                    ),
                ],
            ),
            stmt(StmtKind::Return(Some(null_lit()))),
        ],
    )
}

/// `__c_lfind(key, arr, count, size, cmp)` — linear search of an unsorted
/// array; returns a pointer to the first element `cmp` reports equal, or null.
fn c_lfind() -> Statement {
    let probe = call(
        ident("__c_sort_element"),
        vec![ident("backing"), ident("size"), ident("i")],
    );
    function_stmt(
        "__c_lfind",
        vec!["key", "arr", "count", "size", "cmp"],
        vec![
            var_decl_stmt("backing", call(ident("__c_sort_backing"), vec![ident("arr")])),
            var_decl_stmt("i", int_lit(0)),
            while_stmt(
                bin(BinOp::Lt, ident("i"), ident("count")),
                vec![
                    if_stmt(
                        bin(
                            BinOp::Eq,
                            call(ident("cmp"), vec![ident("key"), probe.clone()]),
                            int_lit(0),
                        ),
                        vec![stmt(StmtKind::Return(Some(probe)))],
                        None,
                    ),
                    set("i", bin(BinOp::Add, ident("i"), int_lit(1))),
                ],
            ),
            stmt(StmtKind::Return(Some(null_lit()))),
        ],
    )
}
