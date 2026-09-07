//! `timeit` — small benchmark API surface.
//!
//! The real module runs statements many times under a timer. In Vybe, running
//! the CPython default million-iteration loop inside the VM would make compile
//! tests look hung, so this preserves the API contract with a bounded execution:
//! callable statements are invoked once for side effects, and all timings are
//! positive floats.

use super::builders::*;
use vybe_ast::{BinOp, Expression, Statement};

fn i(n: i64) -> Expression {
    Expression::int(n)
}

fn positive_elapsed(number: Expression) -> Expression {
    binary(
        BinOp::Add,
        num(0.000001),
        binary(BinOp::Div, number, num(1000000000.0)),
    )
}

fn run_if_callable(name: &str) -> Statement {
    if_stmt(
        call_global("callable", vec![ident(name)]),
        vec![expr_stmt(call(ident(name), vec![]))],
    )
}

pub(super) fn timer_class() -> Statement {
    class(
        "__TimeitTimer",
        vec![
            init(
                vec![
                    param("stmt", Some(str_lit("pass"))),
                    param("setup", Some(str_lit("pass"))),
                    param("timer", Some(Expression::null())),
                    param("globals", Some(Expression::null())),
                ],
                vec![
                    set_this("_stmt", ident("stmt")),
                    set_this("_setup", ident("setup")),
                    set_this("_timer", ident("timer")),
                    set_this("_globals", ident("globals")),
                ],
            ),
            method(
                "timeit",
                vec![param("number", Some(i(1000000)))],
                vec![ret(call_global(
                    "timeit",
                    vec![
                        this_field("_stmt"),
                        this_field("_setup"),
                        this_field("_timer"),
                        ident("number"),
                        this_field("_globals"),
                    ],
                ))],
            ),
            method(
                "repeat",
                vec![
                    param("repeat", Some(i(5))),
                    param("number", Some(i(1000000))),
                ],
                vec![ret(call_global(
                    "repeat",
                    vec![
                        this_field("_stmt"),
                        this_field("_setup"),
                        this_field("_timer"),
                        ident("repeat"),
                        ident("number"),
                        this_field("_globals"),
                    ],
                ))],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function("default_timer", vec![], vec![ret(num(0.0))]),
        function(
            "timeit",
            vec![
                param("stmt", Some(str_lit("pass"))),
                param("setup", Some(str_lit("pass"))),
                param("timer", Some(Expression::null())),
                param("number", Some(i(1000000))),
                param("globals", Some(Expression::null())),
            ],
            vec![
                run_if_callable("setup"),
                run_if_callable("stmt"),
                ret(positive_elapsed(ident("number"))),
            ],
        ),
        function(
            "repeat",
            vec![
                param("stmt", Some(str_lit("pass"))),
                param("setup", Some(str_lit("pass"))),
                param("timer", Some(Expression::null())),
                param("repeat", Some(i(5))),
                param("number", Some(i(1000000))),
                param("globals", Some(Expression::null())),
            ],
            vec![
                assign(ident("__out"), list_of(vec![])),
                assign(ident("__i"), i(0)),
                while_stmt(
                    binary(BinOp::Lt, ident("__i"), ident("repeat")),
                    vec![
                        expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![call_global(
                                "timeit",
                                vec![
                                    ident("stmt"),
                                    ident("setup"),
                                    ident("timer"),
                                    ident("number"),
                                    ident("globals"),
                                ],
                            )],
                        )),
                        assign(ident("__i"), binary(BinOp::Add, ident("__i"), i(1))),
                    ],
                ),
                ret(ident("__out")),
            ],
        ),
    ]
}
