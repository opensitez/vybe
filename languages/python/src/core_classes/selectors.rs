//! `selectors` module classes over ordinary class/object machinery.

use super::builders::*;
use vybe_ast::{BinOp, BreakTarget, Statement, StmtKind};

type Expr = vybe_ast::Expression;

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn len_of(e: Expr) -> Expr {
    call_global("len", vec![e])
}

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn missing(e: Expr) -> Expr {
    op(BinOp::Or, is_none(e.clone()), op(BinOp::StrictEq, e, str_lit("undefined")))
}

fn break_stmt() -> Statement {
    Statement::with_span(StmtKind::Break(BreakTarget::Implicit), span())
}

fn key_for(fileobj: Expr, events: Expr, data: Expr) -> Expr {
    new(
        "SelectorKey",
        vec![
            fileobj.clone(),
            call(member(fileobj, "fileno"), vec![]),
            events,
            data,
        ],
    )
}

pub(super) fn selector_key() -> Statement {
    class(
        "SelectorKey",
        vec![init(
            vec![
                param("fileobj", None),
                param("fd", None),
                param("events", None),
                param("data", Some(null())),
            ],
            vec![
                set_this("fileobj", ident("fileobj")),
                set_this("fd", ident("fd")),
                set_this("events", ident("events")),
                set_this("data", ident("data")),
            ],
        )],
    )
}

pub(super) fn select_selector() -> Statement {
    class(
        "SelectSelector",
        vec![
            init(vec![], vec![set_this("_map", dict_of(vec![]))]),
            method(
                "register",
                vec![
                    param("fileobj", None),
                    param("events", None),
                    param("data", Some(null())),
                ],
                vec![
                    if_stmt(
                        contains(this_field("_map"), ident("fileobj")),
                        vec![raise_call("KeyError", vec![ident("fileobj")])],
                    ),
                    assign(
                        ident("__key"),
                        key_for(ident("fileobj"), ident("events"), ident("data")),
                    ),
                    assign(index(this_field("_map"), ident("fileobj")), ident("__key")),
                    ret(ident("__key")),
                ],
            ),
            method(
                "unregister",
                vec![param("fileobj", None)],
                vec![
                    if_stmt(
                        unary_not(contains(this_field("_map"), ident("fileobj"))),
                        vec![raise_call("KeyError", vec![ident("fileobj")])],
                    ),
                    assign(ident("__out"), index(this_field("_map"), ident("fileobj"))),
                    expr_stmt(call(member(this_field("_map"), "pop"), vec![ident("fileobj")])),
                    ret(ident("__out")),
                ],
            ),
            method(
                "modify",
                vec![
                    param("fileobj", None),
                    param("events", None),
                    param("data", Some(null())),
                ],
                vec![
                    if_stmt(
                        unary_not(contains(this_field("_map"), ident("fileobj"))),
                        vec![raise_call("KeyError", vec![ident("fileobj")])],
                    ),
                    assign(
                        ident("__old"),
                        index(this_field("_map"), ident("fileobj")),
                    ),
                    assign(
                        ident("__data"),
                        ternary(is_none(ident("data")), field_of(ident("__old"), "data"), ident("data")),
                    ),
                    assign(
                        ident("__key"),
                        key_for(ident("fileobj"), ident("events"), ident("__data")),
                    ),
                    assign(index(this_field("_map"), ident("fileobj")), ident("__key")),
                    ret(ident("__key")),
                ],
            ),
            method(
                "get_key",
                vec![param("fileobj", None)],
                vec![
                    if_stmt(
                        unary_not(contains(this_field("_map"), ident("fileobj"))),
                        vec![raise_call("KeyError", vec![ident("fileobj")])],
                    ),
                    ret(index(this_field("_map"), ident("fileobj"))),
                ],
            ),
            method("get_map", vec![], vec![ret(this_field("_map"))]),
            method("select", any_args(), vec![ret(list_of(vec![]))]),
            method("close", vec![], vec![set_this("_map", dict_of(vec![])), ret(null())]),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method(
                "__exit__",
                any_args(),
                vec![expr_stmt(call(member(ident("self"), "close"), vec![])), ret(bool_lit(false))],
            ),
        ],
    )
}

pub(super) fn epoll_selector() -> Statement {
    class_extending("EpollSelector", &["SelectSelector"], vec![])
}

pub(super) fn kqueue_selector() -> Statement {
    class_extending("KqueueSelector", &["SelectSelector"], vec![])
}

pub(super) fn poll_selector() -> Statement {
    class_extending("PollSelector", &["SelectSelector"], vec![])
}

pub(super) fn devpoll_selector() -> Statement {
    class_extending("DevpollSelector", &["SelectSelector"], vec![])
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("EVENT_READ", num(1.0)),
        global_assign("EVENT_WRITE", num(2.0)),
    ]
}
