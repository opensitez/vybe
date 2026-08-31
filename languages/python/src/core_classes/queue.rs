//! `queue` — `Queue` and its three orderings, plus `Empty` / `Full`.
//!
//! One thread of execution, so `block`/`timeout` are accepted and ignored and
//! `get` on an empty queue answers `None` rather than blocking forever. That
//! was the prelude's behaviour too.
//!
//! ⛔ `None` is stored as a SENTINEL string. A queue that holds `None` cannot
//! distinguish "empty" from "holds None" once `get` answers `None` for both,
//! and the corpus puts `None` in queues. Carried across from the prelude
//! deliberately.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

const NONE_SENTINEL: &str = "__py_queue_none__";

/// The exception pair. Parents are catchability: `except queue.Empty` is this
/// declaration.
pub(super) const EXCEPTIONS: &[(&str, &str)] = &[("Empty", "Exception"), ("Full", "Exception")];

pub(super) fn exception(name: &'static str, parent: &'static str) -> Statement {
    class_extending(name, &[parent], vec![])
}

/// `self._items` bound to a local — the receiver rule.
fn items() -> Statement {
    assign(ident("__it"), this_field("_items"))
}

pub(super) fn queue() -> Statement {
    class(
        "Queue",
        vec![
            init(
                vec![param("maxsize", Some(num(0.0)))],
                vec![
                    set_this("maxsize", ident("maxsize")),
                    set_this("_items", call_global("list", vec![])),
                    set_this("_unfinished_tasks", num(0.0)),
                ],
            ),
            method(
                "put",
                vec![
                    param("item", None),
                    param("block", Some(bool_lit(true))),
                    param("timeout", Some(null())),
                ],
                vec![
                    assign(ident("__v"), ident("item")),
                    if_stmt(
                        is_none(ident("__v")),
                        vec![assign(ident("__v"), str_lit(NONE_SENTINEL))],
                    ),
                    items(),
                    expr_stmt(call(member(ident("__it"), "append"), vec![ident("__v")])),
                    set_this(
                        "_unfinished_tasks",
                        add(this_field("_unfinished_tasks"), num(1.0)),
                    ),
                ],
            ),
            method(
                "put_nowait",
                vec![param("item", None)],
                vec![ret(call(
                    member(ident("self"), "put"),
                    vec![ident("item"), bool_lit(false)],
                ))],
            ),
            method(
                "get",
                vec![
                    param("block", Some(bool_lit(true))),
                    param("timeout", Some(null())),
                ],
                vec![
                    items(),
                    if_stmt(
                        binary(
                            BinOp::Eq,
                            call_global("len", vec![ident("__it")]),
                            num(0.0),
                        ),
                        vec![ret(null())],
                    ),
                    assign(
                        ident("__v"),
                        call(member(ident("__it"), "pop"), vec![num(0.0)]),
                    ),
                    if_stmt(
                        binary(BinOp::Eq, ident("__v"), str_lit(NONE_SENTINEL)),
                        vec![ret(null())],
                    ),
                    ret(ident("__v")),
                ],
            ),
            method(
                "get_nowait",
                vec![],
                vec![ret(call(
                    member(ident("self"), "get"),
                    vec![bool_lit(false)],
                ))],
            ),
            method(
                "empty",
                vec![],
                vec![
                    items(),
                    ret(binary(
                        BinOp::Eq,
                        call_global("len", vec![ident("__it")]),
                        num(0.0),
                    )),
                ],
            ),
            method(
                "full",
                vec![],
                vec![
                    items(),
                    ret(binary(
                        BinOp::And,
                        binary(BinOp::Gt, this_field("maxsize"), num(0.0)),
                        binary(
                            BinOp::GtEq,
                            call_global("len", vec![ident("__it")]),
                            this_field("maxsize"),
                        ),
                    )),
                ],
            ),
            method(
                "qsize",
                vec![],
                vec![items(), ret(call_global("len", vec![ident("__it")]))],
            ),
            method(
                "task_done",
                vec![],
                vec![if_stmt(
                    binary(BinOp::Gt, this_field("_unfinished_tasks"), num(0.0)),
                    vec![set_this(
                        "_unfinished_tasks",
                        binary(BinOp::Sub, this_field("_unfinished_tasks"), num(1.0)),
                    )],
                )],
            ),
            method("join", vec![], vec![ret(null())]),
        ],
    )
}

/// LIFO — the ONLY difference is which end `get` pops from.
pub(super) fn lifo_queue() -> Statement {
    class_extending(
        "LifoQueue",
        &["Queue"],
        vec![method(
            "get",
            vec![
                param("block", Some(bool_lit(true))),
                param("timeout", Some(null())),
            ],
            vec![
                items(),
                if_stmt(
                    binary(BinOp::Eq, call_global("len", vec![ident("__it")]), num(0.0)),
                    vec![ret(null())],
                ),
                assign(ident("__v"), call(member(ident("__it"), "pop"), vec![])),
                if_stmt(
                    binary(BinOp::Eq, ident("__v"), str_lit(NONE_SENTINEL)),
                    vec![ret(null())],
                ),
                ret(ident("__v")),
            ],
        )],
    )
}

/// Priority — a linear scan for the smallest, which is what the prelude did and
/// is correct for the sizes the corpus uses.
pub(super) fn priority_queue() -> Statement {
    class_extending(
        "PriorityQueue",
        &["Queue"],
        vec![method(
            "get",
            vec![
                param("block", Some(bool_lit(true))),
                param("timeout", Some(null())),
            ],
            vec![
                items(),
                if_stmt(
                    binary(BinOp::Eq, call_global("len", vec![ident("__it")]), num(0.0)),
                    vec![ret(null())],
                ),
                assign(ident("__bi"), num(0.0)),
                assign(ident("__best"), index(ident("__it"), num(0.0))),
                assign(ident("__i"), num(1.0)),
                while_stmt(
                    binary(
                        BinOp::Lt,
                        ident("__i"),
                        call_global("len", vec![ident("__it")]),
                    ),
                    vec![
                        if_stmt(
                            binary(
                                BinOp::Lt,
                                index(ident("__it"), ident("__i")),
                                ident("__best"),
                            ),
                            vec![
                                assign(ident("__best"), index(ident("__it"), ident("__i"))),
                                assign(ident("__bi"), ident("__i")),
                            ],
                        ),
                        assign(ident("__i"), add(ident("__i"), num(1.0))),
                    ],
                ),
                expr_stmt(call(member(ident("__it"), "pop"), vec![ident("__bi")])),
                if_stmt(
                    binary(BinOp::Eq, ident("__best"), str_lit(NONE_SENTINEL)),
                    vec![ret(null())],
                ),
                ret(ident("__best")),
            ],
        )],
    )
}

/// `SimpleQueue` — a `Queue` with no maxsize. The inherited constructor takes
/// the default, so the subclass adds nothing at all.
pub(super) fn simple_queue() -> Statement {
    class_extending("SimpleQueue", &["Queue"], vec![])
}
