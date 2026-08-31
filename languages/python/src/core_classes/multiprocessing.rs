//! `multiprocessing` — `Pool`, `Manager`, `Pipe`, `Value`/`Array`.
//!
//! There is one process, so a `Pool` maps by calling in-line and a `Pipe` is a
//! pair of in-memory queues. `Process` is `Thread` under another name, which is
//! the honest model when both are the same single thread.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

/// `Process` IS a `Thread` here — the parent is the whole declaration.
pub(super) fn process() -> Statement {
    class_extending("Process", &["Thread"], vec![])
}

pub(super) fn pool() -> Statement {
    class(
        "Pool",
        vec![
            init(
                vec![param("processes", Some(null()))],
                vec![set_this("processes", ident("processes"))],
            ),
            method(
                "map",
                vec![
                    param("fn", None),
                    param("iterable", None),
                    param("chunksize", Some(null())),
                ],
                vec![
                    assign(ident("__out"), call_global("list", vec![])),
                    for_in(
                        "__item",
                        ident("iterable"),
                        vec![expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![call(ident("fn"), vec![ident("__item")])],
                        ))],
                    ),
                    ret(ident("__out")),
                ],
            ),
            method(
                "starmap",
                vec![
                    param("fn", None),
                    param("iterable", None),
                    param("chunksize", Some(null())),
                ],
                vec![
                    assign(ident("__out"), call_global("list", vec![])),
                    for_in(
                        "__args",
                        ident("iterable"),
                        vec![expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![call_spread(ident("fn"), ident("__args"))],
                        ))],
                    ),
                    ret(ident("__out")),
                ],
            ),
            method(
                "apply",
                vec![
                    param("fn", None),
                    param("args", Some(list_of(vec![]))),
                    param("kwds", Some(null())),
                ],
                vec![ret(call_spread(ident("fn"), ident("args")))],
            ),
            method("close", vec![], vec![ret(null())]),
            method("join", vec![], vec![ret(null())]),
            method("terminate", vec![], vec![ret(null())]),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method(
                "__exit__",
                any_args(),
                vec![
                    expr_stmt(call(member(ident("self"), "close"), vec![])),
                    ret(bool_lit(false)),
                ],
            ),
        ],
    )
}

pub(super) fn shared_value() -> Statement {
    class(
        "__PyValue",
        vec![init(
            vec![param("typecode", None), param("value", Some(num(0.0)))],
            vec![set_this("value", ident("value"))],
        )],
    )
}

pub(super) fn process_info() -> Statement {
    class(
        "__PyProcessInfo",
        vec![init(
            vec![param("name", None)],
            vec![set_this("name", ident("name"))],
        )],
    )
}

pub(super) fn manager() -> Statement {
    class(
        "Manager",
        vec![
            init(vec![], vec![]),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method("__exit__", any_args(), vec![ret(bool_lit(false))]),
            method(
                "list",
                vec![param("iterable", Some(null()))],
                vec![
                    if_stmt(
                        is_not_none(ident("iterable")),
                        vec![ret(call_global("list", vec![ident("iterable")]))],
                    ),
                    ret(call_global("list", vec![])),
                ],
            ),
            method(
                "dict",
                vec![param("mapping", Some(null()))],
                vec![
                    if_stmt(
                        is_not_none(ident("mapping")),
                        vec![ret(call_global("dict", vec![ident("mapping")]))],
                    ),
                    ret(call_global("dict", vec![])),
                ],
            ),
        ],
    )
}

pub(super) fn pipe_end() -> Statement {
    class(
        "__PyPipeEnd",
        vec![
            init(vec![], vec![set_this("_items", call_global("list", vec![]))]),
            method(
                "send",
                vec![param("value", None)],
                vec![
                    assign(ident("__it"), this_field("_items")),
                    expr_stmt(call(member(ident("__it"), "append"), vec![ident("value")])),
                ],
            ),
            method(
                "recv",
                vec![],
                vec![
                    assign(ident("__it"), this_field("_items")),
                    if_stmt(
                        binary(BinOp::Eq, call_global("len", vec![ident("__it")]), num(0.0)),
                        vec![ret(null())],
                    ),
                    ret(call(member(ident("__it"), "pop"), vec![num(0.0)])),
                ],
            ),
            method("close", vec![], vec![ret(null())]),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "Value",
            vec![
                param("typecode", None),
                param("value", Some(num(0.0))),
                rest_param("a"),
                kwargs_param("k"),
            ],
            vec![ret(new(
                "__PyValue",
                vec![ident("typecode"), ident("value")],
            ))],
        ),
        function(
            "Array",
            vec![
                param("typecode", None),
                param("initializer", None),
                rest_param("a"),
                kwargs_param("k"),
            ],
            vec![ret(call_global("list", vec![ident("initializer")]))],
        ),
        stub_fn("cpu_count", num(1.0)),
        function(
            "current_process",
            vec![],
            vec![ret(new("__PyProcessInfo", vec![str_lit("MainProcess")]))],
        ),
        stub_fn("active_children", list_of(vec![])),
        // `walker.rs:14849` rewrites `multiprocessing.Process(...)` to this.
        super::threading::thread_factory("__py_process_make", "Process"),
        function(
            "__py_mp_pool_factory",
            vec![param("processes", Some(null()))],
            vec![ret(new("Pool", vec![ident("processes")]))],
        ),
        function(
            "Pipe",
            vec![param("duplex", Some(bool_lit(true)))],
            vec![ret(tuple_of(vec![
                new("__PyPipeEnd", vec![]),
                new("__PyPipeEnd", vec![]),
            ]))],
        ),
    ]
}
