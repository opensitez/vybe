//! `concurrent.futures` — `Future` and the two executors.
//!
//! Both executors run the callable IMMEDIATELY and wrap the answer, which is
//! what a single thread of execution allows. `submit` therefore returns an
//! already-`done` future; that is the prelude's behaviour, carried across.

use super::builders::*;
use vybe_ast::Statement;

pub(super) fn future() -> Statement {
    class(
        "Future",
        vec![
            init(
                vec![param("result", Some(null()))],
                vec![
                    set_this("_result", ident("result")),
                    set_this("_done", bool_lit(true)),
                ],
            ),
            method(
                "result",
                vec![param("timeout", Some(null()))],
                vec![ret(this_field("_result"))],
            ),
            method("done", vec![], vec![ret(this_field("_done"))]),
            method("cancelled", vec![], vec![ret(bool_lit(false))]),
            method("cancel", vec![], vec![ret(bool_lit(false))]),
        ],
    )
}

/// The two executors differ only in NAME — the same body, because neither can
/// actually pool. Two rows rather than two identical builders.
pub(super) const EXECUTORS: &[&str] = &["ThreadPoolExecutor", "ProcessPoolExecutor"];

pub(super) fn executor(name: &'static str) -> Statement {
    class(
        name,
        vec![
            init(
                vec![param("max_workers", Some(null()))],
                vec![set_this("max_workers", ident("max_workers"))],
            ),
            method(
                "submit",
                vec![param("fn", None), rest_param("args"), kwargs_param("kwargs")],
                vec![
                    assign(
                        ident("__f"),
                        new("Future", vec![call_spread(ident("fn"), ident("args"))]),
                    ),
                    assign(ident("__reg"), ident("__py_executor_futures")),
                    expr_stmt(call(member(ident("__reg"), "append"), vec![ident("__f")])),
                    ret(ident("__f")),
                ],
            ),
            method(
                "map",
                vec![param("fn", None), param("iterable", None)],
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
                "shutdown",
                vec![
                    param("wait", Some(bool_lit(true))),
                    param("cancel_futures", Some(bool_lit(false))),
                ],
                vec![ret(null())],
            ),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method(
                "__exit__",
                any_args(),
                vec![
                    expr_stmt(call(member(ident("self"), "shutdown"), vec![])),
                    ret(bool_lit(false)),
                ],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("__py_executor_futures", list_of(vec![])),
        // `walker.rs:19099`/`19106` rewrite `executor.submit(...)` /
        // `executor.map(...)` to these, so they must exist as globals even
        // though the executor classes carry the same methods.
        function(
            "__py_executor_submit",
            vec![
                param("_executor", None),
                param("fn", None),
                param("arg", Some(null())),
            ],
            vec![
                assign(ident("__f"), new("Future", vec![call(ident("fn"), vec![])])),
                if_stmt(
                    is_not_none(ident("arg")),
                    vec![assign(
                        ident("__f"),
                        new("Future", vec![call(ident("fn"), vec![ident("arg")])]),
                    )],
                ),
                assign(ident("__reg"), ident("__py_executor_futures")),
                expr_stmt(call(member(ident("__reg"), "append"), vec![ident("__f")])),
                ret(ident("__f")),
            ],
        ),
        function(
            "__py_executor_map",
            vec![param("_executor", None), param("fn", None), param("iterable", None)],
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
        function(
            "as_completed",
            vec![param("fs", None), param("timeout", Some(null()))],
            vec![
                assign(ident("__out"), call_global("list", vec![])),
                assign(ident("__reg"), ident("__py_executor_futures")),
                for_in(
                    "__f",
                    ident("__reg"),
                    vec![expr_stmt(call(
                        member(ident("__out"), "append"),
                        vec![ident("__f")],
                    ))],
                ),
                if_stmt(
                    binary(
                        vybe_ast::BinOp::Gt,
                        call_global("len", vec![ident("__out")]),
                        num(0.0),
                    ),
                    vec![ret(ident("__out"))],
                ),
                for_in(
                    "__g",
                    ident("fs"),
                    vec![expr_stmt(call(
                        member(ident("__out"), "append"),
                        vec![ident("__g")],
                    ))],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "wait",
            vec![
                param("fs", None),
                param("timeout", Some(null())),
                param("return_when", Some(null())),
            ],
            vec![
                assign(ident("__done"), call_global("list", vec![])),
                for_in(
                    "__f",
                    ident("fs"),
                    vec![expr_stmt(call(
                        member(ident("__done"), "append"),
                        vec![ident("__f")],
                    ))],
                ),
                ret(tuple_of(vec![ident("__done"), list_of(vec![])])),
            ],
        ),
    ]
}
