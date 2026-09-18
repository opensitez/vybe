//! `asyncio` surface adapters.
//!
//! Python async lowers through the common async primitives in the walker for
//! `run`/`gather`/`sleep`/`create_task`. This file supplies the module-level
//! class/function surface that tests and user code also observe.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

pub(super) fn task() -> Statement {
    class_extending(
        "Task",
        &["Future"],
        vec![
            init(
                vec![param("result", Some(null())), param("name", Some(null()))],
                vec![
                    set_this("_result", ident("result")),
                    set_this("_done", bool_lit(true)),
                    set_this("_cancelled", bool_lit(false)),
                    set_this("_callbacks", list_of(vec![])),
                    set_this("_result_append_sinks", list_of(vec![])),
                    set_this("_name", ident("name")),
                ],
            ),
            method("get_name", vec![], vec![ret(this_field("_name"))]),
            method(
                "set_name",
                vec![param("name", None)],
                vec![set_this("_name", ident("name"))],
            ),
            method("cancelled", vec![], vec![ret(this_field("_cancelled"))]),
            method(
                "cancel",
                vec![],
                vec![
                    set_this("_cancelled", bool_lit(true)),
                    set_this("_done", bool_lit(true)),
                    ret(bool_lit(true)),
                ],
            ),
        ],
    )
}

pub(super) fn event_loop() -> Statement {
    class(
        "__PyAsyncEventLoop",
        vec![
            init(vec![], vec![]),
            method("time", vec![], vec![ret(num(0.0))]),
            method("is_running", vec![], vec![ret(bool_lit(true))]),
            method("is_closed", vec![], vec![ret(bool_lit(false))]),
            method("close", vec![], vec![ret(null())]),
            method(
                "create_future",
                vec![],
                vec![ret(new("Future", vec![null()]))],
            ),
            method(
                "call_soon_threadsafe",
                vec![param("callback", None), rest_param("args")],
                vec![ret(call_spread(ident("callback"), ident("args")))],
            ),
            method(
                "run_until_complete",
                vec![param("awaitable", None)],
                vec![ret(ident("awaitable"))],
            ),
        ],
    )
}

pub(super) fn task_group() -> Statement {
    class(
        "TaskGroup",
        vec![
            init(vec![], vec![set_this("_tasks", list_of(vec![]))]),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method("__exit__", any_args(), vec![ret(bool_lit(false))]),
            method("__aenter__", vec![], vec![ret(ident("self"))]),
            method("__aexit__", any_args(), vec![ret(bool_lit(false))]),
            method(
                "create_task",
                vec![param("coro", None), param("name", Some(null()))],
                vec![
                    assign(
                        ident("__t"),
                        new("Task", vec![ident("coro"), ident("name")]),
                    ),
                    expr_stmt(call(
                        member(this_field("_tasks"), "append"),
                        vec![ident("__t")],
                    )),
                    ret(ident("__t")),
                ],
            ),
        ],
    )
}

pub(super) fn timeout() -> Statement {
    class(
        "Timeout",
        vec![
            init(
                vec![param("delay", Some(null()))],
                vec![set_this("delay", ident("delay"))],
            ),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method("__exit__", any_args(), vec![ret(bool_lit(false))]),
            method("__aenter__", vec![], vec![ret(ident("self"))]),
            method("__aexit__", any_args(), vec![ret(bool_lit(false))]),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("__py_async_loop", new("__PyAsyncEventLoop", vec![])),
        global_assign(
            "__py_async_current_task",
            new("Task", vec![null(), str_lit("Task-1")]),
        ),
        global_assign("CancelledError", ident("Exception")),
        global_assign("TimeoutError", ident("Exception")),
        global_assign("FIRST_COMPLETED", str_lit("FIRST_COMPLETED")),
        global_assign("FIRST_EXCEPTION", str_lit("FIRST_EXCEPTION")),
        global_assign("ALL_COMPLETED", str_lit("ALL_COMPLETED")),
        function("run", vec![param("main", None)], vec![ret(ident("main"))]),
        function(
            "get_running_loop",
            vec![],
            vec![ret(ident("__py_async_loop"))],
        ),
        function(
            "get_event_loop",
            vec![],
            vec![ret(ident("__py_async_loop"))],
        ),
        function(
            "new_event_loop",
            vec![],
            vec![ret(new("__PyAsyncEventLoop", vec![]))],
        ),
        function(
            "set_event_loop",
            vec![param("loop", None)],
            vec![ret(null())],
        ),
        function("get_event_loop_policy", vec![], vec![ret(dict_of(vec![]))]),
        function(
            "current_task",
            any_args(),
            vec![ret(ident("__py_async_current_task"))],
        ),
        function(
            "all_tasks",
            any_args(),
            vec![ret(list_of(vec![ident("__py_async_current_task")]))],
        ),
        function(
            "__py_future_add_result_append_sink",
            vec![param("future", None), param("sink", None)],
            vec![
                assign(
                    ident("__sinks"),
                    field_of(ident("future"), "_result_append_sinks"),
                ),
                expr_stmt(call(
                    member(ident("__sinks"), "append"),
                    vec![ident("sink")],
                )),
                ret(null()),
            ],
        ),
        function(
            "create_task",
            vec![param("coro", None), param("name", Some(null()))],
            vec![ret(new("Task", vec![ident("coro"), ident("name")]))],
        ),
        function(
            "ensure_future",
            vec![param("coro", None)],
            vec![ret(ident("coro"))],
        ),
        function(
            "wait_for",
            vec![param("aw", None), param("timeout", Some(null()))],
            vec![ret(ident("aw"))],
        ),
        function("shield", vec![param("aw", None)], vec![ret(ident("aw"))]),
        function(
            "wait",
            vec![
                param("aws", None),
                param("timeout", Some(null())),
                param("return_when", Some(null())),
            ],
            vec![ret(tuple_of(vec![
                call_global("list", vec![ident("aws")]),
                list_of(vec![]),
            ]))],
        ),
        function(
            "as_completed",
            vec![param("aws", None), param("timeout", Some(null()))],
            vec![ret(call_global("list", vec![ident("aws")]))],
        ),
        function(
            "to_thread",
            vec![param("fn", None), rest_param("args")],
            vec![ret(call_spread(ident("fn"), ident("args")))],
        ),
        function(
            "run_coroutine_threadsafe",
            vec![param("coro", None), param("loop", None)],
            vec![ret(new("Future", vec![ident("coro")]))],
        ),
        function(
            "__py_asyncio_unwrap_settled",
            vec![param("results", None)],
            vec![
                assign(ident("__out"), list_of(vec![])),
                for_in(
                    "__item",
                    ident("results"),
                    vec![if_else(
                        binary(
                            BinOp::Eq,
                            field_of(ident("__item"), "status"),
                            str_lit("fulfilled"),
                        ),
                        vec![expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![field_of(ident("__item"), "value")],
                        ))],
                        vec![expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![field_of(ident("__item"), "reason")],
                        ))],
                    )],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "timeout",
            vec![param("delay", Some(null()))],
            vec![ret(new("Timeout", vec![ident("delay")]))],
        ),
        function(
            "iscoroutine",
            vec![param("obj", None)],
            vec![ret(bool_lit(true))],
        ),
        function(
            "iscoroutinefunction",
            vec![param("obj", None)],
            vec![ret(bool_lit(true))],
        ),
        function(
            "sleep",
            vec![
                param("delay", Some(num(0.0))),
                param("result", Some(null())),
            ],
            vec![ret(ident("result"))],
        ),
        function("gather", vec![rest_param("aws")], vec![ret(ident("aws"))]),
    ]
}
