//! `contextlib` — the context managers.
//!
//! `__enter__` / `__exit__` are declared as ordinary dunders. python's
//! `protocol.rs` maps both onto their `ProtocolSlot`, so `with` binds through
//! the shared machinery and none of these classes carries anything
//! context-manager-specific.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

pub(super) fn null_context() -> Statement {
    class(
        "__NullContext",
        vec![
            init(
                vec![param("enter_result", Some(null()))],
                vec![set_this("enter_result", ident("enter_result"))],
            ),
            method("__enter__", vec![], vec![ret(this_field("enter_result"))]),
            method("__exit__", any_args(), vec![ret(bool_lit(false))]),
        ],
    )
}

pub(super) fn closing() -> Statement {
    class(
        "__Closing",
        vec![
            init(
                vec![param("thing", None)],
                vec![set_this("thing", ident("thing"))],
            ),
            method("__enter__", vec![], vec![ret(this_field("thing"))]),
            method(
                "__exit__",
                any_args(),
                vec![
                    expr_stmt(call(member(this_field("thing"), "close"), vec![])),
                    ret(bool_lit(false)),
                ],
            ),
        ],
    )
}

pub(super) fn suppress() -> Statement {
    class(
        "__Suppress",
        vec![
            init(
                vec![rest_param("exc")],
                vec![set_this("exc", ident("exc"))],
            ),
            method("__enter__", vec![], vec![ret(null())]),
            method(
                "__exit__",
                vec![param("exc_type", Some(null())), rest_param("a")],
                vec![
                    if_stmt(
                        is_none(ident("exc_type")),
                        vec![ret(bool_lit(false))],
                    ),
                    for_in(
                        "__e",
                        this_field("exc"),
                        vec![if_stmt(
                            call_global("issubclass", vec![ident("exc_type"), ident("__e")]),
                            vec![ret(bool_lit(true))],
                        )],
                    ),
                    ret(bool_lit(false)),
                ],
            ),
        ],
    )
}

/// The object `@contextmanager` produces around a generator: `__enter__` is the
/// first `next`, `__exit__` the second.
pub(super) fn gen_cm() -> Statement {
    class(
        "__GenCM",
        vec![
            init(vec![param("gen", None)], vec![set_this("gen", ident("gen"))]),
            method(
                "__enter__",
                vec![],
                vec![ret(call_global("next", vec![this_field("gen")]))],
            ),
            method(
                "__exit__",
                any_args(),
                vec![
                    try_except(
                        vec![expr_stmt(call_global("next", vec![this_field("gen")]))],
                        "StopIteration",
                        vec![],
                    ),
                    ret(bool_lit(false)),
                ],
            ),
        ],
    )
}

pub(super) fn redirect() -> Statement {
    class(
        "__Redirect",
        vec![
            init(
                vec![param("target", None)],
                vec![set_this("target", ident("target"))],
            ),
            method("__enter__", vec![], vec![ret(this_field("target"))]),
            method("__exit__", any_args(), vec![ret(bool_lit(false))]),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "nullcontext",
            vec![param("enter_result", Some(null()))],
            vec![ret(new("__NullContext", vec![ident("enter_result")]))],
        ),
        function(
            "closing",
            vec![param("thing", None)],
            vec![ret(new("__Closing", vec![ident("thing")]))],
        ),
        function(
            "redirect_stdout",
            vec![param("target", None)],
            vec![ret(new("__Redirect", vec![ident("target")]))],
        ),
        function(
            "redirect_stderr",
            vec![param("target", None)],
            vec![ret(new("__Redirect", vec![ident("target")]))],
        ),
        // ⛔ `suppress` and `contextmanager` are the two the corpus leans on —
        // dropping them when this module was ported cost 28 tests across
        // `diagnostics_runtime` and the `context_managers_*` suites.
        function(
            "suppress",
            vec![rest_param("exc")],
            vec![ret(new_spread("__Suppress", ident("exc")))],
        ),
        // `@contextmanager` wraps a GENERATOR function: calling the decorated
        // name runs the generator and hands back the `__GenCM` that drives it,
        // so the helper closes over `func` and forwards its own arguments.
        function(
            "contextmanager",
            vec![param("func", None)],
            vec![
                function(
                    "__cm_helper",
                    any_args(),
                    vec![ret(new(
                        "__GenCM",
                        vec![call_spread(ident("func"), ident("a"))],
                    ))],
                ),
                ret(ident("__cm_helper")),
            ],
        ),
    ]
}
