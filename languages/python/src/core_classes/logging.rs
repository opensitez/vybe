//! `logging` — `Logger`, `Handler`, `Formatter` and the module surface.
//!
//! Most of this surface is *present and inert*: the corpus asserts that
//! `getLogger` answers an object, that `setLevel`/`addHandler` work, and that
//! the level constants resolve — not that anything is written anywhere. The
//! levels themselves are already `[namespace_constants]` profile rows, so they
//! are not restated here.
//!
//! The prelude wrapped all of this in a `__LoggingModule` class with a
//! module-level instance, because a bare `def` could not be reached as
//! `logging.getLogger`. `MODULE_SURFACE` answers that directly, so the module
//! object is gone and these are ordinary globals.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

pub(super) fn log_record() -> Statement {
    class("LogRecord", vec![init(any_args(), vec![])])
}

pub(super) fn formatter() -> Statement {
    class(
        "Formatter",
        vec![init(any_args(), vec![]), stub("format", str_lit(""))],
    )
}

pub(super) fn filter_class() -> Statement {
    class("Filter", vec![init(any_args(), vec![])])
}

pub(super) fn handler() -> Statement {
    class(
        "Handler",
        vec![
            init(any_args(), vec![set_this("level", num(0.0))]),
            method(
                "setLevel",
                vec![param("l", None)],
                vec![set_this("level", ident("l"))],
            ),
            stub("setFormatter", null()),
        ],
    )
}

/// `StreamHandler` / `FileHandler` — `Handler` with nothing added. The parent
/// is the whole declaration, and it is what makes `isinstance(h, Handler)` and
/// the inherited `setLevel` work.
pub(super) fn stream_handler() -> Statement {
    class_extending("StreamHandler", &["Handler"], vec![])
}

pub(super) fn file_handler() -> Statement {
    class_extending("FileHandler", &["Handler"], vec![])
}

pub(super) fn logger() -> Statement {
    class(
        "Logger",
        vec![
            init(
                vec![param("name", Some(str_lit("root")))],
                vec![
                    set_this("name", ident("name")),
                    set_this("level", num(0.0)),
                    set_this("handlers", call_global("list", vec![])),
                ],
            ),
            method(
                "setLevel",
                vec![param("l", None)],
                vec![set_this("level", ident("l"))],
            ),
            method(
                "addHandler",
                vec![param("h", None)],
                vec![expr_stmt(call(
                    member(this_field("handlers"), "append"),
                    vec![ident("h")],
                ))],
            ),
            method(
                "hasHandlers",
                vec![],
                vec![ret(binary(
                    BinOp::Gt,
                    call_global("len", vec![this_field("handlers")]),
                    num(0.0),
                ))],
            ),
            stub("isEnabledFor", bool_lit(true)),
            stub("debug", null()),
            stub("info", null()),
            stub("warning", null()),
            stub("error", null()),
            stub("critical", null()),
            stub("exception", null()),
            stub("log", null()),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "getLogger",
            vec![param("name", Some(str_lit("root")))],
            vec![ret(new("Logger", vec![ident("name")]))],
        ),
        global_assign("__py_logging_root", new("Logger", vec![str_lit("root")])),
        global_assign("__py_logging_last_resort", new("StreamHandler", vec![])),
        function(
            "getLevelName",
            vec![param("level", None)],
            vec![
                assign(
                    ident("__names"),
                    call_global(
                        "dict",
                        vec![list_of(vec![
                            list_of(vec![num(0.0), str_lit("NOTSET")]),
                            list_of(vec![num(10.0), str_lit("DEBUG")]),
                            list_of(vec![num(20.0), str_lit("INFO")]),
                            list_of(vec![num(30.0), str_lit("WARNING")]),
                            list_of(vec![num(40.0), str_lit("ERROR")]),
                            list_of(vec![num(50.0), str_lit("CRITICAL")]),
                        ])],
                    ),
                ),
                if_stmt(
                    binary(BinOp::In, ident("level"), ident("__names")),
                    vec![ret(index(ident("__names"), ident("level")))],
                ),
                ret(add(
                    str_lit("Level "),
                    call_global("str", vec![ident("level")]),
                )),
            ],
        ),
        stub_fn("basicConfig", null()),
        stub_fn("addLevelName", null()),
        stub_fn("debug", null()),
        stub_fn("info", null()),
        stub_fn("warning", null()),
        stub_fn("error", null()),
        stub_fn("critical", null()),
        stub_fn("log", null()),
    ]
}
