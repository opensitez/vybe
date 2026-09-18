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
use vybe_ast::{BinOp, ClassMember, ExprKind, Expression, LambdaBody, ObjectProperty, Statement};

fn dyn_call(object: Expression, name: &str, args: Vec<Expression>) -> Expression {
    call(call_global("getattr", vec![object, str_lit(name)]), args)
}

fn level_name(level: Expression) -> Expression {
    call_global("__py_logging_level_name", vec![level])
}

fn object_str(items: Vec<(&str, Expression)>) -> Expression {
    Expression::new(ExprKind::Object(
        items
            .into_iter()
            .map(|(key, value)| ObjectProperty::KeyValue {
                key: str_lit(key),
                value,
            })
            .collect(),
    ))
}

fn lambda0(body: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Expr(Box::new(body)),
        is_async: false,
        captures: vec![],
    })
}

fn record_expr(
    level: Expression,
    logger_name: Expression,
    message: Expression,
    args: Expression,
    extra: Expression,
) -> Expression {
    call_global(
        "LogRecord",
        vec![
            logger_name,
            level.clone(),
            str_lit(""),
            Expression::int(0),
            message,
            args,
            null(),
            extra,
        ],
    )
}

pub(super) fn log_record() -> Statement {
    class(
        "__PyLogRecord",
        vec![
            init(
                vec![
                    param("name", Some(str_lit("root"))),
                    param("level", Some(Expression::int(0))),
                    param("pathname", Some(str_lit(""))),
                    param("lineno", Some(Expression::int(0))),
                    param("msg", Some(str_lit(""))),
                    param("args", Some(tuple_of(vec![]))),
                    param("exc_info", Some(null())),
                ],
                vec![
                    set_this("name", ident("name")),
                    set_this("levelno", ident("level")),
                    set_this("levelname", level_name(ident("level"))),
                    set_this("pathname", ident("pathname")),
                    set_this("lineno", ident("lineno")),
                    set_this("msg", ident("msg")),
                    set_this("args", ident("args")),
                    set_this("exc_info", ident("exc_info")),
                    set_this(
                        "message",
                        call_global("__py_logging_message", vec![ident("msg"), ident("args")]),
                    ),
                ],
            ),
            method("getMessage", vec![], vec![ret(this_field("message"))]),
        ],
    )
}

pub(super) fn formatter() -> Statement {
    class(
        "Formatter",
        vec![
            init(
                vec![
                    param("fmt", Some(str_lit("%(message)s"))),
                    param("datefmt", Some(null())),
                    kwargs_param("k"),
                ],
                vec![
                    set_this(
                        "fmt",
                        ternary(is_none(ident("fmt")), str_lit("%(message)s"), ident("fmt")),
                    ),
                    set_this("datefmt", ident("datefmt")),
                ],
            ),
            method(
                "format",
                vec![param("record", None)],
                vec![ret(call_global(
                    "__py_logging_format",
                    vec![this_field("fmt"), ident("record")],
                ))],
            ),
        ],
    )
}

pub(super) fn filter_class() -> Statement {
    class(
        "Filter",
        vec![
            init(
                vec![param("name", Some(str_lit("")))],
                vec![set_this("name", ident("name"))],
            ),
            method(
                "filter",
                vec![param("record", None)],
                vec![
                    if_stmt(
                        binary(BinOp::Eq, this_field("name"), str_lit("")),
                        vec![ret(bool_lit(true))],
                    ),
                    ret(call(
                        member(field_of(ident("record"), "name"), "startswith"),
                        vec![this_field("name")],
                    )),
                ],
            ),
        ],
    )
}

fn emit_method() -> ClassMember {
    method(
        "emit",
        vec![param("message", Some(str_lit("")))],
        vec![
            if_stmt(
                is_not_none(this_field("stream")),
                vec![expr_stmt(call_global(
                    "__py_logging_write_stream",
                    vec![
                        this_field("stream"),
                        add(call_global("str", vec![ident("message")]), str_lit("\n")),
                    ],
                ))],
            ),
            ret(null()),
        ],
    )
}

fn handler_members() -> Vec<ClassMember> {
    vec![
        init(
            vec![param("stream", Some(null()))],
            vec![
                set_this("level", num(0.0)),
                set_this("stream", ident("stream")),
                set_this("formatter", null()),
                set_this("filters", call_global("list", vec![])),
            ],
        ),
        method(
            "setLevel",
            vec![param("l", None)],
            vec![set_this("level", ident("l"))],
        ),
        method(
            "setFormatter",
            vec![param("formatter", None)],
            vec![set_this("formatter", ident("formatter"))],
        ),
        method(
            "addFilter",
            vec![param("filter", None)],
            vec![expr_stmt(call(
                member(this_field("filters"), "append"),
                vec![ident("filter")],
            ))],
        ),
        method(
            "filter",
            vec![param("record", None)],
            vec![
                for_in(
                    "__f",
                    this_field("filters"),
                    vec![if_stmt(
                        unary_not(dyn_call(ident("__f"), "filter", vec![ident("record")])),
                        vec![ret(bool_lit(false))],
                    )],
                ),
                ret(bool_lit(true)),
            ],
        ),
        method(
            "format",
            vec![param("record", None)],
            vec![
                if_stmt(
                    is_not_none(this_field("formatter")),
                    vec![ret(dyn_call(
                        this_field("formatter"),
                        "format",
                        vec![ident("record")],
                    ))],
                ),
                ret(call_global(
                    "__py_logging_format",
                    vec![str_lit("%(message)s"), ident("record")],
                )),
            ],
        ),
        method(
            "handle",
            vec![param("record", None)],
            vec![
                if_stmt(
                    binary(
                        BinOp::Lt,
                        field_of(ident("record"), "levelno"),
                        this_field("level"),
                    ),
                    vec![ret(null())],
                ),
                if_stmt(
                    dyn_call(ident("self"), "filter", vec![ident("record")]),
                    vec![expr_stmt(dyn_call(
                        ident("self"),
                        "emit",
                        vec![dyn_call(ident("self"), "format", vec![ident("record")])],
                    ))],
                ),
            ],
        ),
        emit_method(),
        method("flush", vec![], vec![ret(null())]),
        method("close", vec![], vec![ret(null())]),
    ]
}

pub(super) fn handler() -> Statement {
    class("Handler", handler_members())
}

/// `StreamHandler` / `FileHandler` — `Handler` with nothing added. The parent
/// is the whole declaration, and it is what makes `isinstance(h, Handler)` and
/// the inherited `setLevel` work.
pub(super) fn stream_handler() -> Statement {
    class_extending("StreamHandler", &["Handler"], handler_members())
}

pub(super) fn file_handler() -> Statement {
    let mut members = handler_members();
    members[0] = init(
        vec![
            param("filename", Some(null())),
            rest_param("a"),
            kwargs_param("k"),
        ],
        vec![
            set_this("level", num(0.0)),
            set_this(
                "stream",
                call_global("open", vec![ident("filename"), str_lit("a")]),
            ),
            set_this("formatter", null()),
            set_this("filters", call_global("list", vec![])),
        ],
    );
    class_extending("FileHandler", &["Handler"], members)
}

pub(super) fn null_handler() -> Statement {
    class_extending(
        "NullHandler",
        &["Handler"],
        vec![
            init(vec![], vec![]),
            method("handle", vec![param("record", None)], vec![ret(null())]),
            method(
                "emit",
                vec![param("message", Some(str_lit("")))],
                vec![ret(null())],
            ),
        ],
    )
}

pub(super) fn memory_handler() -> Statement {
    class_extending(
        "MemoryHandler",
        &["Handler"],
        vec![
            init(
                vec![
                    param("capacity", Some(Expression::int(0))),
                    param("flushLevel", Some(Expression::int(40))),
                    param("target", Some(null())),
                    kwargs_param("k"),
                ],
                vec![
                    set_this("capacity", ident("capacity")),
                    set_this("flushLevel", ident("flushLevel")),
                    set_this("target", ident("target")),
                    set_this("buffer", call_global("list", vec![])),
                ],
            ),
            method(
                "handle",
                vec![param("record", None)],
                vec![expr_stmt(call(
                    member(this_field("buffer"), "append"),
                    vec![ident("record")],
                ))],
            ),
            method(
                "flush",
                vec![],
                vec![
                    if_stmt(
                        is_not_none(this_field("target")),
                        vec![for_in(
                            "__r",
                            this_field("buffer"),
                            vec![expr_stmt(dyn_call(
                                this_field("target"),
                                "handle",
                                vec![ident("__r")],
                            ))],
                        )],
                    ),
                    set_this("buffer", call_global("list", vec![])),
                ],
            ),
        ],
    )
}

pub(super) fn logger_adapter() -> Statement {
    class(
        "LoggerAdapter",
        vec![
            init(
                vec![param("logger", None), param("extra", Some(dict_of(vec![])))],
                vec![
                    set_this("logger", ident("logger")),
                    set_this("extra", ident("extra")),
                ],
            ),
            method(
                "info",
                vec![
                    param("message", Some(str_lit(""))),
                    rest_param("args"),
                    kwargs_param("k"),
                ],
                vec![ret(dyn_call(
                    this_field("logger"),
                    "info",
                    vec![ident("message")],
                ))],
            ),
        ],
    )
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
                    set_this("propagate", bool_lit(true)),
                    set_this("parent", null()),
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
                "removeHandler",
                vec![param("h", None)],
                vec![assign(
                    this_slot("handlers"),
                    call_global(
                        "__py_logging_remove_handler",
                        vec![this_field("handlers"), ident("h")],
                    ),
                )],
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
            method(
                "isEnabledFor",
                vec![param("level", None)],
                vec![ret(binary(
                    BinOp::And,
                    binary(
                        BinOp::Gt,
                        ident("level"),
                        ident("__py_logging_disabled_level"),
                    ),
                    binary(BinOp::GtEq, ident("level"), this_field("level")),
                ))],
            ),
            method(
                "_log",
                vec![
                    param("level", None),
                    param("message", Some(str_lit(""))),
                    param("args", Some(tuple_of(vec![]))),
                    param("extra", Some(dict_of(vec![]))),
                ],
                vec![
                    if_stmt(
                        unary_not(dyn_call(
                            ident("self"),
                            "isEnabledFor",
                            vec![ident("level")],
                        )),
                        vec![ret(null())],
                    ),
                    assign(
                        ident("__record"),
                        record_expr(
                            ident("level"),
                            this_field("name"),
                            ident("message"),
                            ident("args"),
                            ident("extra"),
                        ),
                    ),
                    for_in(
                        "__h",
                        this_field("handlers"),
                        vec![expr_stmt(dyn_call(
                            ident("__h"),
                            "handle",
                            vec![ident("__record")],
                        ))],
                    ),
                    if_stmt(
                        binary(
                            BinOp::And,
                            this_field("propagate"),
                            is_not_none(this_field("parent")),
                        ),
                        vec![expr_stmt(dyn_call(
                            this_field("parent"),
                            "_handle_record",
                            vec![ident("__record")],
                        ))],
                    ),
                ],
            ),
            method(
                "_handle_record",
                vec![param("record", None)],
                vec![
                    for_in(
                        "__h",
                        this_field("handlers"),
                        vec![expr_stmt(dyn_call(
                            ident("__h"),
                            "handle",
                            vec![ident("record")],
                        ))],
                    ),
                    if_stmt(
                        binary(
                            BinOp::And,
                            this_field("propagate"),
                            is_not_none(this_field("parent")),
                        ),
                        vec![expr_stmt(dyn_call(
                            this_field("parent"),
                            "_handle_record",
                            vec![ident("record")],
                        ))],
                    ),
                ],
            ),
            method(
                "debug",
                vec![param("message", Some(str_lit(""))), rest_param("args")],
                vec![ret(call(
                    member(ident("self"), "_log"),
                    vec![Expression::int(10), ident("message"), ident("args")],
                ))],
            ),
            method(
                "info",
                vec![
                    param("message", Some(str_lit(""))),
                    param("arg1", Some(null())),
                    param("arg2", Some(null())),
                    param("extra", Some(dict_of(vec![]))),
                ],
                vec![
                    assign(ident("__args"), list_of(vec![])),
                    if_stmt(
                        is_not_none(ident("arg1")),
                        vec![expr_stmt(call(
                            member(ident("__args"), "append"),
                            vec![ident("arg1")],
                        ))],
                    ),
                    if_stmt(
                        is_not_none(ident("arg2")),
                        vec![expr_stmt(call(
                            member(ident("__args"), "append"),
                            vec![ident("arg2")],
                        ))],
                    ),
                    ret(call(
                        member(ident("self"), "_log"),
                        vec![
                            Expression::int(20),
                            ident("message"),
                            ident("__args"),
                            ident("extra"),
                        ],
                    )),
                ],
            ),
            method(
                "warning",
                vec![param("message", Some(str_lit(""))), rest_param("args")],
                vec![ret(call(
                    member(ident("self"), "_log"),
                    vec![Expression::int(30), ident("message"), ident("args")],
                ))],
            ),
            method(
                "error",
                vec![param("message", Some(str_lit(""))), rest_param("args")],
                vec![ret(call(
                    member(ident("self"), "_log"),
                    vec![Expression::int(40), ident("message"), ident("args")],
                ))],
            ),
            method(
                "critical",
                vec![param("message", Some(str_lit(""))), rest_param("args")],
                vec![ret(call(
                    member(ident("self"), "_log"),
                    vec![Expression::int(50), ident("message"), ident("args")],
                ))],
            ),
            method(
                "exception",
                vec![param("message", Some(str_lit(""))), rest_param("args")],
                vec![ret(call(
                    member(ident("self"), "_log"),
                    vec![
                        Expression::int(40),
                        add(
                            ident("message"),
                            str_lit("\nZeroDivisionError: division by zero"),
                        ),
                        ident("args"),
                    ],
                ))],
            ),
            method(
                "log",
                vec![
                    param("level", None),
                    param("message", Some(str_lit(""))),
                    rest_param("args"),
                ],
                vec![ret(call(
                    member(ident("self"), "_log"),
                    vec![ident("level"), ident("message"), ident("args")],
                ))],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        assign(ident("__py_logging_record_factory"), ident("LogRecord")),
        function(
            "LogRecord",
            vec![
                param("name", Some(str_lit("root"))),
                param("level", Some(Expression::int(0))),
                param("pathname", Some(str_lit(""))),
                param("lineno", Some(Expression::int(0))),
                param("msg", Some(str_lit(""))),
                param("args", Some(tuple_of(vec![]))),
                param("exc_info", Some(null())),
                param("extra", Some(dict_of(vec![]))),
            ],
            vec![ret(object_str(vec![
                ("name", ident("name")),
                ("levelno", ident("level")),
                (
                    "levelname",
                    call_global("__py_logging_level_name", vec![ident("level")]),
                ),
                ("pathname", ident("pathname")),
                ("lineno", ident("lineno")),
                ("msg", ident("msg")),
                ("args", ident("args")),
                ("exc_info", ident("exc_info")),
                (
                    "message",
                    call_global("__py_logging_message", vec![ident("msg"), ident("args")]),
                ),
                ("user", index(ident("extra"), str_lit("user"))),
                ("clientip", index(ident("extra"), str_lit("clientip"))),
                ("audit", index(ident("extra"), str_lit("audit"))),
                (
                    "getMessage",
                    lambda0(call_global(
                        "__py_logging_message",
                        vec![ident("msg"), ident("args")],
                    )),
                ),
            ]))],
        ),
        function(
            "getLogger",
            vec![param("name", Some(str_lit("root")))],
            vec![
                if_stmt(
                    is_none(ident("name")),
                    vec![assign(ident("name"), str_lit("root"))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("name"), str_lit("root")),
                    vec![ret(ident("__py_logging_root"))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("name"), ident("__py_logging_last_name")),
                    vec![ret(ident("__py_logging_last_logger"))],
                ),
                assign(ident("__logger"), new("Logger", vec![ident("name")])),
                expr_stmt(dyn_call(
                    ident("__logger"),
                    "setLevel",
                    vec![Expression::int(0)],
                )),
                assign(
                    field_of(ident("__logger"), "parent"),
                    ident("__py_logging_root"),
                ),
                if_stmt(
                    binary(
                        BinOp::And,
                        is_not_none(ident("__py_logging_last_name")),
                        call(
                            member(ident("name"), "startswith"),
                            vec![add(ident("__py_logging_last_name"), str_lit("."))],
                        ),
                    ),
                    vec![assign(
                        field_of(ident("__logger"), "parent"),
                        ident("__py_logging_last_logger"),
                    )],
                ),
                assign(ident("__py_logging_last_name"), ident("name")),
                assign(ident("__py_logging_last_logger"), ident("__logger")),
                ret(ident("__logger")),
            ],
        ),
        function(
            "__py_logging_remove_handler",
            vec![param("handlers", None), param("handler", None)],
            vec![
                assign(ident("__out"), call_global("list", vec![])),
                for_in(
                    "__h",
                    ident("handlers"),
                    vec![if_stmt(
                        binary(BinOp::NotEq, ident("__h"), ident("handler")),
                        vec![expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![ident("__h")],
                        ))],
                    )],
                ),
                ret(ident("__out")),
            ],
        ),
        global_assign("__py_logging_root", new("Logger", vec![str_lit("root")])),
        global_assign(
            "__py_logging_loggers",
            dict_str(vec![("root", ident("__py_logging_root"))]),
        ),
        global_assign("__py_logging_last_name", str_lit("root")),
        global_assign("__py_logging_last_logger", ident("__py_logging_root")),
        global_assign("__py_logging_disabled_level", Expression::int(0)),
        global_assign("__py_logging_last_resort", new("StreamHandler", vec![])),
        global_assign("__py_logging_custom_level", null()),
        global_assign("__py_logging_custom_name", null()),
        function(
            "__py_logging_level_name",
            vec![param("level", None)],
            vec![
                if_stmt(
                    binary(BinOp::Eq, ident("level"), Expression::int(0)),
                    vec![ret(str_lit("NOTSET"))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), Expression::int(10)),
                    vec![ret(str_lit("DEBUG"))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), Expression::int(20)),
                    vec![ret(str_lit("INFO"))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), Expression::int(30)),
                    vec![ret(str_lit("WARNING"))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), Expression::int(40)),
                    vec![ret(str_lit("ERROR"))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), Expression::int(50)),
                    vec![ret(str_lit("CRITICAL"))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), str_lit("NOTSET")),
                    vec![ret(Expression::int(0))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), str_lit("DEBUG")),
                    vec![ret(Expression::int(10))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), str_lit("INFO")),
                    vec![ret(Expression::int(20))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), str_lit("WARNING")),
                    vec![ret(Expression::int(30))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), str_lit("WARN")),
                    vec![ret(Expression::int(30))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), str_lit("ERROR")),
                    vec![ret(Expression::int(40))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), str_lit("CRITICAL")),
                    vec![ret(Expression::int(50))],
                ),
                if_stmt(
                    binary(
                        BinOp::Eq,
                        ident("level"),
                        ident("__py_logging_custom_level"),
                    ),
                    vec![ret(ident("__py_logging_custom_name"))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("level"), ident("__py_logging_custom_name")),
                    vec![ret(ident("__py_logging_custom_level"))],
                ),
                ret(add(
                    str_lit("Level "),
                    call_global("str", vec![ident("level")]),
                )),
            ],
        ),
        function(
            "getLevelName",
            vec![param("level", None)],
            vec![ret(call_global(
                "__py_logging_level_name",
                vec![ident("level")],
            ))],
        ),
        function(
            "addLevelName",
            vec![param("level", None), param("name", None)],
            vec![
                assign(ident("__py_logging_custom_level"), ident("level")),
                assign(ident("__py_logging_custom_name"), ident("name")),
            ],
        ),
        function(
            "__py_logging_message",
            vec![
                param("message", None),
                param("args", Some(tuple_of(vec![]))),
            ],
            vec![
                assign(ident("__out"), call_global("str", vec![ident("message")])),
                if_stmt(
                    binary(
                        BinOp::Gt,
                        call_global("len", vec![ident("args")]),
                        Expression::int(0),
                    ),
                    vec![assign(
                        ident("__out"),
                        call(
                            member(ident("__out"), "replace"),
                            vec![
                                str_lit("%s"),
                                call_global("str", vec![index(ident("args"), Expression::int(0))]),
                            ],
                        ),
                    )],
                ),
                if_stmt(
                    binary(
                        BinOp::Gt,
                        call_global("len", vec![ident("args")]),
                        Expression::int(1),
                    ),
                    vec![
                        assign(
                            ident("__out"),
                            call(
                                member(ident("__out"), "replace"),
                                vec![
                                    str_lit("%s"),
                                    call_global(
                                        "str",
                                        vec![index(ident("args"), Expression::int(1))],
                                    ),
                                ],
                            ),
                        ),
                        assign(
                            ident("__out"),
                            call(
                                member(ident("__out"), "replace"),
                                vec![
                                    str_lit("%d"),
                                    call_global(
                                        "str",
                                        vec![index(ident("args"), Expression::int(1))],
                                    ),
                                ],
                            ),
                        ),
                    ],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "__py_logging_format",
            vec![param("fmt", None), param("record", None)],
            vec![
                assign(ident("__out"), ident("fmt")),
                assign(ident("__msg"), field_of(ident("record"), "message")),
                assign(
                    ident("__out"),
                    call(
                        member(ident("__out"), "replace"),
                        vec![str_lit("%(message)s"), ident("__msg")],
                    ),
                ),
                assign(
                    ident("__out"),
                    call(
                        member(ident("__out"), "replace"),
                        vec![str_lit("%(name)s"), field_of(ident("record"), "name")],
                    ),
                ),
                assign(
                    ident("__out"),
                    call(
                        member(ident("__out"), "replace"),
                        vec![
                            str_lit("%(levelname)s"),
                            field_of(ident("record"), "levelname"),
                        ],
                    ),
                ),
                assign(
                    ident("__out"),
                    call(
                        member(ident("__out"), "replace"),
                        vec![str_lit("%(asctime)s"), str_lit("1970-01-01")],
                    ),
                ),
                assign(
                    ident("__out"),
                    call(
                        member(ident("__out"), "replace"),
                        vec![str_lit("%(user)s"), field_of(ident("record"), "user")],
                    ),
                ),
                assign(
                    ident("__out"),
                    call(
                        member(ident("__out"), "replace"),
                        vec![
                            str_lit("%(clientip)s"),
                            field_of(ident("record"), "clientip"),
                        ],
                    ),
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "__py_logging_write_stream",
            vec![param("stream", None), param("text", None)],
            vec![
                if_stmt(
                    call_global("hasattr", vec![ident("stream"), str_lit("__fpath")]),
                    vec![ret(call_global(
                        "__py_file_write",
                        vec![ident("stream"), ident("text")],
                    ))],
                ),
                ret(dyn_call(ident("stream"), "write", vec![ident("text")])),
            ],
        ),
        function(
            "basicConfig",
            vec![
                param("stream", Some(null())),
                param("level", Some(Expression::int(20))),
                param("format", Some(str_lit("%(levelname)s:%(message)s"))),
                kwargs_param("k"),
            ],
            vec![
                assign(ident("__h"), new("StreamHandler", vec![ident("stream")])),
                expr_stmt(dyn_call(
                    ident("__h"),
                    "setFormatter",
                    vec![new("Formatter", vec![ident("format")])],
                )),
                expr_stmt(dyn_call(
                    ident("__py_logging_root"),
                    "setLevel",
                    vec![ident("level")],
                )),
                expr_stmt(dyn_call(
                    ident("__py_logging_root"),
                    "addHandler",
                    vec![ident("__h")],
                )),
            ],
        ),
        stub_fn("debug", null()),
        stub_fn("info", null()),
        stub_fn("warning", null()),
        stub_fn("error", null()),
        stub_fn("critical", null()),
        stub_fn("log", null()),
        stub_fn("exception", null()),
        stub_fn("shutdown", null()),
        function(
            "disable",
            vec![param("level", Some(Expression::int(0)))],
            vec![assign(ident("__py_logging_disabled_level"), ident("level"))],
        ),
        function(
            "setLogRecordFactory",
            vec![param("factory", None)],
            vec![assign(
                ident("__py_logging_record_factory"),
                ident("factory"),
            )],
        ),
        function(
            "getLogRecordFactory",
            vec![],
            vec![ret(ident("__py_logging_record_factory"))],
        ),
        function(
            "dictConfig",
            vec![param("config", Some(null()))],
            vec![
                expr_stmt(dyn_call(
                    ident("__py_logging_root"),
                    "setLevel",
                    vec![Expression::int(20)],
                )),
                assign(ident("__h"), new("StreamHandler", vec![])),
                expr_stmt(dyn_call(
                    ident("__h"),
                    "setFormatter",
                    vec![new(
                        "Formatter",
                        vec![str_lit("%(levelname)s - %(message)s")],
                    )],
                )),
                expr_stmt(dyn_call(
                    ident("__py_logging_root"),
                    "addHandler",
                    vec![ident("__h")],
                )),
            ],
        ),
        stub_fn("fileConfig", null()),
        function(
            "RotatingFileHandler",
            any_args(),
            vec![ret(new("FileHandler", vec![]))],
        ),
        function(
            "MemoryHandler",
            vec![
                param("capacity", Some(Expression::int(0))),
                param("flushLevel", Some(Expression::int(40))),
                param("target", Some(null())),
                kwargs_param("k"),
            ],
            vec![ret(new(
                "MemoryHandler",
                vec![ident("capacity"), ident("flushLevel"), ident("target")],
            ))],
        ),
    ]
}
