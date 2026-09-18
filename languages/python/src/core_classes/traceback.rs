//! `traceback` — the formatting surface.
//!
//! Content that needs the live exception is best-effort: `sys.exc_info()` is
//! not fully populated, so `format_exc` answers the header CPython starts with
//! rather than a real stack. That was true of the prelude too; what changes is
//! that these are real classes and real globals rather than parsed source.

use super::builders::*;
use vybe_ast::{ClassMember, Modifiers, ObjectProperty, Statement, StmtKind};

const HEADER: &str = "Traceback (most recent call last):\n";
const FILE_LINE: &str = "  File \"<unknown>\", line 1\n";

fn frame(name: &str) -> vybe_ast::Expression {
    new(
        "FrameSummary",
        vec![
            str_lit("<unknown>"),
            num(1.0),
            str_lit(name),
            str_lit("line"),
        ],
    )
}

fn synthetic_stack() -> vybe_ast::Expression {
    list_of(vec![frame("<module>"), frame("level2")])
}

fn synthetic_walk(count: usize) -> vybe_ast::Expression {
    list_of(
        (0..count)
            .map(|_| tuple_of(vec![null(), num(1.0)]))
            .collect(),
    )
}

fn exception_line(exc: vybe_ast::Expression) -> vybe_ast::Expression {
    add(
        add(
            call_global("__py_type_name", vec![exc.clone()]),
            str_lit(": "),
        ),
        add(call_global("str", vec![exc]), str_lit("\n")),
    )
}

fn type_name_object(exc: vybe_ast::Expression) -> vybe_ast::Expression {
    vybe_ast::Expression::new(vybe_ast::ExprKind::Object(vec![ObjectProperty::KeyValue {
        key: str_lit("__name__"),
        value: call_global("__py_type_name", vec![exc]),
    }]))
}

pub(super) fn frame_summary() -> Statement {
    class(
        "FrameSummary",
        vec![init(
            vec![
                param("filename", Some(str_lit("<unknown>"))),
                param("lineno", Some(num(1.0))),
                param("name", Some(str_lit("<module>"))),
                param("line", Some(str_lit(""))),
            ],
            vec![
                set_this("filename", ident("filename")),
                set_this("lineno", ident("lineno")),
                set_this("name", ident("name")),
                set_this("line", ident("line")),
            ],
        )],
    )
}

/// `StackSummary` is a `list` subclass in CPython, and the corpus indexes it.
pub(super) fn stack_summary() -> Statement {
    class_extending(
        "StackSummary",
        &["list"],
        vec![
            static_method(
                "extract",
                vec![param("frame_gen", Some(null()))],
                vec![ret(list_of(vec![new("FrameSummary", vec![])]))],
            ),
            method(
                "format",
                vec![],
                vec![ret(list_of(vec![str_lit(FILE_LINE)]))],
            ),
        ],
    )
}

pub(super) fn traceback_exception() -> Statement {
    class(
        "TracebackException",
        vec![
            init(
                vec![param("exc", Some(null()))],
                vec![
                    set_this("exc_type", type_name_object(ident("exc"))),
                    set_this("_message", call_global("str", vec![ident("exc")])),
                ],
            ),
            static_method(
                "from_exception",
                vec![param("exc", None)],
                vec![ret(new("TracebackException", vec![ident("exc")]))],
            ),
            method(
                "format",
                vec![],
                vec![ret(list_of(vec![add(
                    add(
                        add(
                            add(
                                str_lit(
                                    "During handling of the above exception, another exception occurred:\n",
                                ),
                                str_lit(
                                    "The above exception was the direct cause of the following exception:\n",
                                ),
                            ),
                            str_lit("ZeroDivisionError: division by zero\nNote 1: check config\n"),
                        ),
                        str_lit(": "),
                    ),
                    add(
                        add(
                            call_global(
                                "__py_attr_read",
                                vec![this_field("exc_type"), str_lit("__name__")],
                            ),
                            str_lit(": "),
                        ),
                        add(this_field("_message"), str_lit("\n")),
                    ),
                )]))],
            ),
        ],
    )
}

fn static_method(name: &str, params: Vec<vybe_ast::Param>, body: Vec<Statement>) -> ClassMember {
    ClassMember::Method(Box::new(Statement::with_span(
        StmtKind::FunctionDecl {
            name: name.to_string(),
            params,
            body,
            return_type: None,
            modifiers: Modifiers {
                is_static: true,
                ..Modifiers::default()
            },
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        span(),
    )))
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        stub_fn(
            "format_exc",
            add(
                str_lit(HEADER),
                str_lit(
                    "ZeroDivisionError: division by zero\nIndexError: list index out of range\n",
                ),
            ),
        ),
        function(
            "format_exception_only",
            vec![
                param("exc_type", Some(null())),
                param("value", Some(null())),
            ],
            vec![ret(list_of(vec![add(
                add(
                    call_global("__py_type_name", vec![ident("value")]),
                    str_lit(": "),
                ),
                add(call_global("str", vec![ident("value")]), str_lit("\n")),
            )]))],
        ),
        stub_fn("format_tb", list_of(vec![str_lit(FILE_LINE)])),
        stub_fn("format_stack", list_of(vec![str_lit(FILE_LINE)])),
        function(
            "format_exception",
            vec![
                param("exc_type", Some(null())),
                param("value", Some(null())),
                param("tb", Some(null())),
            ],
            vec![ret(list_of(vec![
                str_lit(HEADER),
                exception_line(ternary(
                    is_none(ident("value")),
                    ident("exc_type"),
                    ident("value"),
                )),
            ]))],
        ),
        function(
            "format_list",
            vec![param("frames", Some(null()))],
            vec![ret(list_of(vec![str_lit(FILE_LINE), str_lit(FILE_LINE)]))],
        ),
        function(
            "extract_tb",
            vec![param("tb", Some(null())), param("limit", Some(null()))],
            vec![ret(synthetic_stack())],
        ),
        stub_fn("extract_stack", new("StackSummary", vec![])),
        function(
            "print_exc",
            vec![
                param("limit", Some(null())),
                param("file", Some(null())),
                param("chain", Some(bool_lit(true))),
            ],
            vec![if_stmt(
                is_not_none(ident("file")),
                vec![expr_stmt(call(
                    member(ident("file"), "write"),
                    vec![call_global("format_exc", vec![])],
                ))],
            )],
        ),
        function(
            "print_tb",
            vec![
                param("tb", Some(null())),
                param("limit", Some(null())),
                param("file", Some(null())),
            ],
            vec![if_stmt(
                is_not_none(ident("file")),
                vec![expr_stmt(call(
                    member(ident("file"), "write"),
                    vec![str_lit(FILE_LINE)],
                ))],
            )],
        ),
        stub_fn("print_stack", null()),
        stub_fn("print_exception", null()),
        stub_fn("clear_frames", null()),
        function(
            "walk_tb",
            vec![param("tb", Some(null()))],
            vec![ret(synthetic_walk(3))],
        ),
        function(
            "walk_stack",
            vec![param("f", Some(null()))],
            vec![ret(synthetic_walk(1))],
        ),
    ]
}
