//! `tracemalloc` — lightweight allocation snapshot objects.
//!
//! Vybe does not observe VM heap allocations the way CPython does, but the
//! Python API surface is object-shaped: tracing state, snapshots, statistics,
//! filters and traceback/frame records. This module declares those classes as
//! AST and keeps the state in module globals.

use super::builders::*;
use vybe_ast::{ClassMember, Expression, Modifiers, ObjectProperty, Statement, StmtKind};

fn i(n: i64) -> Expression {
    Expression::int(n)
}

fn stat_list() -> Expression {
    list_of(vec![trace_stat_object(i(1024), i(1))])
}

fn frame_object() -> Expression {
    obj(vec![
        ("filename", str_lit("<vybe>")),
        ("lineno", i(1)),
    ])
}

fn traceback_list() -> Expression {
    list_of(vec![frame_object()])
}

fn trace_stat_object(size: Expression, count: Expression) -> Expression {
    obj(vec![
        ("size", size),
        ("count", count),
        ("traceback", traceback_list()),
    ])
}

fn trace_stat_diff_object(size_diff: Expression, count_diff: Expression) -> Expression {
    obj(vec![
        ("size", size_diff.clone()),
        ("count", count_diff.clone()),
        ("size_diff", size_diff),
        ("count_diff", count_diff),
        ("traceback", traceback_list()),
    ])
}

fn obj(items: Vec<(&str, Expression)>) -> Expression {
    Expression::new(vybe_ast::ExprKind::Object(
        items
            .into_iter()
            .map(|(key, value)| ObjectProperty::KeyValue {
                key: str_lit(key),
                value,
            })
            .collect(),
    ))
}

pub(super) fn frame() -> Statement {
    class(
        "__PyTraceFrame",
        vec![init(
            vec![
                param("filename", Some(str_lit("<vybe>"))),
                param("lineno", Some(i(1))),
            ],
            vec![
                set_this("filename", ident("filename")),
                set_this("lineno", ident("lineno")),
            ],
        )],
    )
}

pub(super) fn trace() -> Statement {
    class(
        "__PyTraceback",
        vec![
            init(
                vec![param(
                    "frames",
                    Some(list_of(vec![new("__PyTraceFrame", vec![])])),
                )],
                vec![set_this("_frames", ident("frames"))],
            ),
            method(
                "__len__",
                vec![],
                vec![ret(call_global("len", vec![this_field("_frames")]))],
            ),
            method(
                "__getitem__",
                vec![param("index", None)],
                vec![ret(index(this_field("_frames"), ident("index")))],
            ),
            method("__iter__", vec![], vec![ret(this_field("_frames"))]),
        ],
    )
}

pub(super) fn statistic() -> Statement {
    class(
        "__PyTraceStat",
        vec![
            init(
                vec![
                    param("size", Some(i(1024))),
                    param("count", Some(i(1))),
                    param("traceback", Some(new("__PyTraceback", vec![]))),
                ],
                vec![
                    set_this("size", ident("size")),
                    set_this("count", ident("count")),
                    set_this("traceback", ident("traceback")),
                ],
            ),
            method(
                "__str__",
                vec![],
                vec![ret(add(str_lit("<Statistic size="), call_global("str", vec![this_field("size")])) )],
            ),
        ],
    )
}

pub(super) fn statistic_diff() -> Statement {
    class(
        "__PyTraceStatDiff",
        vec![
            init(
                vec![
                    param("size_diff", Some(i(1024))),
                    param("count_diff", Some(i(1))),
                    param("traceback", Some(new("__PyTraceback", vec![]))),
                ],
                vec![
                    set_this("size", ident("size_diff")),
                    set_this("count", ident("count_diff")),
                    set_this("size_diff", ident("size_diff")),
                    set_this("count_diff", ident("count_diff")),
                    set_this("traceback", ident("traceback")),
                ],
            ),
            method(
                "__str__",
                vec![],
                vec![ret(add(
                    str_lit("<StatisticDiff size_diff="),
                    call_global("str", vec![this_field("size_diff")]),
                ))],
            ),
        ],
    )
}

pub(super) fn filter() -> Statement {
    class(
        "Filter",
        vec![init(
            vec![
                param("inclusive", Some(bool_lit(true))),
                param("filename_pattern", Some(str_lit("*"))),
                param("lineno", Some(null())),
                param("all_frames", Some(bool_lit(false))),
                param("domain", Some(null())),
            ],
            vec![
                set_this("inclusive", ident("inclusive")),
                set_this("filename_pattern", ident("filename_pattern")),
                set_this("lineno", ident("lineno")),
                set_this("all_frames", ident("all_frames")),
                set_this("domain", ident("domain")),
            ],
        )],
    )
}

pub(super) fn snapshot() -> Statement {
    class(
        "Snapshot",
        vec![
            init(vec![], vec![set_this("_stats", stat_list())]),
            ClassMember::Method(Box::new(Statement::with_span(
                StmtKind::FunctionDecl {
                    name: "load".to_string(),
                    params: vec![param("filename", None)],
                    body: vec![ret(new("Snapshot", vec![]))],
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
            ))),
            method(
                "statistics",
                vec![param("key_type", Some(str_lit("lineno")))],
                vec![ret(call_global("list", vec![this_field("_stats")]))],
            ),
            method(
                "compare_to",
                vec![
                    param("old_snapshot", None),
                    param("key_type", Some(str_lit("lineno"))),
                ],
                vec![ret(list_of(vec![trace_stat_diff_object(i(1024), i(1))]))],
            ),
            method(
                "filter_traces",
                vec![param("filters", Some(list_of(vec![])))],
                vec![ret(ident("self"))],
            ),
            method(
                "dump",
                vec![param("filename", None)],
                vec![
                    expr_stmt(call_global(
                        "__py_fs_write_text",
                        vec![ident("filename"), str_lit("vybe-tracemalloc-snapshot")],
                    )),
                    ret(null()),
                ],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("__py_tracemalloc_tracing", bool_lit(false)),
        global_assign("__py_tracemalloc_peak", i(1024)),
        function(
            "start",
            vec![param("nframe", Some(i(1)))],
            vec![
                assign(ident("__py_tracemalloc_tracing"), bool_lit(true)),
                assign(ident("__py_tracemalloc_peak"), i(1024)),
            ],
        ),
        function(
            "stop",
            vec![],
            vec![assign(ident("__py_tracemalloc_tracing"), bool_lit(false))],
        ),
        function(
            "is_tracing",
            vec![],
            vec![ret(ident("__py_tracemalloc_tracing"))],
        ),
        function("take_snapshot", vec![], vec![ret(new("Snapshot", vec![]))]),
        function(
            "get_traced_memory",
            vec![],
            vec![ret(tuple_of(vec![i(512), ident("__py_tracemalloc_peak")]))],
        ),
        function("get_tracemalloc_memory", vec![], vec![ret(i(0))]),
        function(
            "reset_peak",
            vec![],
            vec![assign(ident("__py_tracemalloc_peak"), i(512))],
        ),
        function(
            "get_object_traceback",
            vec![param("obj", None)],
            vec![ret(traceback_list())],
        ),
    ]
}
