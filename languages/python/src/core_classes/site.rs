//! `site` — path setup helpers as an AST module surface.
//!
//! The behavior here is the small deterministic part of CPython's `site`
//! module that matters in Vybe: user-site paths, `sys.path` mutation and `.pth`
//! processing. It is declared directly as AST and uses the existing filesystem
//! and Python container primitives.

use super::builders::*;
use vybe_ast::{BinOp, ExprKind, Expression, ObjectProperty, Statement};

const USER_BASE: &str = "/home/vybe/.local";
const USER_SITE: &str = "/home/vybe/.local/lib/python3.12/site-packages";
const PREFIX: &str = "/usr";

type Expr = Expression;

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn not_contains(haystack: Expr, needle: Expr) -> Expr {
    unary_not(contains(haystack, needle))
}

fn and(left: Expr, right: Expr) -> Expr {
    binary(BinOp::And, left, right)
}

fn ne(left: Expr, right: Expr) -> Expr {
    binary(BinOp::NotEq, left, right)
}

fn eq(left: Expr, right: Expr) -> Expr {
    binary(BinOp::Eq, left, right)
}

fn slash_join(left: Expr, right: Expr) -> Expr {
    add(add(left, str_lit("/")), right)
}

fn str_splitlines(value: Expr) -> Expr {
    call_global("__py_str_splitlines", vec![value])
}

fn str_startswith(value: Expr, prefix: Expr) -> Expr {
    call_global("__py_str_startswith", vec![value, prefix])
}

fn str_endswith(value: Expr, suffix: Expr) -> Expr {
    call_global("__py_str_endswith", vec![value, suffix])
}

fn str_strip(value: Expr) -> Expr {
    call_global("__py_str_strip", vec![value])
}

fn sys_path() -> Expr {
    ident("__py_sys_path")
}

fn append(list: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(list, "append"), vec![value]))
}

fn add_to_set(set: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(set, "add"), vec![value]))
}

fn assign_module_attr(module: &str, field: &str, value: Expr) -> Statement {
    assign(
        Expression::with_span(
            ExprKind::Member {
                object: Box::new(ident(module)),
                field: field.to_string(),
                null_safe: false,
            },
            span(),
        ),
        value,
    )
}

fn helper_object(text: &str) -> Expr {
    Expression::with_span(
        ExprKind::Object(vec![ObjectProperty::KeyValue {
            key: str_lit("__str__"),
            value: Expression::with_span(
                ExprKind::Lambda {
                    params: vec![],
                    body: vybe_ast::LambdaBody::Expr(Box::new(str_lit(text))),
                    is_async: false,
                    captures: vec![],
                },
                span(),
            ),
        }]),
        span(),
    )
}

fn site_packages_for(prefix: Expr) -> Expr {
    slash_join(
        slash_join(prefix, str_lit("lib/python3.12")),
        str_lit("site-packages"),
    )
}

fn pth_process_line_body() -> Vec<Statement> {
    vec![
        assign(
            ident("__py_site_line"),
            str_strip(ident("__py_site_raw_line")),
        ),
        if_stmt(
            and(
                ne(ident("__py_site_line"), str_lit("")),
                unary_not(str_startswith(ident("__py_site_line"), str_lit("#"))),
            ),
            vec![
                if_stmt(
                    str_startswith(ident("__py_site_line"), str_lit("import sys;")),
                    vec![assign(ident("__py_sys_pth_test_var"), str_lit("executed"))],
                ),
                if_stmt(
                    unary_not(str_startswith(ident("__py_site_line"), str_lit("import "))),
                    vec![
                        assign(
                            ident("__py_site_path"),
                            slash_join(ident("sitedir"), ident("__py_site_line")),
                        ),
                        if_stmt(
                            and(
                                call_global("__py_fs_exists", vec![ident("__py_site_path")]),
                                not_contains(ident("known_paths"), ident("__py_site_path")),
                            ),
                            vec![
                                add_to_set(ident("known_paths"), ident("__py_site_path")),
                                if_stmt(
                                    not_contains(sys_path(), ident("__py_site_path")),
                                    vec![append(sys_path(), ident("__py_site_path"))],
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        ),
    ]
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("USER_BASE", str_lit(USER_BASE)),
        global_assign("USER_SITE", str_lit(USER_SITE)),
        global_assign("PREFIXES", list_of(vec![str_lit(PREFIX)])),
        function("getuserbase", vec![], vec![ret(str_lit(USER_BASE))]),
        function("getusersitepackages", vec![], vec![ret(str_lit(USER_SITE))]),
        function(
            "makepath",
            vec![rest_param("paths")],
            vec![
                assign(ident("__out"), str_lit("")),
                for_in(
                    "__part",
                    ident("paths"),
                    vec![if_stmt(
                        ne(ident("__part"), str_lit("")),
                        vec![assign(
                            ident("__out"),
                            ternary(
                                eq(ident("__out"), str_lit("")),
                                slash_join(str_lit(""), ident("__part")),
                                slash_join(ident("__out"), ident("__part")),
                            ),
                        )],
                    )],
                ),
                ret(tuple_of(vec![ident("__out"), ident("__out")])),
            ],
        ),
        function(
            "getsitepackages",
            vec![param("prefixes", Some(Expression::null()))],
            vec![
                assign(ident("__out"), list_of(vec![])),
                assign(
                    ident("__prefixes"),
                    ternary(
                        is_none(ident("prefixes")),
                        list_of(vec![str_lit(PREFIX)]),
                        ident("prefixes"),
                    ),
                ),
                for_in(
                    "__prefix",
                    ident("__prefixes"),
                    vec![append(ident("__out"), site_packages_for(ident("__prefix")))],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "addpackage",
            vec![
                param("sitedir", None),
                param("name", None),
                param("known_paths", Some(Expression::null())),
            ],
            vec![
                if_stmt(
                    is_none(ident("known_paths")),
                    vec![assign(ident("known_paths"), call_global("set", vec![]))],
                ),
                assign(ident("__pth"), slash_join(ident("sitedir"), ident("name"))),
                if_stmt(
                    unary_not(call_global("__py_fs_exists", vec![ident("__pth")])),
                    vec![ret(ident("known_paths"))],
                ),
                assign(
                    ident("__py_site_text"),
                    call_global("__py_fs_read_text", vec![ident("__pth")]),
                ),
                for_in(
                    "__py_site_raw_line",
                    str_splitlines(ident("__py_site_text")),
                    pth_process_line_body(),
                ),
                ret(ident("known_paths")),
            ],
        ),
        function(
            "addsitedir",
            vec![
                param("sitedir", None),
                param("known_paths", Some(Expression::null())),
            ],
            vec![
                if_stmt(
                    is_none(ident("known_paths")),
                    vec![assign(ident("known_paths"), call_global("set", vec![]))],
                ),
                if_stmt(
                    not_contains(sys_path(), ident("sitedir")),
                    vec![append(sys_path(), ident("sitedir"))],
                ),
                add_to_set(ident("known_paths"), ident("sitedir")),
                if_stmt(
                    call_global("__py_fs_exists", vec![ident("sitedir")]),
                    vec![for_in(
                        "__py_site_entry",
                        call_global("__py_fs_list_dir", vec![ident("sitedir")]),
                        vec![if_stmt(
                            str_endswith(ident("__py_site_entry"), str_lit(".pth")),
                            vec![expr_stmt(call_global(
                                "addpackage",
                                vec![
                                    ident("sitedir"),
                                    ident("__py_site_entry"),
                                    ident("known_paths"),
                                ],
                            ))],
                        )],
                    )],
                ),
                ret(ident("known_paths")),
            ],
        ),
        function(
            "sethelper",
            vec![],
            vec![
                assign_module_attr(
                    "builtins",
                    "help",
                    helper_object("Type help() for interactive help."),
                ),
                ret(null()),
            ],
        ),
        function(
            "setcopyright",
            vec![],
            vec![
                assign_module_attr(
                    "builtins",
                    "copyright",
                    helper_object("Copyright (c) Vybe Python."),
                ),
                ret(null()),
            ],
        ),
        function(
            "setquit",
            vec![],
            vec![
                assign_module_attr(
                    "builtins",
                    "quit",
                    helper_object("Use quit() or exit() to leave."),
                ),
                assign_module_attr(
                    "builtins",
                    "exit",
                    helper_object("Use quit() or exit() to leave."),
                ),
                ret(null()),
            ],
        ),
        function(
            "main",
            vec![],
            vec![
                expr_stmt(call_global(
                    "print",
                    vec![str_lit("sys.path = "), sys_path()],
                )),
                ret(null()),
            ],
        ),
    ]
}
