//! `linecache` — file line reads with the public module cache.
//!
//! The previous adapter intentionally skipped caching. This module declares the
//! cache as a real dict and routes reads through the existing filesystem
//! primitive so `linecache.cache`, `updatecache` and `checkcache` agree.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn cache() -> Expr {
    ident("__py_linecache_cache")
}

fn append(list: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(list, "append"), vec![value]))
}

fn call_clear_cache() -> Statement {
    expr_stmt(call(member(cache(), "clear"), vec![]))
}

fn make_lines_body() -> Vec<Statement> {
    vec![
        assign(ident("__lines"), list_of(vec![])),
        if_stmt(
            binary(BinOp::Eq, ident("text"), str_lit("")),
            vec![append(ident("__lines"), str_lit("\n"))],
        ),
        if_stmt(
            binary(BinOp::NotEq, ident("text"), str_lit("")),
            vec![for_in(
                "__line",
                call(member(ident("text"), "splitlines"), vec![]),
                vec![append(
                    ident("__lines"),
                    add(ident("__line"), str_lit("\n")),
                )],
            )],
        ),
        ret(ident("__lines")),
    ]
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("__py_linecache_cache", dict_str(vec![])),
        function(
            "__py_linecache_make_lines",
            vec![param("text", None)],
            make_lines_body(),
        ),
        function(
            "__py_linecache_updatecache",
            vec![
                param("filename", None),
                param("module_globals", Some(null())),
            ],
            vec![
                if_stmt(
                    unary_not(call_global("__py_fs_exists", vec![ident("filename")])),
                    vec![ret(list_of(vec![]))],
                ),
                assign(
                    ident("__text"),
                    call_global("__py_fs_read_text", vec![ident("filename")]),
                ),
                assign(
                    ident("__lines"),
                    call_global("__py_linecache_make_lines", vec![ident("__text")]),
                ),
                assign(
                    index(cache(), ident("filename")),
                    tuple_of(vec![
                        call_global("len", vec![ident("__text")]),
                        num(0.0),
                        ident("__lines"),
                        ident("filename"),
                    ]),
                ),
                ret(ident("__lines")),
            ],
        ),
        function(
            "__py_linecache_getlines",
            vec![
                param("filename", None),
                param("module_globals", Some(null())),
            ],
            vec![ret(call_global(
                "__py_linecache_updatecache",
                vec![ident("filename"), ident("module_globals")],
            ))],
        ),
        function(
            "__py_linecache_getline",
            vec![
                param("filename", None),
                param("lineno", None),
                param("module_globals", Some(null())),
                param("lazycache", Some(bool_lit(true))),
            ],
            vec![
                if_stmt(
                    binary(BinOp::LtEq, ident("lineno"), num(0.0)),
                    vec![ret(str_lit(""))],
                ),
                assign(
                    ident("__lines"),
                    call_global(
                        "__py_linecache_updatecache",
                        vec![ident("filename"), ident("module_globals")],
                    ),
                ),
                if_stmt(
                    binary(
                        BinOp::Gt,
                        ident("lineno"),
                        call_global("len", vec![ident("__lines")]),
                    ),
                    vec![ret(str_lit(""))],
                ),
                ret(index(
                    ident("__lines"),
                    binary(BinOp::Sub, ident("lineno"), num(1.0)),
                )),
            ],
        ),
        function(
            "__py_linecache_clearcache",
            vec![],
            vec![call_clear_cache(), ret(null())],
        ),
        function(
            "__py_linecache_checkcache",
            vec![param("filename", Some(null()))],
            vec![call_clear_cache(), ret(null())],
        ),
        function(
            "__py_linecache_lazycache",
            vec![
                param("filename", None),
                param("module_globals", Some(null())),
            ],
            vec![ret(bool_lit(false))],
        ),
    ]
}
