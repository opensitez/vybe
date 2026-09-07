//! `fnmatch` and `configparser` — declared, not parsed.
//!
//! Both were prelude SOURCE whose only job was to define a handful of globals
//! and a module object. The module objects are gone: `MODULE_SURFACE` maps
//! `fnmatch.filter` to its global directly, which is what the walker already
//! consults for every other converted module.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn add(l: Expr, r: Expr) -> Expr {
    op(BinOp::Add, l, r)
}

/// `c == "<ch>"`
fn is_ch(ch: &str) -> Expr {
    op(BinOp::Eq, ident("c"), str_lit(ch))
}

pub(super) fn module_functions() -> Vec<Statement> {
    // The regex metacharacters `fnmatch.translate` escapes.
    let specials = [".", "\\", "+", "(", ")", "|", "^", "$", "{", "}"];
    let mut needs_escape = is_ch(specials[0]);
    for ch in &specials[1..] {
        needs_escape = binary(BinOp::Or, needs_escape, is_ch(ch));
    }
    vec![
        function(
            "__fn_translate",
            vec![param("pat", Some(str_lit("")))],
            vec![
                assign(ident("res"), str_lit("")),
                assign(ident("i"), i(0)),
                while_stmt(
                    op(
                        BinOp::Lt,
                        ident("i"),
                        call_global("len", vec![ident("pat")]),
                    ),
                    vec![
                        assign(ident("c"), index(ident("pat"), ident("i"))),
                        assign(ident("done"), bool_lit(false)),
                        if_stmt(
                            is_ch("*"),
                            vec![
                                assign(ident("res"), add(ident("res"), str_lit(".*"))),
                                assign(ident("done"), bool_lit(true)),
                            ],
                        ),
                        if_stmt(
                            is_ch("?"),
                            vec![
                                assign(ident("res"), add(ident("res"), str_lit("."))),
                                assign(ident("done"), bool_lit(true)),
                            ],
                        ),
                        if_stmt(
                            binary(BinOp::And, unary_not(ident("done")), needs_escape.clone()),
                            vec![
                                assign(
                                    ident("res"),
                                    add(
                                        add(ident("res"), call_global("chr", vec![i(92)])),
                                        ident("c"),
                                    ),
                                ),
                                assign(ident("done"), bool_lit(true)),
                            ],
                        ),
                        if_stmt(
                            unary_not(ident("done")),
                            vec![assign(ident("res"), add(ident("res"), ident("c")))],
                        ),
                        assign(ident("i"), add(ident("i"), i(1))),
                    ],
                ),
                ret(add(
                    add(str_lit("(?s:"), ident("res")),
                    add(call_global("chr", vec![i(92)]), str_lit("Z)")),
                )),
            ],
        ),
        function(
            "__py_fnmatch_match",
            vec![param("name", Some(null())), param("pat", Some(null()))],
            vec![ret(call_global(
                "__glob_match",
                vec![ident("name"), ident("pat")],
            ))],
        ),
        // ⛔ NOT a global called `filter`: that is a python builtin, and the
        // module object the prelude used kept it namespaced. `MODULE_SURFACE`
        // does the same job by mapping `fnmatch.filter` onto this name.
        function(
            "__py_fnmatch_filter",
            vec![param("names", Some(null())), param("pat", Some(null()))],
            vec![
                assign(ident("result"), list_of(vec![])),
                for_in(
                    "nm",
                    ident("names"),
                    vec![if_stmt(
                        call_global("__glob_match", vec![ident("nm"), ident("pat")]),
                        vec![expr_stmt(call(
                            member(ident("result"), "append"),
                            vec![ident("nm")],
                        ))],
                    )],
                ),
                ret(ident("result")),
            ],
        ),
    ]
}
