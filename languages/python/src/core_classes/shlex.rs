//! `shlex` module helpers declared as AST, not parsed source prelude.

use super::builders::*;
use vybe_ast::{BinOp, BreakTarget, Expression, Statement, StmtKind};

type Expr = Expression;

fn op(op: BinOp, left: Expr, right: Expr) -> Expr {
    binary(op, left, right)
}

fn eq(left: Expr, right: Expr) -> Expr {
    op(BinOp::Eq, left, right)
}

fn ne(left: Expr, right: Expr) -> Expr {
    op(BinOp::NotEq, left, right)
}

fn and(left: Expr, right: Expr) -> Expr {
    op(BinOp::And, left, right)
}

fn or_expr(left: Expr, right: Expr) -> Expr {
    op(BinOp::Or, left, right)
}

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn py_attr(object: Expr, field: &str) -> Expr {
    call_global("__py_attr_read", vec![object, str_lit(field)])
}

fn if_else(cond: Expr, then_body: Vec<Statement>, else_body: Vec<Statement>) -> Statement {
    Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            elifs: vec![],
            else_body: Some(else_body),
        },
        span(),
    )
}

fn break_stmt() -> Statement {
    Statement::with_span(StmtKind::Break(BreakTarget::Implicit), span())
}

fn append(list: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(list, "append"), vec![value]))
}

fn add_assign(name: &str, value: Expr) -> Statement {
    assign(ident(name), add(ident(name), value))
}

fn shlex_without_comments(text: Expr, commenters: Expr) -> Vec<Statement> {
    vec![
        assign(ident("__shlex_text"), text),
        assign(ident("__shlex_commenter"), commenters),
        assign(ident("text"), str_lit("")),
        assign(ident("__shlex_skip"), bool_lit(false)),
        for_in(
            "__shlex_ch",
            ident("__shlex_text"),
            vec![if_else(
                ident("__shlex_skip"),
                vec![if_stmt(
                    eq(ident("__shlex_ch"), str_lit("\n")),
                    vec![
                        assign(ident("__shlex_skip"), bool_lit(false)),
                        add_assign("text", ident("__shlex_ch")),
                    ],
                )],
                vec![if_else(
                    eq(ident("__shlex_ch"), ident("__shlex_commenter")),
                    vec![assign(ident("__shlex_skip"), bool_lit(true))],
                    vec![add_assign("text", ident("__shlex_ch"))],
                )],
            )],
        ),
    ]
}

pub(super) fn shlex_class() -> Statement {
    class(
        "__py_shlex_class",
        vec![
            init(
                vec![
                    param("instream", Some(null())),
                    param("posix", Some(bool_lit(false))),
                    param("punctuation_chars", Some(bool_lit(false))),
                ],
                vec![
                    if_else(
                        is_none(ident("instream")),
                        vec![set_this("text", str_lit(""))],
                        vec![set_this("text", call_global("str", vec![ident("instream")]))],
                    ),
                    set_this("posix", ident("posix")),
                    set_this("whitespace_split", bool_lit(false)),
                    set_this("commenters", str_lit("#")),
                ],
            ),
            method(
                "__iter__",
                vec![],
                vec![
                    assign(ident("text"), this_field("text")),
                    if_stmt(
                        ne(this_field("commenters"), str_lit("")),
                        shlex_without_comments(ident("text"), this_field("commenters")),
                    ),
                    ret(call_global(
                        "iter",
                        vec![call_global(
                            "__py_shlex_split",
                            vec![ident("text"), bool_lit(false), this_field("posix")],
                        )],
                    )),
                ],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "__py_shlex_needs_quote",
            vec![param("s", None)],
            vec![
                if_stmt(eq(ident("s"), str_lit("")), vec![ret(bool_lit(true))]),
                assign(
                    ident("safe"),
                    str_lit("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_@%+=:,./-"),
                ),
                for_in(
                    "ch",
                    ident("s"),
                    vec![if_stmt(
                        unary_not(contains(ident("safe"), ident("ch"))),
                        vec![ret(bool_lit(true))],
                    )],
                ),
                ret(bool_lit(false)),
            ],
        ),
        function(
            "__py_shlex_quote",
            vec![param("s", None)],
            vec![
                if_stmt(
                    unary_not(call_global("__py_shlex_needs_quote", vec![ident("s")])),
                    vec![ret(ident("s"))],
                ),
                assign(ident("out"), str_lit("'")),
                for_in(
                    "ch",
                    ident("s"),
                    vec![if_else(
                        eq(ident("ch"), str_lit("'")),
                        vec![add_assign("out", str_lit("'\"'\"'"))],
                        vec![add_assign("out", ident("ch"))],
                    )],
                ),
                ret(add(ident("out"), str_lit("'"))),
            ],
        ),
        function(
            "__py_shlex_split",
            vec![
                param("s", None),
                param("comments", Some(bool_lit(false))),
                param("posix", Some(bool_lit(true))),
            ],
            vec![
                assign(ident("out"), list_of(vec![])),
                assign(ident("cur"), str_lit("")),
                assign(ident("quote"), str_lit("")),
                assign(ident("esc"), bool_lit(false)),
                for_in(
                    "ch",
                    ident("s"),
                    vec![if_else(
                        ident("esc"),
                        vec![add_assign("cur", ident("ch")), assign(ident("esc"), bool_lit(false))],
                        vec![if_else(
                            ne(ident("quote"), str_lit("")),
                            vec![if_else(
                                eq(ident("ch"), ident("quote")),
                                vec![if_else(
                                    ident("posix"),
                                    vec![assign(ident("quote"), str_lit(""))],
                                    vec![add_assign("cur", ident("ch")), assign(ident("quote"), str_lit(""))],
                                )],
                                vec![add_assign("cur", ident("ch"))],
                            )],
                            vec![if_else(
                                and(eq(ident("ch"), str_lit("\\")), ident("posix")),
                                vec![assign(ident("esc"), bool_lit(true))],
                                vec![if_else(
                                    or_expr(eq(ident("ch"), str_lit("'")), eq(ident("ch"), str_lit("\""))),
                                    vec![if_else(
                                        ident("posix"),
                                        vec![assign(ident("quote"), ident("ch"))],
                                        vec![
                                            assign(ident("quote"), ident("ch")),
                                            add_assign("cur", ident("ch")),
                                        ],
                                    )],
                                    vec![if_else(
                                        and(ident("comments"), eq(ident("ch"), str_lit("#"))),
                                        vec![break_stmt()],
                                        vec![if_else(
                                            or_expr(
                                                or_expr(
                                                    eq(ident("ch"), str_lit(" ")),
                                                    eq(ident("ch"), str_lit("\t")),
                                                ),
                                                eq(ident("ch"), str_lit("\n")),
                                            ),
                                            vec![if_stmt(
                                                ne(ident("cur"), str_lit("")),
                                                vec![
                                                    append(ident("out"), ident("cur")),
                                                    assign(ident("cur"), str_lit("")),
                                                ],
                                            )],
                                            vec![add_assign("cur", ident("ch"))],
                                        )],
                                    )],
                                )],
                            )],
                        )],
                    )],
                ),
                if_stmt(
                    ne(ident("quote"), str_lit("")),
                    vec![raise_call("ValueError", vec![str_lit("No closing quotation")])],
                ),
                if_stmt(
                    ne(ident("cur"), str_lit("")),
                    vec![append(ident("out"), ident("cur"))],
                ),
                ret(ident("out")),
            ],
        ),
        function(
            "__py_shlex_join",
            vec![param("parts", None)],
            vec![
                assign(ident("joined"), str_lit("")),
                assign(ident("first"), bool_lit(true)),
                for_in(
                    "p",
                    ident("parts"),
                    vec![
                        assign(ident("q"), call_global("__py_shlex_quote", vec![ident("p")])),
                        if_else(
                            ident("first"),
                            vec![
                                assign(ident("joined"), ident("q")),
                                assign(ident("first"), bool_lit(false)),
                            ],
                            vec![assign(
                                ident("joined"),
                                add(add(ident("joined"), str_lit(" ")), ident("q")),
                            )],
                        ),
                    ],
                ),
                ret(ident("joined")),
            ],
        ),
        function(
            "__py_shlex_tokens",
            vec![param("obj", None)],
            vec![
                assign(ident("text"), py_attr(ident("obj"), "text")),
                if_stmt(
                    ne(py_attr(ident("obj"), "commenters"), str_lit("")),
                    shlex_without_comments(ident("text"), py_attr(ident("obj"), "commenters")),
                ),
                ret(call_global(
                    "__py_shlex_split",
                    vec![ident("text"), bool_lit(false), py_attr(ident("obj"), "posix")],
                )),
            ],
        ),
    ]
}
