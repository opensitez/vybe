use super::builders::*;
use vybe_ast::{BinOp, Expression as Expr, Statement};

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn int(value: i64) -> Expr {
    Expr::int(value)
}

fn stmt_expr(expr: Expr) -> Statement {
    expr_stmt(expr)
}

fn call_method(object: Expr, name: &str, args: Vec<Expr>) -> Expr {
    call(member(object, name), args)
}

fn len_of(expr: Expr) -> Expr {
    call_global("len", vec![expr])
}

fn startswith(expr: Expr, prefix: &str, start: Expr) -> Expr {
    call_method(expr, "startswith", vec![str_lit(prefix), start])
}

fn find_from(expr: Expr, needle: &str, start: Expr) -> Expr {
    call_method(expr, "find", vec![str_lit(needle), start])
}

fn strip(expr: Expr) -> Expr {
    call_method(expr, "strip", vec![])
}

fn strip_quotes(expr: Expr) -> Expr {
    call_method(expr, "strip", vec![str_lit("\"'")])
}

fn slice_to(object: Expr, end: Expr) -> Expr {
    index(
        object,
        Expr::with_span(
            vybe_ast::ExprKind::Slice {
                lower: None,
                upper: Some(Box::new(end)),
                step: None,
            },
            span(),
        ),
    )
}

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn append_attr(value: Expr) -> Statement {
    stmt_expr(call_method(ident("__attrs"), "append", vec![value]))
}

fn callback(name: &str, args: Vec<Expr>) -> Statement {
    stmt_expr(call(member(ident("self"), name), args))
}

fn set_i(value: Expr) -> Statement {
    assign(ident("__i"), value)
}

fn feed_body() -> Vec<Statement> {
    vec![
        assign(ident("__data"), call_global("str", vec![ident("data")])),
        assign(ident("__i"), int(0)),
        while_stmt(
            op(BinOp::Lt, ident("__i"), len_of(ident("__data"))),
            vec![if_else(
                startswith(ident("__data"), "<!--", ident("__i")),
                vec![
                    assign(
                        ident("__j"),
                        find_from(ident("__data"), "-->", op(BinOp::Add, ident("__i"), int(4))),
                    ),
                    if_stmt(
                        op(BinOp::Eq, ident("__j"), int(-1)),
                        vec![assign(ident("__j"), len_of(ident("__data")))],
                    ),
                    callback(
                        "handle_comment",
                        vec![slice_range(
                            ident("__data"),
                            op(BinOp::Add, ident("__i"), int(4)),
                            ident("__j"),
                        )],
                    ),
                    set_i(op(BinOp::Add, ident("__j"), int(3))),
                ],
                vec![if_else(
                    startswith(ident("__data"), "</", ident("__i")),
                    vec![
                        assign(
                            ident("__j"),
                            find_from(ident("__data"), ">", op(BinOp::Add, ident("__i"), int(2))),
                        ),
                        if_stmt(
                            op(BinOp::Eq, ident("__j"), int(-1)),
                            vec![assign(ident("__j"), len_of(ident("__data")))],
                        ),
                        callback(
                            "handle_endtag",
                            vec![strip(slice_range(
                                ident("__data"),
                                op(BinOp::Add, ident("__i"), int(2)),
                                ident("__j"),
                            ))],
                        ),
                        set_i(op(BinOp::Add, ident("__j"), int(1))),
                    ],
                    vec![if_else(
                        startswith(ident("__data"), "<?", ident("__i")),
                        vec![
                            assign(
                                ident("__j"),
                                find_from(
                                    ident("__data"),
                                    ">",
                                    op(BinOp::Add, ident("__i"), int(2)),
                                ),
                            ),
                            if_stmt(
                                op(BinOp::Eq, ident("__j"), int(-1)),
                                vec![assign(ident("__j"), len_of(ident("__data")))],
                            ),
                            callback(
                                "handle_pi",
                                vec![slice_range(
                                    ident("__data"),
                                    op(BinOp::Add, ident("__i"), int(2)),
                                    ident("__j"),
                                )],
                            ),
                            set_i(op(BinOp::Add, ident("__j"), int(1))),
                        ],
                        vec![if_else(
                            startswith(ident("__data"), "<!", ident("__i")),
                            vec![
                                assign(
                                    ident("__j"),
                                    find_from(
                                        ident("__data"),
                                        ">",
                                        op(BinOp::Add, ident("__i"), int(2)),
                                    ),
                                ),
                                if_stmt(
                                    op(BinOp::Eq, ident("__j"), int(-1)),
                                    vec![assign(ident("__j"), len_of(ident("__data")))],
                                ),
                                callback(
                                    "handle_decl",
                                    vec![strip(slice_range(
                                        ident("__data"),
                                        op(BinOp::Add, ident("__i"), int(2)),
                                        ident("__j"),
                                    ))],
                                ),
                                set_i(op(BinOp::Add, ident("__j"), int(1))),
                            ],
                            vec![if_else(
                                op(
                                    BinOp::Eq,
                                    index(ident("__data"), ident("__i")),
                                    str_lit("<"),
                                ),
                                vec![
                                    assign(
                                        ident("__j"),
                                        find_from(
                                            ident("__data"),
                                            ">",
                                            op(BinOp::Add, ident("__i"), int(1)),
                                        ),
                                    ),
                                    if_stmt(
                                        op(BinOp::Eq, ident("__j"), int(-1)),
                                        vec![assign(ident("__j"), len_of(ident("__data")))],
                                    ),
                                    assign(
                                        ident("__content"),
                                        strip(slice_range(
                                            ident("__data"),
                                            op(BinOp::Add, ident("__i"), int(1)),
                                            ident("__j"),
                                        )),
                                    ),
                                    assign(
                                        ident("__self_closing"),
                                        call_method(
                                            ident("__content"),
                                            "endswith",
                                            vec![str_lit("/")],
                                        ),
                                    ),
                                    if_stmt(
                                        ident("__self_closing"),
                                        vec![assign(
                                            ident("__content"),
                                            strip(slice_to(
                                                ident("__content"),
                                                op(BinOp::Sub, len_of(ident("__content")), int(1)),
                                            )),
                                        )],
                                    ),
                                    assign(
                                        ident("__parts"),
                                        call_method(ident("__content"), "split", vec![]),
                                    ),
                                    assign(ident("__tag"), str_lit("")),
                                    if_stmt(
                                        op(BinOp::Gt, len_of(ident("__parts")), int(0)),
                                        vec![assign(
                                            ident("__tag"),
                                            index(ident("__parts"), int(0)),
                                        )],
                                    ),
                                    assign(ident("__attrs"), list_of(vec![])),
                                    for_in(
                                        "__part",
                                        slice_from(ident("__parts"), int(1)),
                                        vec![if_else(
                                            contains(ident("__part"), str_lit("=")),
                                            vec![
                                                assign(
                                                    ident("__kv"),
                                                    call_method(
                                                        ident("__part"),
                                                        "split",
                                                        vec![str_lit("="), int(1)],
                                                    ),
                                                ),
                                                append_attr(tuple_of(vec![
                                                    index(ident("__kv"), int(0)),
                                                    strip_quotes(index(ident("__kv"), int(1))),
                                                ])),
                                            ],
                                            vec![append_attr(tuple_of(vec![
                                                ident("__part"),
                                                null(),
                                            ]))],
                                        )],
                                    ),
                                    if_else(
                                        ident("__self_closing"),
                                        vec![callback(
                                            "handle_startendtag",
                                            vec![ident("__tag"), ident("__attrs")],
                                        )],
                                        vec![callback(
                                            "handle_starttag",
                                            vec![ident("__tag"), ident("__attrs")],
                                        )],
                                    ),
                                    set_i(op(BinOp::Add, ident("__j"), int(1))),
                                ],
                                vec![
                                    assign(
                                        ident("__j"),
                                        find_from(ident("__data"), "<", ident("__i")),
                                    ),
                                    if_stmt(
                                        op(BinOp::Eq, ident("__j"), int(-1)),
                                        vec![assign(ident("__j"), len_of(ident("__data")))],
                                    ),
                                    callback(
                                        "handle_data",
                                        vec![slice_range(
                                            ident("__data"),
                                            ident("__i"),
                                            ident("__j"),
                                        )],
                                    ),
                                    set_i(ident("__j")),
                                ],
                            )],
                        )],
                    )],
                )],
            )],
        ),
        ret(null()),
    ]
}

pub(super) fn html_parser() -> Statement {
    class(
        "HTMLParser",
        vec![
            init(
                vec![],
                vec![set_this("_line", int(1)), set_this("_col", int(0))],
            ),
            method(
                "reset",
                any_args(),
                vec![
                    set_this("_line", int(1)),
                    set_this("_col", int(0)),
                    ret(null()),
                ],
            ),
            method(
                "getpos",
                any_args(),
                vec![ret(tuple_of(vec![this_field("_line"), this_field("_col")]))],
            ),
            method("feed", vec![param("data", None)], feed_body()),
            method(
                "handle_starttag",
                vec![param("tag", None), param("attrs", None)],
                vec![ret(null())],
            ),
            method(
                "handle_startendtag",
                vec![param("tag", None), param("attrs", None)],
                vec![ret(null())],
            ),
            method("handle_endtag", vec![param("tag", None)], vec![ret(null())]),
            method("handle_data", vec![param("data", None)], vec![ret(null())]),
            method(
                "handle_comment",
                vec![param("data", None)],
                vec![ret(null())],
            ),
            method("handle_decl", vec![param("decl", None)], vec![ret(null())]),
            method("handle_pi", vec![param("data", None)], vec![ret(null())]),
        ],
    )
}
