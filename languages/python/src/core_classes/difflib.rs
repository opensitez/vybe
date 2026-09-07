//! `difflib` — small sequence comparison classes/functions as AST declarations.

use super::builders::*;
use vybe_ast::{BinOp, Expression, Statement, StmtKind};

type Expr = Expression;

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn i(v: i64) -> Expr {
    Expression::int(v)
}

fn len_of(e: Expr) -> Expr {
    call_global("len", vec![e])
}

fn append(list: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(list, "append"), vec![value]))
}

fn startswith(value: Expr, prefix: &str) -> Expr {
    call(member(value, "startswith"), vec![str_lit(prefix)])
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

pub(super) fn match_class() -> Statement {
    class(
        "__PyDiffMatch",
        vec![init(
            vec![param("a", None), param("b", None), param("size", None)],
            vec![
                set_this("a", ident("a")),
                set_this("b", ident("b")),
                set_this("size", ident("size")),
            ],
        )],
    )
}

pub(super) fn sequence_matcher() -> Statement {
    class(
        "SequenceMatcher",
        vec![
            init(
                vec![
                    param("isjunk", Some(null())),
                    param("a", Some(str_lit(""))),
                    param("b", Some(str_lit(""))),
                    param("autojunk", Some(bool_lit(true))),
                ],
                vec![
                    set_this("isjunk", ident("isjunk")),
                    set_this("a", ident("a")),
                    set_this("b", ident("b")),
                    set_this("autojunk", ident("autojunk")),
                ],
            ),
            method(
                "set_seq1",
                vec![param("a", None)],
                vec![set_this("a", ident("a"))],
            ),
            method(
                "set_seq2",
                vec![param("b", None)],
                vec![set_this("b", ident("b"))],
            ),
            method(
                "set_seqs",
                vec![param("a", None), param("b", None)],
                vec![set_this("a", ident("a")), set_this("b", ident("b"))],
            ),
            method(
                "ratio",
                vec![],
                vec![ret(call_global(
                    "__py_difflib_ratio",
                    vec![this_field("a"), this_field("b")],
                ))],
            ),
            method(
                "quick_ratio",
                vec![],
                vec![ret(call_global(
                    "__py_difflib_ratio",
                    vec![this_field("a"), this_field("b")],
                ))],
            ),
            method(
                "real_quick_ratio",
                vec![],
                vec![ret(call_global(
                    "__py_difflib_ratio",
                    vec![this_field("a"), this_field("b")],
                ))],
            ),
            method(
                "find_longest_match",
                vec![
                    param("alo", Some(i(0))),
                    param("ahi", Some(null())),
                    param("blo", Some(i(0))),
                    param("bhi", Some(null())),
                ],
                vec![
                    if_stmt(is_none(ident("ahi")), vec![assign(ident("ahi"), len_of(this_field("a")))]),
                    if_stmt(is_none(ident("bhi")), vec![assign(ident("bhi"), len_of(this_field("b")))]),
                    assign(ident("__best_i"), ident("alo")),
                    assign(ident("__best_j"), ident("blo")),
                    assign(ident("__best_size"), i(0)),
                    assign(ident("__i"), ident("alo")),
                    while_stmt(
                        op(BinOp::Lt, ident("__i"), ident("ahi")),
                        vec![
                            assign(ident("__j"), ident("blo")),
                            while_stmt(
                                op(BinOp::Lt, ident("__j"), ident("bhi")),
                                vec![
                                    assign(ident("__k"), i(0)),
                                    while_stmt(
                                        op(
                                            BinOp::And,
                                            op(
                                                BinOp::And,
                                                op(BinOp::Lt, op(BinOp::Add, ident("__i"), ident("__k")), ident("ahi")),
                                                op(BinOp::Lt, op(BinOp::Add, ident("__j"), ident("__k")), ident("bhi")),
                                            ),
                                            op(
                                                BinOp::Eq,
                                                index(this_field("a"), op(BinOp::Add, ident("__i"), ident("__k"))),
                                                index(this_field("b"), op(BinOp::Add, ident("__j"), ident("__k"))),
                                            ),
                                        ),
                                        vec![assign(ident("__k"), op(BinOp::Add, ident("__k"), i(1)))],
                                    ),
                                    if_stmt(
                                        op(BinOp::Gt, ident("__k"), ident("__best_size")),
                                        vec![
                                            assign(ident("__best_i"), ident("__i")),
                                            assign(ident("__best_j"), ident("__j")),
                                            assign(ident("__best_size"), ident("__k")),
                                        ],
                                    ),
                                    assign(ident("__j"), op(BinOp::Add, ident("__j"), i(1))),
                                ],
                            ),
                            assign(ident("__i"), op(BinOp::Add, ident("__i"), i(1))),
                        ],
                    ),
                    ret(new(
                        "__PyDiffMatch",
                        vec![ident("__best_i"), ident("__best_j"), ident("__best_size")],
                    )),
                ],
            ),
            method(
                "get_matching_blocks",
                vec![],
                vec![ret(list_of(vec![
                    call(
                        member(ident("self"), "find_longest_match"),
                        vec![i(0), len_of(this_field("a")), i(0), len_of(this_field("b"))],
                    ),
                    new("__PyDiffMatch", vec![len_of(this_field("a")), len_of(this_field("b")), i(0)]),
                ]))],
            ),
            method(
                "get_opcodes",
                vec![],
                vec![ret(list_of(vec![
                    tuple_of(vec![str_lit("equal"), i(0), i(0), i(0), i(0)]),
                    tuple_of(vec![
                        str_lit("replace"),
                        i(0),
                        len_of(this_field("a")),
                        i(0),
                        len_of(this_field("b")),
                    ]),
                ]))],
            ),
        ],
    )
}

pub(super) fn differ() -> Statement {
    class(
        "Differ",
        vec![
            init(vec![], vec![]),
            method(
                "compare",
                vec![param("a", None), param("b", None)],
                vec![
                    assign(ident("__out"), list_of(vec![])),
                    assign(ident("__i"), i(0)),
                    while_stmt(
                        op(
                            BinOp::Or,
                            op(BinOp::Lt, ident("__i"), len_of(ident("a"))),
                            op(BinOp::Lt, ident("__i"), len_of(ident("b"))),
                        ),
                        vec![
                            if_stmt(
                                op(
                                    BinOp::And,
                                    op(BinOp::Lt, ident("__i"), len_of(ident("a"))),
                                    op(BinOp::Lt, ident("__i"), len_of(ident("b"))),
                                ),
                                vec![if_else(
                                    op(
                                        BinOp::Eq,
                                        index(ident("a"), ident("__i")),
                                        index(ident("b"), ident("__i")),
                                    ),
                                    vec![append(
                                        ident("__out"),
                                        op(BinOp::Add, str_lit("  "), index(ident("a"), ident("__i"))),
                                    )],
                                    vec![
                                        append(
                                            ident("__out"),
                                            op(BinOp::Add, str_lit("- "), index(ident("a"), ident("__i"))),
                                        ),
                                        append(
                                            ident("__out"),
                                            op(BinOp::Add, str_lit("+ "), index(ident("b"), ident("__i"))),
                                        ),
                                    ],
                                )],
                            ),
                            if_stmt(
                                op(
                                    BinOp::And,
                                    op(BinOp::Lt, ident("__i"), len_of(ident("a"))),
                                    op(BinOp::GtEq, ident("__i"), len_of(ident("b"))),
                                ),
                                vec![append(
                                    ident("__out"),
                                    op(BinOp::Add, str_lit("- "), index(ident("a"), ident("__i"))),
                                )],
                            ),
                            if_stmt(
                                op(
                                    BinOp::And,
                                    op(BinOp::GtEq, ident("__i"), len_of(ident("a"))),
                                    op(BinOp::Lt, ident("__i"), len_of(ident("b"))),
                                ),
                                vec![append(
                                    ident("__out"),
                                    op(BinOp::Add, str_lit("+ "), index(ident("b"), ident("__i"))),
                                )],
                            ),
                            assign(ident("__i"), op(BinOp::Add, ident("__i"), i(1))),
                        ],
                    ),
                    ret(ident("__out")),
                ],
            ),
        ],
    )
}

pub(super) fn html_diff() -> Statement {
    class(
        "HtmlDiff",
        vec![
            init(vec![], vec![]),
            method("make_table", any_args(), vec![ret(str_lit("<table></table>"))]),
            method("make_file", any_args(), vec![ret(str_lit("<html><table></table></html>"))]),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "get_close_matches",
            vec![
                param("word", None),
                param("possibilities", None),
                param("n", Some(i(3))),
                param("cutoff", Some(num(0.6))),
            ],
            vec![
                assign(ident("__out"), list_of(vec![])),
                for_in(
                    "__candidate",
                    ident("possibilities"),
                    vec![if_stmt(
                        op(
                            BinOp::And,
                            op(
                                BinOp::Gt,
                                call_global(
                                    "__py_difflib_ratio",
                                    vec![ident("word"), ident("__candidate")],
                                ),
                                ident("cutoff"),
                            ),
                            op(BinOp::Lt, len_of(ident("__out")), ident("n")),
                        ),
                        vec![append(ident("__out"), ident("__candidate"))],
                    )],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "restore",
            vec![param("delta", None), param("which", None)],
            vec![
                assign(ident("__out"), list_of(vec![])),
                for_in(
                    "__line",
                    ident("delta"),
                    vec![
                        if_stmt(
                            op(
                                BinOp::And,
                                op(BinOp::Eq, ident("which"), i(1)),
                                op(
                                    BinOp::Or,
                                    startswith(ident("__line"), "  "),
                                    startswith(ident("__line"), "- "),
                                ),
                            ),
                            vec![append(ident("__out"), slice_from(ident("__line"), i(2)))],
                        ),
                        if_stmt(
                            op(
                                BinOp::And,
                                op(BinOp::Eq, ident("which"), i(2)),
                                op(
                                    BinOp::Or,
                                    startswith(ident("__line"), "  "),
                                    startswith(ident("__line"), "+ "),
                                ),
                            ),
                            vec![append(ident("__out"), slice_from(ident("__line"), i(2)))],
                        ),
                    ],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "unified_diff",
            vec![
                param("a", None),
                param("b", None),
                param("fromfile", Some(str_lit(""))),
                param("tofile", Some(str_lit(""))),
            ],
            vec![ret(list_of(vec![
                op(BinOp::Add, op(BinOp::Add, str_lit("--- "), ident("fromfile")), str_lit("\n")),
                op(BinOp::Add, op(BinOp::Add, str_lit("+++ "), ident("tofile")), str_lit("\n")),
            ]))],
        ),
        function(
            "context_diff",
            vec![param("a", None), param("b", None)],
            vec![ret(list_of(vec![str_lit("*** \n"), str_lit("--- \n")]))],
        ),
        function(
            "IS_CHARACTER_JUNK",
            vec![param("ch", None)],
            vec![ret(op(
                BinOp::Or,
                op(BinOp::Eq, ident("ch"), str_lit(" ")),
                op(BinOp::Eq, ident("ch"), str_lit("\t")),
            ))],
        ),
        function(
            "IS_LINE_JUNK",
            vec![param("line", None)],
            vec![ret(bool_lit(false))],
        ),
    ]
}
