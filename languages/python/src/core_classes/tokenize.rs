//! `token` / `tokenize` — token constants plus a compact lexical scanner.
//!
//! This is not a source prelude. The module declares the `TokenInfo` class and
//! helpers as AST so imports resolve through the Python namespace tree.

use super::builders::*;
use vybe_ast::{BinOp, Expression, Statement};

type Expr = Expression;

const ENDMARKER: i64 = 0;
const NAME: i64 = 1;
const NUMBER: i64 = 2;
const STRING: i64 = 3;
const NEWLINE: i64 = 4;
const INDENT: i64 = 5;
const DEDENT: i64 = 6;
const OP: i64 = 54;
const COMMENT: i64 = 61;
const NL: i64 = 62;
const ENCODING: i64 = 63;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn and(l: Expr, r: Expr) -> Expr {
    op(BinOp::And, l, r)
}

fn or(l: Expr, r: Expr) -> Expr {
    op(BinOp::Or, l, r)
}

fn add(l: Expr, r: Expr) -> Expr {
    op(BinOp::Add, l, r)
}

fn len_of(e: Expr) -> Expr {
    call_global("len", vec![e])
}

fn char_at(name: &str, idx: Expr) -> Expr {
    index(ident(name), idx)
}

fn append_token(
    out: Expr,
    kind: i64,
    text: Expr,
    line_no: Expr,
    start_col: Expr,
    end_col: Expr,
    line_text: Expr,
) -> Statement {
    expr_stmt(call(
        member(out, "append"),
        vec![new(
            "TokenInfo",
            vec![
                i(kind),
                text,
                tuple_of(vec![line_no.clone(), start_col]),
                tuple_of(vec![line_no, end_col]),
                line_text,
            ],
        )],
    ))
}

fn token_type(tok: Expr) -> Expr {
    field_of(tok, "type")
}

fn token_string(tok: Expr) -> Expr {
    field_of(tok, "string")
}

fn normal_token(tok: Expr) -> Expr {
    and(
        and(
            and(
                op(BinOp::NotEq, token_type(tok.clone()), i(ENCODING)),
                op(BinOp::NotEq, token_type(tok.clone()), i(ENDMARKER)),
            ),
            and(
                op(BinOp::NotEq, token_type(tok.clone()), i(NEWLINE)),
                op(BinOp::NotEq, token_type(tok.clone()), i(NL)),
            ),
        ),
        and(
            and(
                op(BinOp::NotEq, token_type(tok.clone()), i(INDENT)),
                op(BinOp::NotEq, token_type(tok.clone()), i(DEDENT)),
            ),
            op(BinOp::NotEq, token_type(tok), i(COMMENT)),
        ),
    )
}

pub(super) fn token_info() -> Statement {
    class(
        "TokenInfo",
        vec![
            init(
                vec![
                    param("type", Some(i(ENDMARKER))),
                    param("string", Some(str_lit(""))),
                    param("start", Some(tuple_of(vec![i(0), i(0)]))),
                    param("end", Some(tuple_of(vec![i(0), i(0)]))),
                    param("line", Some(str_lit(""))),
                ],
                vec![
                    set_this("type", ident("type")),
                    set_this("string", ident("string")),
                    set_this("start", ident("start")),
                    set_this("end", ident("end")),
                    set_this("line", ident("line")),
                    set_this("exact_type", ident("type")),
                ],
            ),
            method(
                "__iter__",
                vec![],
                vec![ret(call_global(
                    "iter",
                    vec![tuple_of(vec![
                        this_field("type"),
                        this_field("string"),
                        this_field("start"),
                        this_field("end"),
                        this_field("line"),
                    ])],
                ))],
            ),
        ],
    )
}

pub(super) fn token_error() -> Statement {
    class_extending("TokenError", &["Exception"], vec![])
}

fn read_all_fn() -> Statement {
    function(
        "__py_tokenize_read_all",
        vec![param("readline", None)],
        vec![
            assign(ident("__text"), str_lit("")),
            while_stmt(
                bool_lit(true),
                vec![
                    assign(ident("__line"), call(ident("readline"), vec![])),
                    if_stmt(
                        op(BinOp::Eq, len_of(ident("__line")), i(0)),
                        vec![Statement::with_span(
                            vybe_ast::StmtKind::Break(vybe_ast::BreakTarget::Implicit),
                            vybe_ast::Span::default(),
                        )],
                    ),
                    assign(
                        ident("__text"),
                        add(
                            ident("__text"),
                            call_global("__py_io_text", vec![ident("__line")]),
                        ),
                    ),
                ],
            ),
            ret(ident("__text")),
        ],
    )
}

fn tokenize_text_fn() -> Statement {
    function(
        "__py_tokenize_text",
        vec![
            param("text", Some(str_lit(""))),
            param("include_encoding", Some(bool_lit(false))),
        ],
        vec![
            assign(ident("__out"), list_of(vec![])),
            if_stmt(
                ident("include_encoding"),
                vec![append_token(
                    ident("__out"),
                    ENCODING,
                    str_lit("utf-8"),
                    i(0),
                    i(0),
                    i(0),
                    str_lit(""),
                )],
            ),
            assign(ident("__current_indent"), i(0)),
            assign(ident("__line_no"), i(1)),
            for_in(
                "__line",
                call(member(ident("text"), "split"), vec![str_lit("\n")]),
                vec![
                    if_stmt(
                        op(BinOp::Gt, len_of(ident("__line")), i(0)),
                        vec![
                            assign(ident("__indent"), i(0)),
                            while_stmt(
                                and(
                                    op(BinOp::Lt, ident("__indent"), len_of(ident("__line"))),
                                    op(
                                        BinOp::Eq,
                                        char_at("__line", ident("__indent")),
                                        str_lit(" "),
                                    ),
                                ),
                                vec![assign(ident("__indent"), add(ident("__indent"), i(1)))],
                            ),
                            if_stmt(
                                op(BinOp::Gt, ident("__indent"), ident("__current_indent")),
                                vec![append_token(
                                    ident("__out"),
                                    INDENT,
                                    slice_range(ident("__line"), i(0), ident("__indent")),
                                    ident("__line_no"),
                                    i(0),
                                    ident("__indent"),
                                    ident("__line"),
                                )],
                            ),
                            if_stmt(
                                op(BinOp::Lt, ident("__indent"), ident("__current_indent")),
                                vec![append_token(
                                    ident("__out"),
                                    DEDENT,
                                    str_lit(""),
                                    ident("__line_no"),
                                    ident("__indent"),
                                    ident("__indent"),
                                    ident("__line"),
                                )],
                            ),
                            assign(ident("__current_indent"), ident("__indent")),
                            assign(ident("__pos"), ident("__indent")),
                            while_stmt(
                                op(BinOp::Lt, ident("__pos"), len_of(ident("__line"))),
                                vec![
                                    assign(ident("__done"), bool_lit(false)),
                                    assign(ident("__ch"), char_at("__line", ident("__pos"))),
                                    if_stmt(
                                        and(
                                            unary_not(ident("__done")),
                                            or(
                                                op(BinOp::Eq, ident("__ch"), str_lit(" ")),
                                                op(BinOp::Eq, ident("__ch"), str_lit("\t")),
                                            ),
                                        ),
                                        vec![
                                            assign(ident("__pos"), add(ident("__pos"), i(1))),
                                            assign(ident("__done"), bool_lit(true)),
                                        ],
                                    ),
                                    if_stmt(
                                        and(
                                            unary_not(ident("__done")),
                                            op(BinOp::Eq, ident("__ch"), str_lit("#")),
                                        ),
                                        vec![
                                            append_token(
                                                ident("__out"),
                                                COMMENT,
                                                slice_from(ident("__line"), ident("__pos")),
                                                ident("__line_no"),
                                                ident("__pos"),
                                                len_of(ident("__line")),
                                                ident("__line"),
                                            ),
                                            assign(ident("__pos"), len_of(ident("__line"))),
                                            assign(ident("__done"), bool_lit(true)),
                                        ],
                                    ),
                                    if_stmt(
                                        and(
                                            unary_not(ident("__done")),
                                            or(
                                                op(BinOp::Eq, ident("__ch"), str_lit("\"")),
                                                op(BinOp::Eq, ident("__ch"), str_lit("'")),
                                            ),
                                        ),
                                        vec![
                                            assign(ident("__quote"), ident("__ch")),
                                            assign(ident("__end"), add(ident("__pos"), i(1))),
                                            while_stmt(
                                                and(
                                                    op(
                                                        BinOp::Lt,
                                                        ident("__end"),
                                                        len_of(ident("__line")),
                                                    ),
                                                    op(
                                                        BinOp::NotEq,
                                                        char_at("__line", ident("__end")),
                                                        ident("__quote"),
                                                    ),
                                                ),
                                                vec![assign(
                                                    ident("__end"),
                                                    add(ident("__end"), i(1)),
                                                )],
                                            ),
                                            if_stmt(
                                                op(
                                                    BinOp::Lt,
                                                    ident("__end"),
                                                    len_of(ident("__line")),
                                                ),
                                                vec![assign(
                                                    ident("__end"),
                                                    add(ident("__end"), i(1)),
                                                )],
                                            ),
                                            append_token(
                                                ident("__out"),
                                                STRING,
                                                slice_range(
                                                    ident("__line"),
                                                    ident("__pos"),
                                                    ident("__end"),
                                                ),
                                                ident("__line_no"),
                                                ident("__pos"),
                                                ident("__end"),
                                                ident("__line"),
                                            ),
                                            assign(ident("__pos"), ident("__end")),
                                            assign(ident("__done"), bool_lit(true)),
                                        ],
                                    ),
                                    if_stmt(
                                        and(
                                            unary_not(ident("__done")),
                                            call(member(ident("__ch"), "isdigit"), vec![]),
                                        ),
                                        vec![
                                            assign(ident("__end"), add(ident("__pos"), i(1))),
                                            while_stmt(
                                                and(
                                                    op(
                                                        BinOp::Lt,
                                                        ident("__end"),
                                                        len_of(ident("__line")),
                                                    ),
                                                    or(
                                                        call(
                                                            member(
                                                                char_at("__line", ident("__end")),
                                                                "isalnum",
                                                            ),
                                                            vec![],
                                                        ),
                                                        op(
                                                            BinOp::Eq,
                                                            char_at("__line", ident("__end")),
                                                            str_lit("."),
                                                        ),
                                                    ),
                                                ),
                                                vec![assign(
                                                    ident("__end"),
                                                    add(ident("__end"), i(1)),
                                                )],
                                            ),
                                            append_token(
                                                ident("__out"),
                                                NUMBER,
                                                slice_range(
                                                    ident("__line"),
                                                    ident("__pos"),
                                                    ident("__end"),
                                                ),
                                                ident("__line_no"),
                                                ident("__pos"),
                                                ident("__end"),
                                                ident("__line"),
                                            ),
                                            assign(ident("__pos"), ident("__end")),
                                            assign(ident("__done"), bool_lit(true)),
                                        ],
                                    ),
                                    if_stmt(
                                        and(
                                            unary_not(ident("__done")),
                                            or(
                                                call(member(ident("__ch"), "isalpha"), vec![]),
                                                op(BinOp::Eq, ident("__ch"), str_lit("_")),
                                            ),
                                        ),
                                        vec![
                                            assign(ident("__end"), add(ident("__pos"), i(1))),
                                            while_stmt(
                                                and(
                                                    op(
                                                        BinOp::Lt,
                                                        ident("__end"),
                                                        len_of(ident("__line")),
                                                    ),
                                                    or(
                                                        call(
                                                            member(
                                                                char_at("__line", ident("__end")),
                                                                "isalnum",
                                                            ),
                                                            vec![],
                                                        ),
                                                        op(
                                                            BinOp::Eq,
                                                            char_at("__line", ident("__end")),
                                                            str_lit("_"),
                                                        ),
                                                    ),
                                                ),
                                                vec![assign(
                                                    ident("__end"),
                                                    add(ident("__end"), i(1)),
                                                )],
                                            ),
                                            append_token(
                                                ident("__out"),
                                                NAME,
                                                slice_range(
                                                    ident("__line"),
                                                    ident("__pos"),
                                                    ident("__end"),
                                                ),
                                                ident("__line_no"),
                                                ident("__pos"),
                                                ident("__end"),
                                                ident("__line"),
                                            ),
                                            assign(ident("__pos"), ident("__end")),
                                            assign(ident("__done"), bool_lit(true)),
                                        ],
                                    ),
                                    if_stmt(
                                        unary_not(ident("__done")),
                                        vec![
                                            append_token(
                                                ident("__out"),
                                                OP,
                                                ident("__ch"),
                                                ident("__line_no"),
                                                ident("__pos"),
                                                add(ident("__pos"), i(1)),
                                                ident("__line"),
                                            ),
                                            assign(ident("__pos"), add(ident("__pos"), i(1))),
                                        ],
                                    ),
                                ],
                            ),
                            append_token(
                                ident("__out"),
                                NEWLINE,
                                str_lit("\n"),
                                ident("__line_no"),
                                len_of(ident("__line")),
                                len_of(ident("__line")),
                                ident("__line"),
                            ),
                        ],
                    ),
                    assign(ident("__line_no"), add(ident("__line_no"), i(1))),
                ],
            ),
            if_stmt(
                op(BinOp::Gt, ident("__current_indent"), i(0)),
                vec![append_token(
                    ident("__out"),
                    DEDENT,
                    str_lit(""),
                    ident("__line_no"),
                    i(0),
                    i(0),
                    str_lit(""),
                )],
            ),
            append_token(
                ident("__out"),
                ENDMARKER,
                str_lit(""),
                ident("__line_no"),
                i(0),
                i(0),
                str_lit(""),
            ),
            ret(ident("__out")),
        ],
    )
}

fn untokenize_fn() -> Statement {
    function(
        "untokenize",
        vec![param("tokens", None)],
        vec![
            assign(ident("__out"), str_lit("")),
            assign(ident("__first"), bool_lit(true)),
            assign(ident("__bytes"), bool_lit(false)),
            for_in(
                "__tok",
                ident("tokens"),
                vec![
                    if_stmt(
                        op(BinOp::Eq, token_type(ident("__tok")), i(ENCODING)),
                        vec![assign(ident("__bytes"), bool_lit(true))],
                    ),
                    if_stmt(
                        normal_token(ident("__tok")),
                        vec![
                            if_stmt(
                                unary_not(ident("__first")),
                                vec![assign(ident("__out"), add(ident("__out"), str_lit(" ")))],
                            ),
                            assign(
                                ident("__out"),
                                add(ident("__out"), token_string(ident("__tok"))),
                            ),
                            assign(ident("__first"), bool_lit(false)),
                        ],
                    ),
                ],
            ),
            if_stmt(
                ident("__bytes"),
                vec![ret(call_global(
                    "bytes",
                    vec![ident("__out"), str_lit("utf-8")],
                ))],
            ),
            ret(ident("__out")),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        read_all_fn(),
        tokenize_text_fn(),
        function(
            "generate_tokens",
            vec![param("readline", None)],
            vec![ret(call_global(
                "__py_tokenize_text",
                vec![
                    call_global("__py_tokenize_read_all", vec![ident("readline")]),
                    bool_lit(false),
                ],
            ))],
        ),
        function(
            "tokenize",
            vec![param("readline", None)],
            vec![ret(call_global(
                "__py_tokenize_text",
                vec![
                    call_global("__py_tokenize_read_all", vec![ident("readline")]),
                    bool_lit(true),
                ],
            ))],
        ),
        untokenize_fn(),
        function(
            "detect_encoding",
            vec![param("readline", None)],
            vec![ret(tuple_of(vec![str_lit("utf-8"), list_of(vec![])]))],
        ),
    ]
}
