//! The bytes helper surface — `repr`, the str/bytes conversions, `bytes.join`
//! and the incremental codecs.
//!
//! ⛔ 55 lines of PARSED PYTHON that fired on 1368 corpus tests — more than any
//! other prelude — because the gate is any `b'` or `b"` in the source. Declared
//! here as AST instead: the gate is unchanged, but nothing is parsed.

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

/// `codecs.getincrementalencoder(...)()` — the pair only has to encode and
/// decode; the incremental state a real codec keeps is never observed here.
pub(super) fn incremental_encoder() -> Statement {
    class(
        "__PyIncrementalEncoder",
        vec![method(
            "encode",
            vec![
                param("value", Some(str_lit(""))),
                param("final", Some(bool_lit(false))),
            ],
            vec![ret(call_global(
                "bytes",
                vec![ident("value"), str_lit("utf-8")],
            ))],
        )],
    )
}

pub(super) fn incremental_decoder() -> Statement {
    class(
        "__PyIncrementalDecoder",
        vec![method(
            "decode",
            vec![
                param("value", Some(null())),
                param("final", Some(bool_lit(false))),
            ],
            vec![ret(call_global(
                "__vybe_bytes_decode",
                vec![ident("value")],
            ))],
        )],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        // `b'\t\n\\'` — the escapes CPython's repr emits, then printable
        // ASCII, then `\xNN`. The backslash comes from `chr(92)`: a literal
        // one does not survive python's string-escape lowering.
        function(
            "__vybe_bytes_repr",
            vec![param("a", Some(null()))],
            vec![
                assign(ident("bs"), call_global("chr", vec![i(92)])),
                assign(ident("hexd"), str_lit("0123456789abcdef")),
                assign(ident("r"), str_lit("b'")),
                for_in(
                    "b",
                    ident("a"),
                    vec![
                        assign(ident("done"), bool_lit(false)),
                        escape(9, "t"),
                        escape(10, "n"),
                        escape(13, "r"),
                        // The backslash escapes to two backslashes.
                        if_stmt(
                            op(BinOp::Eq, ident("b"), i(92)),
                            vec![
                                assign(ident("r"), add(ident("r"), add(ident("bs"), ident("bs")))),
                                assign(ident("done"), bool_lit(true)),
                            ],
                        ),
                        escape(39, "'"),
                        if_stmt(
                            binary(
                                BinOp::And,
                                unary_not(ident("done")),
                                binary(
                                    BinOp::And,
                                    op(BinOp::GtEq, ident("b"), i(32)),
                                    op(BinOp::LtEq, ident("b"), i(126)),
                                ),
                            ),
                            vec![
                                assign(
                                    ident("r"),
                                    add(ident("r"), call_global("chr", vec![ident("b")])),
                                ),
                                assign(ident("done"), bool_lit(true)),
                            ],
                        ),
                        if_stmt(
                            unary_not(ident("done")),
                            vec![assign(
                                ident("r"),
                                add(
                                    add(
                                        add(ident("r"), add(ident("bs"), str_lit("x"))),
                                        index(ident("hexd"), op(BinOp::Shr, ident("b"), i(4))),
                                    ),
                                    index(ident("hexd"), op(BinOp::BitAnd, ident("b"), i(15))),
                                ),
                            )],
                        ),
                    ],
                ),
                ret(add(ident("r"), str_lit("'"))),
            ],
        ),
        function(
            "__vybe_str_encode",
            vec![param("s", Some(str_lit("")))],
            vec![
                assign(ident("out"), list_of(vec![])),
                for_in(
                    "ch",
                    ident("s"),
                    vec![expr_stmt(call(
                        member(ident("out"), "append"),
                        vec![call_global("ord", vec![ident("ch")])],
                    ))],
                ),
                ret(ident("out")),
            ],
        ),
        function(
            "__vybe_bytes_decode",
            vec![param("a", Some(null()))],
            vec![
                assign(ident("r"), str_lit("")),
                for_in(
                    "b",
                    ident("a"),
                    vec![assign(
                        ident("r"),
                        add(ident("r"), call_global("chr", vec![ident("b")])),
                    )],
                ),
                ret(ident("r")),
            ],
        ),
        function(
            "__py_bytes_join",
            vec![param("sep", Some(null())), param("iterable", Some(null()))],
            vec![
                assign(
                    ident("s"),
                    call_global("__vybe_bytes_decode", vec![ident("sep")]),
                ),
                assign(ident("out"), str_lit("")),
                assign(ident("first"), bool_lit(true)),
                for_in(
                    "item",
                    ident("iterable"),
                    vec![
                        if_stmt(
                            unary_not(ident("first")),
                            vec![assign(ident("out"), add(ident("out"), ident("s")))],
                        ),
                        // `str(x)` of a str is that str; of a bytes it is the
                        // repr, so the lengths separate them without a type
                        // test — a spliced class never gets `isinstance`
                        // resolved.
                        if_stmt(
                            op(
                                BinOp::Eq,
                                call_global("len", vec![call_global("str", vec![ident("item")])]),
                                call_global("len", vec![ident("item")]),
                            ),
                            vec![assign(ident("out"), add(ident("out"), ident("item")))],
                        ),
                        if_stmt(
                            op(
                                BinOp::NotEq,
                                call_global("len", vec![call_global("str", vec![ident("item")])]),
                                call_global("len", vec![ident("item")]),
                            ),
                            vec![assign(
                                ident("out"),
                                add(
                                    ident("out"),
                                    call_global("__vybe_bytes_decode", vec![ident("item")]),
                                ),
                            )],
                        ),
                        assign(ident("first"), bool_lit(false)),
                    ],
                ),
                ret(call_global("bytes", vec![ident("out"), str_lit("utf-8")])),
            ],
        ),
    ]
}

/// `if b == <code> and not done: r += bs + "<ch>"`
fn escape(code: i64, ch: &str) -> Statement {
    if_stmt(
        op(BinOp::Eq, ident("b"), i(code)),
        vec![
            assign(ident("r"), add(ident("r"), add(ident("bs"), str_lit(ch)))),
            assign(ident("done"), bool_lit(true)),
        ],
    )
}
