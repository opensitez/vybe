//! `io.StringIO` / `io.BytesIO` — in-memory streams.
//!
//! ⛔ These were 158 lines of PARSED PYTHON behind
//! `contains("import io") || contains(", io")`, which fired on 221 corpus
//! tests and was recompiled for every one of them. They are classes, so they
//! are declared here as AST: no scan of the program text, no second parse.
//!
//! Declaring them also fixes construction. A prelude is parsed AFTER the user
//! body is walked, so `BytesIO` was not in `py_defined_classes` when
//! `io.BytesIO(b'data')` was lowered: the call stayed an `ExprKind::Call`
//! instead of normalising to `New`, the constructor bound its arguments one
//! slot over, `initial` fell back to its default and `getvalue()` answered
//! `b''`. `needed_classes` seeds the name BEFORE the walk.
//!
//! The buffer is ONE STRING, not the prelude's list of parts: every reader
//! joined the list first, so the join was pure overhead.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn buf() -> Expr {
    this_field("_buf")
}

fn pos() -> Expr {
    this_field("_pos")
}

fn len_of(e: Expr) -> Expr {
    call_global("len", vec![e])
}

type ClassMemberList = vybe_ast::ClassMember;

/// `self._buf[self._pos:]` — everything not yet read.
fn rest() -> Expr {
    slice_from(buf(), pos())
}

fn common(wrap: fn(Expr) -> Expr) -> Vec<ClassMemberList> {
    vec![
        method(
            "read",
            vec![param("size", Some(i(-1)))],
            vec![
                assign(ident("__r"), rest()),
                if_stmt(
                    op(BinOp::GtEq, ident("size"), i(0)),
                    vec![assign(
                        ident("__r"),
                        slice_range(rest(), i(0), ident("size")),
                    )],
                ),
                set_this("_pos", op(BinOp::Add, pos(), len_of(ident("__r")))),
                ret(wrap(ident("__r"))),
            ],
        ),
        method(
            "read1",
            vec![param("size", Some(i(-1)))],
            vec![ret(call(
                member(ident("self"), "read"),
                vec![ident("size")],
            ))],
        ),
        method(
            "readline",
            vec![],
            vec![
                assign(ident("__rest"), rest()),
                assign(
                    ident("__i"),
                    call(member(ident("__rest"), "find"), vec![str_lit("\n")]),
                ),
                assign(ident("__r"), ident("__rest")),
                if_stmt(
                    op(BinOp::GtEq, ident("__i"), i(0)),
                    vec![assign(
                        ident("__r"),
                        slice_range(ident("__rest"), i(0), op(BinOp::Add, ident("__i"), i(1))),
                    )],
                ),
                set_this("_pos", op(BinOp::Add, pos(), len_of(ident("__r")))),
                ret(wrap(ident("__r"))),
            ],
        ),
        method(
            "readlines",
            vec![],
            vec![
                assign(ident("__out"), list_of(vec![])),
                while_stmt(
                    op(BinOp::Lt, pos(), len_of(buf())),
                    vec![
                        assign(
                            ident("__line"),
                            call(member(ident("self"), "readline"), vec![]),
                        ),
                        if_stmt(
                            op(BinOp::Eq, len_of(ident("__line")), i(0)),
                            vec![Statement::with_span(
                                vybe_ast::StmtKind::Break(vybe_ast::BreakTarget::Implicit),
                                vybe_ast::Span::default(),
                            )],
                        ),
                        expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![ident("__line")],
                        )),
                    ],
                ),
                ret(ident("__out")),
            ],
        ),
        method(
            "__iter__",
            vec![],
            vec![ret(call_global(
                "iter",
                vec![call(member(ident("self"), "readlines"), vec![])],
            ))],
        ),
        method(
            "seek",
            vec![param("pos", Some(i(0))), param("whence", Some(i(0)))],
            vec![
                if_stmt(
                    op(BinOp::Eq, ident("whence"), i(1)),
                    vec![set_this("_pos", op(BinOp::Add, pos(), ident("pos")))],
                ),
                if_stmt(
                    op(BinOp::Eq, ident("whence"), i(2)),
                    vec![set_this(
                        "_pos",
                        op(BinOp::Add, len_of(buf()), ident("pos")),
                    )],
                ),
                if_stmt(
                    op(BinOp::Eq, ident("whence"), i(0)),
                    vec![set_this("_pos", ident("pos"))],
                ),
                ret(pos()),
            ],
        ),
        method("tell", vec![], vec![ret(pos())]),
        method(
            "truncate",
            vec![param("size", Some(null()))],
            vec![
                assign(ident("__end"), pos()),
                if_stmt(
                    is_not_none(ident("size")),
                    vec![assign(ident("__end"), ident("size"))],
                ),
                set_this("_buf", slice_range(buf(), i(0), ident("__end"))),
                ret(ident("__end")),
            ],
        ),
        method("readable", vec![], vec![ret(bool_lit(true))]),
        method("writable", vec![], vec![ret(bool_lit(true))]),
        method("seekable", vec![], vec![ret(bool_lit(true))]),
        method("flush", vec![], vec![ret(null())]),
        method("detach", vec![], vec![ret(null())]),
        method("close", vec![], vec![set_this("closed", bool_lit(true))]),
        method("__enter__", vec![], vec![ret(ident("self"))]),
        method(
            "__exit__",
            any_args(),
            vec![set_this("closed", bool_lit(true)), ret(bool_lit(false))],
        ),
    ]
}

fn text(e: Expr) -> Expr {
    e
}

fn as_bytes(e: Expr) -> Expr {
    call_global("bytes", vec![e, str_lit("utf-8")])
}

pub(super) fn string_io() -> Statement {
    let mut members = vec![
        init(
            vec![param("initial", Some(str_lit("")))],
            vec![
                set_this(
                    "_buf",
                    ternary(is_none(ident("initial")), str_lit(""), ident("initial")),
                ),
                set_this("_pos", i(0)),
                set_this("closed", bool_lit(false)),
            ],
        ),
        method(
            "write",
            vec![param("s", Some(str_lit("")))],
            vec![
                set_this("_buf", op(BinOp::Add, buf(), ident("s"))),
                set_this("_pos", op(BinOp::Add, pos(), len_of(ident("s")))),
                ret(len_of(ident("s"))),
            ],
        ),
        method(
            "writelines",
            vec![param("lines", Some(null()))],
            vec![for_in(
                "__l",
                ident("lines"),
                vec![expr_stmt(call(
                    member(ident("self"), "write"),
                    vec![ident("__l")],
                ))],
            )],
        ),
        method("getvalue", vec![], vec![ret(buf())]),
    ];
    members.extend(common(text));
    class("StringIO", members)
}

pub(super) fn bytes_io() -> Statement {
    let mut members = vec![
        init(
            vec![param("initial", Some(null()))],
            vec![
                set_this("_buf", str_lit("")),
                if_stmt(
                    is_not_none(ident("initial")),
                    vec![set_this(
                        "_buf",
                        call_global("__py_io_text", vec![ident("initial")]),
                    )],
                ),
                set_this("_pos", i(0)),
                set_this("closed", bool_lit(false)),
            ],
        ),
        method(
            "write",
            vec![param("b", Some(null()))],
            vec![
                assign(ident("__s"), call_global("__py_io_text", vec![ident("b")])),
                set_this("_buf", op(BinOp::Add, buf(), ident("__s"))),
                set_this("_pos", op(BinOp::Add, pos(), len_of(ident("__s")))),
                ret(len_of(ident("__s"))),
            ],
        ),
        method(
            "writelines",
            vec![param("lines", Some(null()))],
            vec![for_in(
                "__l",
                ident("lines"),
                vec![expr_stmt(call(
                    member(ident("self"), "write"),
                    vec![ident("__l")],
                ))],
            )],
        ),
        method("getvalue", vec![], vec![ret(as_bytes(buf()))]),
        method("getbuffer", vec![], vec![ret(as_bytes(buf()))]),
    ];
    members.extend(common(as_bytes));
    class("BytesIO", members)
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        // Bytes reach the buffer as text so one representation serves both
        // streams; a str passes through unchanged.
        //
        // ⛔ NO `isinstance`. A spliced class is never walked, so the type
        // name is an unresolved identifier and the test does not mean what it
        // says — `BytesIO(b'data').getvalue()` came back `b'1009711697'`,
        // the element integers concatenated as text. `str(x)` of a str is that
        // same str, but of a bytes it is the REPR (`b'data'`, 7 chars against
        // a length of 4), so the lengths separate them with no type test.
        function(
            "__py_io_text",
            vec![param("value", Some(null()))],
            vec![
                if_stmt(is_none(ident("value")), vec![ret(str_lit(""))]),
                if_stmt(
                    op(
                        BinOp::Eq,
                        len_of(call_global("str", vec![ident("value")])),
                        len_of(ident("value")),
                    ),
                    vec![ret(ident("value"))],
                ),
                assign(ident("__out"), str_lit("")),
                for_in(
                    "__c",
                    ident("value"),
                    vec![assign(
                        ident("__out"),
                        op(
                            BinOp::Add,
                            ident("__out"),
                            call_global("chr", vec![ident("__c")]),
                        ),
                    )],
                ),
                ret(ident("__out")),
            ],
        ),
    ]
}
