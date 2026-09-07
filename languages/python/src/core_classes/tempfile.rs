//! `tempfile.NamedTemporaryFile` / `TemporaryFile` — a real file OBJECT.
//!
//! ⛔ The adapter used to build this with `class_slots::emit_class_alloc` plus
//! stamped fields. That is the hand-rolled construction `core_classes/mod.rs`
//! warns about: an anonymous struct with no type, no vtable and no prototype,
//! so `type(f).__name__` answered with the path string and the object had no
//! methods at all. The PATH stays a primitive (`__py_temp_path` — genuine host
//! work); only the wrapper moves here, where it gets an rtt and real methods.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn path() -> Expr {
    this_field("name")
}

pub(super) fn named_temp_file() -> Statement {
    class(
        "__PyNamedTempFile",
        vec![
            // CPython's signature order; the corpus passes these by keyword.
            init(
                vec![
                    param("mode", Some(str_lit("w+b"))),
                    param("buffering", Some(i(-1))),
                    param("encoding", Some(null())),
                    param("newline", Some(null())),
                    param("suffix", Some(str_lit(""))),
                    param("prefix", Some(str_lit(""))),
                    param("dir", Some(str_lit(""))),
                    param("delete", Some(bool_lit(true))),
                ],
                vec![
                    set_this(
                        "name",
                        call_global(
                            "__py_temp_path",
                            vec![ident("prefix"), ident("suffix"), ident("dir")],
                        ),
                    ),
                    set_this("__fpath", this_field("name")),
                    set_this("__fmode", ident("mode")),
                    set_this("mode", ident("mode")),
                    set_this("delete", ident("delete")),
                    set_this("__fdata", str_lit("")),
                    set_this("closed", bool_lit(false)),
                    set_this("_pos", i(0)),
                ],
            ),
            method(
                "write",
                vec![param("data", Some(str_lit("")))],
                vec![
                    assign(
                        ident("__s"),
                        call_global("__py_tmp_text", vec![ident("data")]),
                    ),
                    expr_stmt(call_global(
                        "__py_fs_write_text",
                        vec![
                            path(),
                            op(
                                BinOp::Add,
                                call_global("__py_fs_read_text", vec![path()]),
                                ident("__s"),
                            ),
                        ],
                    )),
                    ret(call_global("len", vec![ident("__s")])),
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
            method(
                "read",
                vec![param("size", Some(i(-1)))],
                vec![ret(call_global("__py_fs_read_text", vec![path()]))],
            ),
            method(
                "readline",
                vec![],
                vec![
                    assign(
                        ident("__all"),
                        call_global("__py_fs_read_text", vec![path()]),
                    ),
                    assign(
                        ident("__rest"),
                        slice_from(ident("__all"), this_field("_pos")),
                    ),
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
                    set_this(
                        "_pos",
                        op(
                            BinOp::Add,
                            this_field("_pos"),
                            call_global("len", vec![ident("__r")]),
                        ),
                    ),
                    ret(ident("__r")),
                ],
            ),
            method(
                "readlines",
                vec![],
                vec![ret(call(
                    member(call_global("__py_fs_read_text", vec![path()]), "splitlines"),
                    vec![],
                ))],
            ),
            method(
                "seek",
                vec![param("pos", Some(i(0)))],
                vec![set_this("_pos", ident("pos"))],
            ),
            method("tell", vec![], vec![ret(this_field("_pos"))]),
            method("flush", vec![], vec![ret(null())]),
            method("fileno", vec![], vec![ret(i(3))]),
            method("readable", vec![], vec![ret(bool_lit(true))]),
            method("writable", vec![], vec![ret(bool_lit(true))]),
            method("seekable", vec![], vec![ret(bool_lit(true))]),
            // ⛔ `delete=True` removes the file on close, which is what the
            // corpus checks with `os.path.exists` afterwards.
            method(
                "close",
                vec![],
                vec![
                    if_stmt(
                        unary_not(this_field("closed")),
                        vec![
                            set_this("closed", bool_lit(true)),
                            if_stmt(
                                is_true(this_field("delete")),
                                vec![expr_stmt(call_global("__py_fs_unlink", vec![path()]))],
                            ),
                        ],
                    ),
                    ret(null()),
                ],
            ),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method(
                "__exit__",
                any_args(),
                vec![
                    expr_stmt(call(member(ident("self"), "close"), vec![])),
                    ret(bool_lit(false)),
                ],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        // ⛔ Its OWN helper, not `io`'s: `__py_io_text` only splices when the
        // program imports `io`, and a tempfile program need not.
        //
        // `str(x)` of a str is that same str, but of a bytes it is the REPR, so
        // the lengths separate them with no type test — `isinstance` does not
        // resolve inside a declaration.
        function(
            "__py_tmp_text",
            vec![param("value", Some(str_lit("")))],
            vec![
                if_stmt(is_none(ident("value")), vec![ret(str_lit(""))]),
                if_stmt(
                    op(
                        BinOp::Eq,
                        call_global("len", vec![call_global("str", vec![ident("value")])]),
                        call_global("len", vec![ident("value")]),
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
