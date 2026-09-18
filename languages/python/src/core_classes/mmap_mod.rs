//! `mmap` — in-memory mapping adapter over Python file objects.
//!
//! The host surface does not expose POSIX memory maps. For Python-over-WASM the
//! useful contract is the Python object behavior: bytes-like reads, indexing,
//! seek/tell, copy-on-write vs write-through, and anonymous maps. File-backed
//! maps consume the file object token returned by `fileno()` and flush through
//! the existing filesystem primitives.

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

fn to_text(e: Expr) -> Expr {
    call_global("__py_mmap_text", vec![e])
}

fn to_bytes(e: Expr) -> Expr {
    call_global("bytes", vec![e, str_lit("utf-8")])
}

fn key_start(key: Expr) -> Expr {
    ternary(
        is_none(field_of(key.clone(), "start")),
        i(0),
        field_of(key, "start"),
    )
}

fn key_stop(key: Expr) -> Expr {
    ternary(
        is_none(field_of(key.clone(), "stop")),
        len_of(buf()),
        field_of(key, "stop"),
    )
}

fn has_start(key: Expr) -> Expr {
    call_global("hasattr", vec![key, str_lit("start")])
}

pub(super) fn mmap_class() -> Statement {
    class(
        "mmap",
        vec![
            init(
                vec![
                    param("fileno", None),
                    param("length", Some(i(0))),
                    param("flags", Some(i(0))),
                    param("prot", Some(i(0))),
                    param("access", Some(i(0))),
                    param("offset", Some(i(0))),
                ],
                vec![
                    set_this("_fd", ident("fileno")),
                    set_this("_access", ident("access")),
                    set_this("_pos", i(0)),
                    set_this("closed", bool_lit(false)),
                    set_this(
                        "_buf",
                        call_global("__py_mmap_initial", vec![ident("fileno"), ident("length")]),
                    ),
                ],
            ),
            method("__len__", vec![], vec![ret(len_of(buf()))]),
            method(
                "__getitem__",
                vec![param("key", None)],
                vec![
                    if_stmt(
                        has_start(ident("key")),
                        vec![
                            assign(ident("__start"), key_start(ident("key"))),
                            assign(ident("__stop"), key_stop(ident("key"))),
                            ret(to_bytes(slice_range(
                                buf(),
                                ident("__start"),
                                ident("__stop"),
                            ))),
                        ],
                    ),
                    ret(call_global("ord", vec![index(buf(), ident("key"))])),
                ],
            ),
            method(
                "__getslice__",
                vec![param("start", Some(null())), param("stop", Some(null()))],
                vec![
                    assign(ident("__start"), ident("start")),
                    if_stmt(
                        is_none(ident("__start")),
                        vec![assign(ident("__start"), i(0))],
                    ),
                    assign(ident("__stop"), ident("stop")),
                    if_stmt(
                        is_none(ident("__stop")),
                        vec![assign(ident("__stop"), len_of(buf()))],
                    ),
                    ret(to_bytes(slice_range(
                        buf(),
                        ident("__start"),
                        ident("__stop"),
                    ))),
                ],
            ),
            method(
                "__setitem__",
                vec![param("key", None), param("value", None)],
                vec![
                    if_stmt(
                        op(BinOp::Eq, this_field("_access"), i(1)),
                        vec![raise_new(
                            "TypeError",
                            vec![str_lit("mmap can't modify a readonly memory map")],
                        )],
                    ),
                    assign(ident("__s"), to_text(ident("value"))),
                    if_stmt(
                        has_start(ident("key")),
                        vec![
                            assign(ident("__start"), key_start(ident("key"))),
                            assign(ident("__stop"), key_stop(ident("key"))),
                            set_this(
                                "_buf",
                                op(
                                    BinOp::Add,
                                    op(
                                        BinOp::Add,
                                        slice_range(buf(), i(0), ident("__start")),
                                        ident("__s"),
                                    ),
                                    slice_from(buf(), ident("__stop")),
                                ),
                            ),
                            expr_stmt(call(member(ident("self"), "flush"), vec![])),
                            ret(null()),
                        ],
                    ),
                    set_this(
                        "_buf",
                        op(
                            BinOp::Add,
                            op(
                                BinOp::Add,
                                slice_range(buf(), i(0), ident("key")),
                                ident("__s"),
                            ),
                            slice_from(buf(), op(BinOp::Add, ident("key"), i(1))),
                        ),
                    ),
                    expr_stmt(call(member(ident("self"), "flush"), vec![])),
                    ret(null()),
                ],
            ),
            method(
                "__setslice__",
                vec![
                    param("start", Some(null())),
                    param("stop", Some(null())),
                    param("value", None),
                ],
                vec![
                    if_stmt(
                        op(BinOp::Eq, this_field("_access"), i(1)),
                        vec![raise_new(
                            "TypeError",
                            vec![str_lit("mmap can't modify a readonly memory map")],
                        )],
                    ),
                    assign(ident("__start"), ident("start")),
                    if_stmt(
                        is_none(ident("__start")),
                        vec![assign(ident("__start"), i(0))],
                    ),
                    assign(ident("__stop"), ident("stop")),
                    if_stmt(
                        is_none(ident("__stop")),
                        vec![assign(ident("__stop"), len_of(buf()))],
                    ),
                    assign(ident("__s"), to_text(ident("value"))),
                    set_this(
                        "_buf",
                        op(
                            BinOp::Add,
                            op(
                                BinOp::Add,
                                slice_range(buf(), i(0), ident("__start")),
                                ident("__s"),
                            ),
                            slice_from(buf(), ident("__stop")),
                        ),
                    ),
                    expr_stmt(call(member(ident("self"), "flush"), vec![])),
                    ret(null()),
                ],
            ),
            method(
                "read",
                vec![param("size", Some(i(-1)))],
                vec![
                    assign(ident("__end"), len_of(buf())),
                    if_stmt(
                        op(BinOp::GtEq, ident("size"), i(0)),
                        vec![assign(ident("__end"), op(BinOp::Add, pos(), ident("size")))],
                    ),
                    assign(ident("__r"), slice_range(buf(), pos(), ident("__end"))),
                    set_this("_pos", ident("__end")),
                    ret(to_bytes(ident("__r"))),
                ],
            ),
            method(
                "readline",
                vec![],
                vec![
                    assign(ident("__rest"), slice_from(buf(), pos())),
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
                    ret(to_bytes(ident("__r"))),
                ],
            ),
            method(
                "write",
                vec![param("data", None)],
                vec![
                    assign(ident("__s"), to_text(ident("data"))),
                    set_this(
                        "_buf",
                        op(
                            BinOp::Add,
                            op(BinOp::Add, slice_range(buf(), i(0), pos()), ident("__s")),
                            slice_from(buf(), op(BinOp::Add, pos(), len_of(ident("__s")))),
                        ),
                    ),
                    set_this("_pos", op(BinOp::Add, pos(), len_of(ident("__s")))),
                    expr_stmt(call(member(ident("self"), "flush"), vec![])),
                    ret(len_of(ident("__s"))),
                ],
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
            method("size", vec![], vec![ret(len_of(buf()))]),
            method(
                "find",
                vec![
                    param("sub", None),
                    param("start", Some(i(0))),
                    param("end", Some(null())),
                ],
                vec![
                    assign(ident("__end"), ident("end")),
                    if_stmt(
                        is_none(ident("__end")),
                        vec![assign(ident("__end"), len_of(buf()))],
                    ),
                    ret(call(
                        member(slice_range(buf(), ident("start"), ident("__end")), "find"),
                        vec![to_text(ident("sub"))],
                    )),
                ],
            ),
            method(
                "rfind",
                vec![
                    param("sub", None),
                    param("start", Some(i(0))),
                    param("end", Some(null())),
                ],
                vec![
                    assign(ident("__end"), ident("end")),
                    if_stmt(
                        is_none(ident("__end")),
                        vec![assign(ident("__end"), len_of(buf()))],
                    ),
                    ret(call(
                        member(slice_range(buf(), ident("start"), ident("__end")), "rfind"),
                        vec![to_text(ident("sub"))],
                    )),
                ],
            ),
            method(
                "move",
                vec![
                    param("dest", None),
                    param("src", None),
                    param("count", None),
                ],
                vec![
                    assign(
                        ident("__frag"),
                        slice_range(
                            buf(),
                            ident("src"),
                            op(BinOp::Add, ident("src"), ident("count")),
                        ),
                    ),
                    set_this(
                        "_buf",
                        op(
                            BinOp::Add,
                            op(
                                BinOp::Add,
                                slice_range(buf(), i(0), ident("dest")),
                                ident("__frag"),
                            ),
                            slice_from(buf(), op(BinOp::Add, ident("dest"), ident("count"))),
                        ),
                    ),
                    expr_stmt(call(member(ident("self"), "flush"), vec![])),
                    ret(null()),
                ],
            ),
            method(
                "flush",
                vec![],
                vec![
                    if_stmt(
                        op(BinOp::NotEq, this_field("_access"), i(3)),
                        vec![expr_stmt(call_global(
                            "__py_mmap_flush",
                            vec![this_field("_fd"), buf()],
                        ))],
                    ),
                    ret(i(0)),
                ],
            ),
            method(
                "close",
                vec![],
                vec![
                    expr_stmt(call(member(ident("self"), "flush"), vec![])),
                    set_this("closed", bool_lit(true)),
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
        function(
            "__py_mmap_text",
            vec![param("value", Some(str_lit("")))],
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
        function(
            "__py_mmap_initial",
            vec![param("fileno", None), param("length", Some(i(0)))],
            vec![
                if_stmt(
                    op(BinOp::Eq, ident("fileno"), i(-1)),
                    vec![ret(str_lit(""))],
                ),
                if_stmt(
                    call_global("hasattr", vec![ident("fileno"), str_lit("__fpath")]),
                    vec![ret(call_global(
                        "__py_fs_read_text",
                        vec![member(ident("fileno"), "__fpath")],
                    ))],
                ),
                ret(to_text(call_global(
                    "__py_file_read",
                    vec![ident("fileno")],
                ))),
            ],
        ),
        function(
            "__py_mmap_flush",
            vec![param("fileno", None), param("data", Some(str_lit("")))],
            vec![
                if_stmt(op(BinOp::Eq, ident("fileno"), i(-1)), vec![ret(null())]),
                if_stmt(
                    call_global("hasattr", vec![ident("fileno"), str_lit("__fpath")]),
                    vec![
                        assign(member(ident("fileno"), "__fdata"), to_bytes(ident("data"))),
                        expr_stmt(call_global(
                            "__py_fs_write_text",
                            vec![member(ident("fileno"), "__fpath"), ident("data")],
                        )),
                    ],
                ),
                ret(null()),
            ],
        ),
    ]
}
