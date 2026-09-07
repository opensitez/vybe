//! `filecmp` — file and directory comparison over shared filesystem primitives.

use super::builders::*;
use vybe_ast::{BinOp, Expression, Statement};

type Expr = Expression;

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn add(l: Expr, r: Expr) -> Expr {
    op(BinOp::Add, l, r)
}

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn path_join(left: Expr, name: Expr) -> Expr {
    call_global("__py_filecmp_join", vec![left, name])
}

fn append(list: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(list, "append"), vec![value]))
}

fn empty_dict() -> Expr {
    Expression::with_span(
        vybe_ast::ExprKind::Map(Vec::new()),
        vybe_ast::Span::default(),
    )
}

fn compute_dir_list(name: &str, cond: Expr, value: Expr) -> Statement {
    if_stmt(cond, vec![append(ident(name), value)])
}

pub(super) fn dircmp() -> Statement {
    class(
        "__PyDirCmp",
        vec![init(
            vec![
                param("left", None),
                param("right", None),
                param("ignore", Some(list_of(vec![]))),
                param("hide", Some(list_of(vec![]))),
            ],
            vec![
                set_this("left", ident("left")),
                set_this("right", ident("right")),
                set_this(
                    "left_list",
                    call_global(
                        "sorted",
                        vec![call_global("__py_fs_list_dir", vec![ident("left")])],
                    ),
                ),
                set_this(
                    "right_list",
                    call_global(
                        "sorted",
                        vec![call_global("__py_fs_list_dir", vec![ident("right")])],
                    ),
                ),
                assign(ident("__left_only"), list_of(vec![])),
                assign(ident("__right_only"), list_of(vec![])),
                assign(ident("__common"), list_of(vec![])),
                assign(ident("__common_files"), list_of(vec![])),
                assign(ident("__common_dirs"), list_of(vec![])),
                assign(ident("__common_funny"), list_of(vec![])),
                assign(ident("__same"), list_of(vec![])),
                assign(ident("__diff"), list_of(vec![])),
                assign(ident("__subdirs"), empty_dict()),
                for_in(
                    "__name",
                    this_field("left_list"),
                    vec![
                        if_stmt(
                            contains(this_field("right_list"), ident("__name")),
                            vec![append(ident("__common"), ident("__name"))],
                        ),
                        if_stmt(
                            unary_not(contains(this_field("right_list"), ident("__name"))),
                            vec![append(ident("__left_only"), ident("__name"))],
                        ),
                    ],
                ),
                for_in(
                    "__name",
                    this_field("right_list"),
                    vec![if_stmt(
                        unary_not(contains(this_field("left_list"), ident("__name"))),
                        vec![append(ident("__right_only"), ident("__name"))],
                    )],
                ),
                for_in(
                    "__name",
                    ident("__common"),
                    vec![
                        assign(ident("__lp"), path_join(ident("left"), ident("__name"))),
                        assign(ident("__rp"), path_join(ident("right"), ident("__name"))),
                        assign(
                            ident("__lf"),
                            call_global("__py_fs_is_file", vec![ident("__lp")]),
                        ),
                        assign(
                            ident("__rf"),
                            call_global("__py_fs_is_file", vec![ident("__rp")]),
                        ),
                        assign(
                            ident("__ld"),
                            call_global("__py_fs_is_dir", vec![ident("__lp")]),
                        ),
                        assign(
                            ident("__rd"),
                            call_global("__py_fs_is_dir", vec![ident("__rp")]),
                        ),
                        compute_dir_list(
                            "__common_files",
                            op(BinOp::And, ident("__lf"), ident("__rf")),
                            ident("__name"),
                        ),
                        compute_dir_list(
                            "__common_dirs",
                            op(BinOp::And, ident("__ld"), ident("__rd")),
                            ident("__name"),
                        ),
                        compute_dir_list(
                            "__common_funny",
                            op(
                                BinOp::Or,
                                op(BinOp::And, ident("__lf"), ident("__rd")),
                                op(BinOp::And, ident("__ld"), ident("__rf")),
                            ),
                            ident("__name"),
                        ),
                    ],
                ),
                for_in(
                    "__name",
                    ident("__common_files"),
                    vec![
                        assign(ident("__lp"), path_join(ident("left"), ident("__name"))),
                        assign(ident("__rp"), path_join(ident("right"), ident("__name"))),
                        if_stmt(
                            call_global("cmp", vec![ident("__lp"), ident("__rp"), bool_lit(false)]),
                            vec![append(ident("__same"), ident("__name"))],
                        ),
                        if_stmt(
                            unary_not(call_global(
                                "cmp",
                                vec![ident("__lp"), ident("__rp"), bool_lit(false)],
                            )),
                            vec![append(ident("__diff"), ident("__name"))],
                        ),
                    ],
                ),
                for_in(
                    "__name",
                    ident("__common_dirs"),
                    vec![assign(
                        index(ident("__subdirs"), ident("__name")),
                        bool_lit(true),
                    )],
                ),
                set_this("left_only", ident("__left_only")),
                set_this("right_only", ident("__right_only")),
                set_this("common", ident("__common")),
                set_this("common_files", ident("__common_files")),
                set_this("common_dirs", ident("__common_dirs")),
                set_this("common_funny", ident("__common_funny")),
                set_this("same_files", ident("__same")),
                set_this("diff_files", ident("__diff")),
                set_this("funny_files", list_of(vec![])),
                set_this("subdirs", ident("__subdirs")),
            ],
        )],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "__py_filecmp_join",
            vec![param("left", None), param("name", None)],
            vec![ret(add(add(ident("left"), str_lit("/")), ident("name")))],
        ),
        function(
            "cmp",
            vec![
                param("f1", None),
                param("f2", None),
                param("shallow", Some(bool_lit(true))),
            ],
            vec![
                if_stmt(
                    op(BinOp::Eq, ident("f1"), ident("f2")),
                    vec![ret(bool_lit(true))],
                ),
                if_stmt(
                    unary_not(call_global("__py_fs_exists", vec![ident("f1")])),
                    vec![ret(bool_lit(false))],
                ),
                if_stmt(
                    unary_not(call_global("__py_fs_exists", vec![ident("f2")])),
                    vec![ret(bool_lit(false))],
                ),
                ret(op(
                    BinOp::Eq,
                    call_global("__py_fs_read_bytes", vec![ident("f1")]),
                    call_global("__py_fs_read_bytes", vec![ident("f2")]),
                )),
            ],
        ),
        function(
            "cmpfiles",
            vec![
                param("dir1", None),
                param("dir2", None),
                param("common", None),
                param("shallow", Some(bool_lit(true))),
            ],
            vec![
                assign(ident("__match"), list_of(vec![])),
                assign(ident("__mismatch"), list_of(vec![])),
                assign(ident("__errors"), list_of(vec![])),
                for_in(
                    "__name",
                    ident("common"),
                    vec![
                        assign(ident("__p1"), path_join(ident("dir1"), ident("__name"))),
                        assign(ident("__p2"), path_join(ident("dir2"), ident("__name"))),
                        if_stmt(
                            op(
                                BinOp::Or,
                                unary_not(call_global("__py_fs_exists", vec![ident("__p1")])),
                                unary_not(call_global("__py_fs_exists", vec![ident("__p2")])),
                            ),
                            vec![append(ident("__errors"), ident("__name"))],
                        ),
                        if_stmt(
                            op(
                                BinOp::And,
                                call_global("__py_fs_exists", vec![ident("__p1")]),
                                call_global("__py_fs_exists", vec![ident("__p2")]),
                            ),
                            vec![
                                if_stmt(
                                    call_global(
                                        "cmp",
                                        vec![ident("__p1"), ident("__p2"), ident("shallow")],
                                    ),
                                    vec![append(ident("__match"), ident("__name"))],
                                ),
                                if_stmt(
                                    unary_not(call_global(
                                        "cmp",
                                        vec![ident("__p1"), ident("__p2"), ident("shallow")],
                                    )),
                                    vec![append(ident("__mismatch"), ident("__name"))],
                                ),
                            ],
                        ),
                    ],
                ),
                ret(tuple_of(vec![
                    ident("__match"),
                    ident("__mismatch"),
                    ident("__errors"),
                ])),
            ],
        ),
        function(
            "dircmp",
            vec![
                param("a", None),
                param("b", None),
                param("ignore", Some(list_of(vec![]))),
                param("hide", Some(list_of(vec![]))),
            ],
            vec![ret(new(
                "__PyDirCmp",
                vec![ident("a"), ident("b"), ident("ignore"), ident("hide")],
            ))],
        ),
        function("clear_cache", vec![], vec![ret(null())]),
    ]
}
