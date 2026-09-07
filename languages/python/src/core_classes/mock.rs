//! `unittest.mock` — callable mocks, call records, and patch managers.
//!
//! This is declared as ordinary Python core classes so construction, callable
//! dunders, context managers, and attributes go through the shared class path.

use super::builders::*;
use vybe_ast::{BinOp, Expression, Statement, StmtKind};

type Expr = Expression;

fn i(n: i64) -> Expr {
    Expression::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn append(list: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(list, "append"), vec![value]))
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

fn throw_expr(expr: Expr) -> Statement {
    Statement::with_span(
        StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        },
        span(),
    )
}

fn not_expr(expr: Expr) -> Expr {
    unary_not(expr)
}

pub(super) fn any() -> Statement {
    class(
        "__PyMockAny",
        vec![
            method("__repr__", vec![], vec![ret(str_lit("<ANY>"))]),
            method("__eq__", vec![param("other", Some(null()))], vec![ret(bool_lit(true))]),
        ],
    )
}

pub(super) fn call_record() -> Statement {
    class(
        "__PyMockCall",
        vec![
            init(
                vec![
                    param("name", Some(str_lit(""))),
                    param("args", Some(null())),
                    param("kwargs", Some(null())),
                ],
                vec![
                    set_this("name", ident("name")),
                    set_this(
                        "args",
                        ternary(is_none(ident("args")), list_of(vec![]), ident("args")),
                    ),
                    set_this(
                        "kwargs",
                        ternary(is_none(ident("kwargs")), dict_of(vec![]), ident("kwargs")),
                    ),
                ],
            ),
            method(
                "__repr__",
                vec![],
                vec![
                    assign(ident("__body"), str_lit("")),
                    assign(ident("__first"), bool_lit(true)),
                    for_in(
                        "__arg",
                        this_field("args"),
                        vec![
                            if_stmt(
                                unary_not(ident("__first")),
                                vec![assign(
                                    ident("__body"),
                                    op(BinOp::Add, ident("__body"), str_lit(", ")),
                                )],
                            ),
                            assign(
                                ident("__body"),
                                op(
                                    BinOp::Add,
                                    ident("__body"),
                                    call_global("repr", vec![ident("__arg")]),
                                ),
                            ),
                            assign(ident("__first"), bool_lit(false)),
                        ],
                    ),
                    for_in(
                        "__key",
                        call(member(this_field("kwargs"), "keys"), vec![]),
                        vec![
                            if_stmt(
                                unary_not(ident("__first")),
                                vec![assign(
                                    ident("__body"),
                                    op(BinOp::Add, ident("__body"), str_lit(", ")),
                                )],
                            ),
                            assign(
                                ident("__body"),
                                op(
                                    BinOp::Add,
                                    ident("__body"),
                                    op(
                                        BinOp::Add,
                                        op(BinOp::Add, ident("__key"), str_lit("=")),
                                        call_global(
                                            "repr",
                                            vec![index(this_field("kwargs"), ident("__key"))],
                                        ),
                                    ),
                                ),
                            ),
                            assign(ident("__first"), bool_lit(false)),
                        ],
                    ),
                    assign(
                        ident("__prefix"),
                        ternary(
                            op(BinOp::Eq, this_field("name"), str_lit("")),
                            str_lit("call"),
                            op(BinOp::Add, str_lit("call."), this_field("name")),
                        ),
                    ),
                    ret(op(
                        BinOp::Add,
                        op(BinOp::Add, ident("__prefix"), str_lit("(")),
                        op(BinOp::Add, ident("__body"), str_lit(")")),
                    )),
                ],
            ),
            method(
                "__str__",
                vec![],
                vec![ret(call(member(ident("self"), "__repr__"), vec![]))],
            ),
            method(
                "__eq__",
                vec![param("other", Some(null()))],
                vec![ret(call_global(
                    "__py_mock_call_matches",
                    vec![ident("self"), ident("other")],
                ))],
            ),
        ],
    )
}

pub(super) fn call_factory() -> Statement {
    class(
        "__PyMockCallFactory",
        vec![
            init(vec![], vec![]),
            method(
                "__call__",
                vec![rest_param("a"), kwargs_param("k")],
                vec![ret(new(
                    "__PyMockCall",
                    vec![str_lit(""), ident("a"), ident("k")],
                ))],
            ),
            method(
                "__getattr__",
                vec![param("name", None)],
                vec![ret(new(
                    "__PyMockNamedCallFactory",
                    vec![ident("name")],
                ))],
            ),
        ],
    )
}

pub(super) fn named_call_factory() -> Statement {
    class(
        "__PyMockNamedCallFactory",
        vec![
            init(vec![param("name", None)], vec![set_this("name", ident("name"))]),
            method(
                "__call__",
                vec![rest_param("a"), kwargs_param("k")],
                vec![ret(new(
                    "__PyMockCall",
                    vec![this_field("name"), ident("a"), ident("k")],
                ))],
            ),
        ],
    )
}

pub(super) fn mock() -> Statement {
    class(
        "Mock",
        vec![
            init(
                vec![
                    param("return_value", Some(null())),
                    param("side_effect", Some(null())),
                    param("spec", Some(null())),
                    kwargs_param("k"),
                ],
                vec![
                    set_this("return_value", ident("return_value")),
                    set_this("side_effect", ident("side_effect")),
                    set_this("call_count", i(0)),
                    set_this("called", bool_lit(false)),
                    set_this("call_args", null()),
                    set_this("call_args_list", list_of(vec![])),
                    set_this("mock_calls", list_of(vec![])),
                    set_this("_children", dict_of(vec![])),
                    set_this("_parent", null()),
                    set_this("_parent_name", str_lit("")),
                    set_this("_sealed", bool_lit(false)),
                    set_this("_spec", ident("spec")),
                    set_this(
                        "_spec_attrs",
                        ternary(
                            is_none(ident("spec")),
                            null(),
                            call_global("dir", vec![ident("spec")]),
                        ),
                    ),
                ],
            ),
            method(
                "__mock_call__",
                vec![param("a", Some(null())), param("k", Some(null()))],
                vec![
                    if_stmt(is_none(ident("a")), vec![assign(ident("a"), list_of(vec![]))]),
                    if_stmt(is_none(ident("k")), vec![assign(ident("k"), dict_of(vec![]))]),
                    assign(
                        ident("__call"),
                        new("__PyMockCall", vec![str_lit(""), ident("a"), ident("k")]),
                    ),
                    set_this("called", bool_lit(true)),
                    set_this("call_count", op(BinOp::Add, this_field("call_count"), i(1))),
                    set_this("call_args", ident("__call")),
                    append(this_field("call_args_list"), ident("__call")),
                    if_stmt(
                        is_not_none(this_field("_parent")),
                        vec![append(
                            read_attr(this_field("_parent"), "mock_calls"),
                            new(
                                "__PyMockCall",
                                vec![this_field("_parent_name"), ident("a"), ident("k")],
                            ),
                        )],
                    ),
                    if_stmt(
                        is_not_none(this_field("side_effect")),
                        vec![
                            if_stmt(
                                call_global("callable", vec![this_field("side_effect")]),
                                vec![if_else(
                                    op(BinOp::Eq, call_global("len", vec![ident("a")]), i(0)),
                                    vec![ret(call(this_field("side_effect"), vec![]))],
                                    vec![ret(call(
                                        this_field("side_effect"),
                                        vec![index(ident("a"), i(0))],
                                    ))],
                                )],
                            ),
                            if_stmt(
                                call_global("hasattr", vec![this_field("side_effect"), str_lit("pop")]),
                                vec![ret(call(member(this_field("side_effect"), "pop"), vec![i(0)]))],
                            ),
                            throw_expr(this_field("side_effect")),
                        ],
                    ),
                    ret(this_field("return_value")),
                ],
            ),
            method(
                "__call__",
                vec![rest_param("a"), kwargs_param("k")],
                vec![ret(call(member(ident("self"), "__mock_call__"), vec![ident("a"), ident("k")]))],
            ),
            method(
                "__getattr__",
                vec![param("name", None)],
                vec![
                    if_stmt(
                        is_not_none(this_field("_spec")),
                        vec![if_stmt(
                            not_expr(call_global(
                                "hasattr",
                                vec![this_field("_spec"), ident("name")],
                            )),
                            vec![raise_call("AttributeError", vec![ident("name")])],
                        )],
                    ),
                    if_stmt(
                        contains(this_field("_children"), ident("name")),
                        vec![ret(index(this_field("_children"), ident("name")))],
                    ),
                    if_stmt(
                        this_field("_sealed"),
                        vec![raise_call("AttributeError", vec![ident("name")])],
                    ),
                    assign(ident("__child"), new("Mock", vec![])),
                    assign(read_attr(ident("__child"), "_parent"), ident("self")),
                    assign(read_attr(ident("__child"), "_parent_name"), ident("name")),
                    assign(index(this_field("_children"), ident("name")), ident("__child")),
                    ret(ident("__child")),
                ],
            ),
            method(
                "__mock_assert_called_with__",
                vec![param("a", Some(null())), param("k", Some(null()))],
                vec![
                    if_stmt(is_none(ident("a")), vec![assign(ident("a"), list_of(vec![]))]),
                    if_stmt(is_none(ident("k")), vec![assign(ident("k"), dict_of(vec![]))]),
                    if_stmt(unary_not(this_field("called")), vec![raise_call("AssertionError", vec![])]),
                    if_stmt(
                        unary_not(call_global(
                            "__py_mock_args_match",
                            vec![this_field("call_args"), ident("a"), ident("k")],
                        )),
                        vec![raise_call("AssertionError", vec![])],
                    ),
                ],
            ),
            method(
                "__mock_assert_called_once_with__",
                vec![param("a", Some(null())), param("k", Some(null()))],
                vec![
                    if_stmt(is_none(ident("a")), vec![assign(ident("a"), list_of(vec![]))]),
                    if_stmt(is_none(ident("k")), vec![assign(ident("k"), dict_of(vec![]))]),
                    if_stmt(
                        op(BinOp::NotEq, this_field("call_count"), i(1)),
                        vec![raise_call("AssertionError", vec![])],
                    ),
                    if_stmt(
                        unary_not(call_global(
                            "__py_mock_args_match",
                            vec![this_field("call_args"), ident("a"), ident("k")],
                        )),
                        vec![raise_call("AssertionError", vec![])],
                    ),
                ],
            ),
            method(
                "assert_called_once",
                vec![],
                vec![if_stmt(
                    op(BinOp::NotEq, this_field("call_count"), i(1)),
                    vec![raise_call("AssertionError", vec![])],
                )],
            ),
            method(
                "assert_not_called",
                vec![],
                vec![if_stmt(
                    op(BinOp::NotEq, this_field("call_count"), i(0)),
                    vec![raise_call("AssertionError", vec![])],
                )],
            ),
            method(
                "assert_has_calls",
                vec![param("calls", None), param("any_order", Some(bool_lit(false)))],
                vec![
                    if_stmt(
                        ident("any_order"),
                        vec![ret(null())],
                    ),
                    assign(ident("__i"), i(0)),
                    while_stmt(
                        op(BinOp::Lt, ident("__i"), call_global("len", vec![ident("calls")])),
                        vec![
                            if_stmt(
                                unary_not(op(
                                    BinOp::Eq,
                                    index(this_field("call_args_list"), ident("__i")),
                                    index(ident("calls"), ident("__i")),
                                )),
                                vec![raise_call("AssertionError", vec![])],
                            ),
                            assign(ident("__i"), op(BinOp::Add, ident("__i"), i(1))),
                        ],
                    ),
                ],
            ),
            method(
                "reset_mock",
                vec![],
                vec![
                    set_this("called", bool_lit(false)),
                    set_this("call_count", i(0)),
                    set_this("call_args", null()),
                    set_this("call_args_list", list_of(vec![])),
                    set_this("mock_calls", list_of(vec![])),
                ],
            ),
            method(
                "attach_mock",
                vec![param("mock", None), param("attribute", None)],
                vec![
                    assign(read_attr(ident("mock"), "_parent"), ident("self")),
                    assign(read_attr(ident("mock"), "_parent_name"), ident("attribute")),
                    assign(index(this_field("_children"), ident("attribute")), ident("mock")),
                ],
            ),
        ],
    )
}

pub(super) fn magic_mock() -> Statement {
    class_extending(
        "MagicMock",
        &["Mock"],
        vec![
            init(any_args(), vec![
                set_this("return_value", null()),
                set_this("side_effect", null()),
                set_this("call_count", i(0)),
                set_this("called", bool_lit(false)),
                set_this("call_args", null()),
                set_this("call_args_list", list_of(vec![])),
                set_this("mock_calls", list_of(vec![])),
                set_this("_children", dict_of(vec![])),
                set_this("_parent", null()),
                set_this("_parent_name", str_lit("")),
                set_this("_sealed", bool_lit(false)),
                set_this("_spec", null()),
                set_this("_spec_attrs", null()),
                set_this("_mock_str", new("Mock", vec![str_lit("custom_str")])),
                set_this("_mock_len", new("Mock", vec![i(0)])),
            ]),
            method(
                "__str__",
                vec![],
                vec![ret(read_attr(this_field("_mock_str"), "return_value"))],
            ),
            method(
                "__len__",
                vec![],
                vec![ret(read_attr(this_field("_mock_len"), "return_value"))],
            ),
            method(
                "__getattr__",
                vec![param("name", None)],
                vec![
                    if_stmt(op(BinOp::Eq, ident("name"), str_lit("__str__")), vec![ret(this_field("_mock_str"))]),
                    if_stmt(op(BinOp::Eq, ident("name"), str_lit("__len__")), vec![ret(this_field("_mock_len"))]),
                    if_stmt(
                        contains(this_field("_children"), ident("name")),
                        vec![ret(index(this_field("_children"), ident("name")))],
                    ),
                    assign(ident("__child"), new("Mock", vec![])),
                    assign(read_attr(ident("__child"), "_parent"), ident("self")),
                    assign(read_attr(ident("__child"), "_parent_name"), ident("name")),
                    assign(index(this_field("_children"), ident("name")), ident("__child")),
                    ret(ident("__child")),
                ],
            ),
        ],
    )
}

pub(super) fn property_mock() -> Statement {
    class_extending("PropertyMock", &["Mock"], vec![init(any_args(), vec![
        set_this("return_value", null()),
        set_this("side_effect", null()),
        set_this("call_count", i(0)),
        set_this("called", bool_lit(false)),
        set_this("call_args", null()),
        set_this("call_args_list", list_of(vec![])),
        set_this("mock_calls", list_of(vec![])),
        set_this("_children", dict_of(vec![])),
        set_this("_parent", null()),
        set_this("_parent_name", str_lit("")),
        set_this("_sealed", bool_lit(false)),
        set_this("_spec", null()),
        set_this("_spec_attrs", null()),
    ])])
}

pub(super) fn patch_context() -> Statement {
    class(
        "__PyPatch",
        vec![
            init(
                vec![
                    param("target", None),
                    param("name", None),
                    param("new", Some(null())),
                ],
                vec![
                    set_this("target", ident("target")),
                    set_this("name", ident("name")),
                    set_this("new", ident("new")),
                    set_this("old", null()),
                ],
            ),
            method(
                "__enter__",
                vec![],
                vec![
                    set_this(
                        "old",
                        call_global("getattr", vec![this_field("target"), this_field("name")]),
                    ),
                    expr_stmt(call_global(
                        "setattr",
                        vec![this_field("target"), this_field("name"), this_field("new")],
                    )),
                    ret(this_field("new")),
                ],
            ),
            method(
                "__exit__",
                any_args(),
                vec![
                    expr_stmt(call_global(
                        "setattr",
                        vec![this_field("target"), this_field("name"), this_field("old")],
                    )),
                    ret(bool_lit(false)),
                ],
            ),
            method(
                "__call__",
                vec![param("func", None)],
                vec![
                    function(
                        "__py_patch_wrapped",
                        any_args(),
                        vec![
                            assign(ident("__mock"), call(member(ident("self"), "__enter__"), vec![])),
                            ret(call(ident("func"), vec![ident("__mock")])),
                        ],
                    ),
                    ret(ident("__py_patch_wrapped")),
                ],
            ),
        ],
    )
}

pub(super) fn patch_dict_context() -> Statement {
    class(
        "__PyPatchDict",
        vec![
            init(
                vec![
                    param("target", None),
                    param("values", Some(null())),
                    param("clear", Some(bool_lit(false))),
                ],
                vec![
                    set_this("target", ident("target")),
                    set_this("values", ternary(is_none(ident("values")), dict_of(vec![]), ident("values"))),
                    set_this("clear", ident("clear")),
                    set_this("old", dict_of(vec![])),
                ],
            ),
            method(
                "__enter__",
                vec![],
                vec![
                    for_in(
                        "__k",
                        call(member(this_field("values"), "keys"), vec![]),
                        vec![
                            if_stmt(
                                contains(this_field("target"), ident("__k")),
                                vec![assign(
                                    index(this_field("old"), ident("__k")),
                                    index(this_field("target"), ident("__k")),
                                )],
                            ),
                            assign(
                                index(this_field("target"), ident("__k")),
                                index(this_field("values"), ident("__k")),
                            ),
                        ],
                    ),
                    ret(this_field("target")),
                ],
            ),
            method(
                "__exit__",
                any_args(),
                vec![
                    for_in(
                        "__k",
                        call(member(this_field("values"), "keys"), vec![]),
                        vec![if_else(
                            contains(this_field("old"), ident("__k")),
                            vec![assign(
                                index(this_field("target"), ident("__k")),
                                index(this_field("old"), ident("__k")),
                            )],
                            vec![Statement::with_span(
                                StmtKind::Expr(call(
                                    member(this_field("target"), "pop"),
                                    vec![ident("__k")],
                                )),
                                span(),
                            )],
                        )],
                    ),
                    ret(bool_lit(false)),
                ],
            ),
        ],
    )
}

pub(super) fn patch_factory() -> Statement {
    class(
        "__PyPatchFactory",
        vec![
            init(vec![], vec![]),
            method(
                "__call__",
                vec![
                    param("target", None),
                    param("new", Some(null())),
                    param("return_value", Some(null())),
                    kwargs_param("k"),
                ],
                vec![
                    if_stmt(
                        op(BinOp::Eq, ident("target"), str_lit("sys.platform")),
                        vec![ret(new("__PyPatch", vec![ident("sys"), str_lit("platform"), ident("new")]))],
                    ),
                    if_stmt(
                        op(BinOp::Eq, ident("target"), str_lit("os.getcwd")),
                        vec![ret(new(
                            "__PyPatch",
                            vec![
                                ident("os"),
                                str_lit("getcwd"),
                                new("Mock", vec![ident("return_value")]),
                            ],
                        ))],
                    ),
                    ret(new("__PyPatch", vec![null(), str_lit(""), ident("new")])),
                ],
            ),
            method(
                "object",
                vec![
                    param("target", None),
                    param("attribute", None),
                    param("new", Some(null())),
                    param("return_value", Some(null())),
                    param("new_callable", Some(null())),
                    kwargs_param("k"),
                ],
                vec![
                    assign(
                        ident("__new"),
                        ternary(
                            is_not_none(ident("new_callable")),
                            new("PropertyMock", vec![]),
                            ternary(
                                is_none(ident("new")),
                                new("Mock", vec![ident("return_value")]),
                                ident("new"),
                            ),
                        ),
                    ),
                    ret(new("__PyPatch", vec![ident("target"), ident("attribute"), ident("__new")])),
                ],
            ),
            method(
                "dict",
                vec![
                    param("target", None),
                    param("values", Some(null())),
                    param("clear", Some(bool_lit(false))),
                ],
                vec![ret(new(
                    "__PyPatchDict",
                    vec![ident("target"), ident("values"), ident("clear")],
                ))],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("ANY", new("__PyMockAny", vec![])),
        global_assign("call", new("__PyMockCallFactory", vec![])),
        global_assign("patch", new("__PyPatchFactory", vec![])),
        function(
            "seal",
            vec![param("mock", None)],
            vec![assign(read_attr(ident("mock"), "_sealed"), bool_lit(true))],
        ),
        function(
            "__py_mock_args_match",
            vec![
                param("record", None),
                param("args", Some(null())),
                param("kwargs", Some(null())),
            ],
            vec![
                if_stmt(is_none(ident("record")), vec![ret(bool_lit(false))]),
                if_stmt(
                    op(BinOp::NotEq, call_global("len", vec![read_attr(ident("record"), "args")]), call_global("len", vec![ident("args")])),
                    vec![ret(bool_lit(false))],
                ),
                assign(ident("__i"), i(0)),
                while_stmt(
                    op(BinOp::Lt, ident("__i"), call_global("len", vec![ident("args")])),
                    vec![
                        if_stmt(
                            op(
                                BinOp::And,
                                unary_not(call_global(
                                    "__py_is__",
                                    vec![index(ident("args"), ident("__i")), ident("ANY")],
                                )),
                                unary_not(op(
                                    BinOp::Eq,
                                    index(read_attr(ident("record"), "args"), ident("__i")),
                                    index(ident("args"), ident("__i")),
                                )),
                            ),
                            vec![ret(bool_lit(false))],
                        ),
                        assign(ident("__i"), op(BinOp::Add, ident("__i"), i(1))),
                    ],
                ),
                ret(op(BinOp::Eq, read_attr(ident("record"), "kwargs"), ident("kwargs"))),
            ],
        ),
        function(
            "__py_mock_call_matches",
            vec![param("left", None), param("right", None)],
            vec![
                if_stmt(
                    unary_not(call_global("hasattr", vec![ident("right"), str_lit("args")])),
                    vec![ret(bool_lit(false))],
                ),
                ret(call_global(
                    "__py_mock_args_match",
                    vec![ident("left"), read_attr(ident("right"), "args"), read_attr(ident("right"), "kwargs")],
                )),
            ],
        ),
    ]
}
