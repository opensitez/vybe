//! The `collections` surface, declared rather than parsed.
//!
//! ⛔ 351 lines of prelude SOURCE behind `contains("collections")`. The helpers
//! are what the walker rewrites `Counter(...)`, `deque(...)`, `defaultdict(...)`
//! and friends into, so they have to exist as globals — but nothing about them
//! needs a second parse of Python text.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn len_of(e: Expr) -> Expr {
    call_global("len", vec![e])
}

/// `d[len(d)] = v` — python's walker lowers `list.append` to an index write,
/// and the prelude used the same shape.
fn push(target: Expr, value: Expr) -> Statement {
    assign(index(target.clone(), len_of(target)), value)
}

fn pop_front(d: Expr) -> Statement {
    expr_stmt(call(member(d, "pop"), vec![i(0)]))
}

fn maxlen_trim(d: &str, drop_front: bool) -> Statement {
    if_stmt(
        is_not_none(ident("maxlen")),
        vec![while_stmt(
            op(BinOp::Gt, len_of(ident(d)), ident("maxlen")),
            vec![if drop_front {
                pop_front(ident(d))
            } else {
                expr_stmt(call(member(ident(d), "pop"), vec![]))
            }],
        )],
    )
}

pub(super) fn deque_functions() -> Vec<Statement> {
    vec![
        function(
            "__py_deque",
            vec![
                param("iterable", Some(null())),
                param("maxlen", Some(null())),
            ],
            vec![
                assign(ident("d"), list_of(vec![])),
                if_stmt(
                    is_not_none(ident("iterable")),
                    vec![assign(
                        ident("d"),
                        call_global("list", vec![ident("iterable")]),
                    )],
                ),
                maxlen_trim("d", true),
                ret(ident("d")),
            ],
        ),
        function(
            "__py_deque_append",
            vec![
                param("d", Some(null())),
                param("value", Some(null())),
                param("maxlen", Some(null())),
            ],
            vec![
                push(ident("d"), ident("value")),
                maxlen_trim("d", true),
                ret(null()),
            ],
        ),
        function(
            "__py_deque_appendleft",
            vec![
                param("d", Some(null())),
                param("value", Some(null())),
                param("maxlen", Some(null())),
            ],
            vec![
                assign(ident("i"), len_of(ident("d"))),
                while_stmt(
                    op(BinOp::Gt, ident("i"), i(0)),
                    vec![
                        assign(
                            index(ident("d"), ident("i")),
                            index(ident("d"), op(BinOp::Sub, ident("i"), i(1))),
                        ),
                        assign(ident("i"), op(BinOp::Sub, ident("i"), i(1))),
                    ],
                ),
                assign(index(ident("d"), i(0)), ident("value")),
                maxlen_trim("d", false),
                ret(null()),
            ],
        ),
        function(
            "__py_deque_extend",
            vec![
                param("d", Some(null())),
                param("values", Some(null())),
                param("maxlen", Some(null())),
            ],
            vec![
                for_in(
                    "v",
                    ident("values"),
                    vec![expr_stmt(call_global(
                        "__py_deque_append",
                        vec![ident("d"), ident("v"), ident("maxlen")],
                    ))],
                ),
                ret(null()),
            ],
        ),
        function(
            "__py_deque_extendleft",
            vec![
                param("d", Some(null())),
                param("values", Some(null())),
                param("maxlen", Some(null())),
            ],
            vec![
                for_in(
                    "v",
                    ident("values"),
                    vec![expr_stmt(call_global(
                        "__py_deque_appendleft",
                        vec![ident("d"), ident("v"), ident("maxlen")],
                    ))],
                ),
                ret(null()),
            ],
        ),
        function(
            "__py_deque_drop_left",
            vec![param("d", Some(null()))],
            vec![
                if_stmt(
                    op(BinOp::Gt, len_of(ident("d")), i(0)),
                    vec![pop_front(ident("d"))],
                ),
                ret(null()),
            ],
        ),
        // Shift every element after the first match down one, then drop the
        // tail — `list.remove` semantics without a second pass.
        function(
            "__py_deque_remove",
            vec![param("d", Some(null())), param("value", Some(null()))],
            vec![
                assign(ident("i"), i(0)),
                assign(ident("found"), bool_lit(false)),
                while_stmt(
                    op(BinOp::Lt, ident("i"), len_of(ident("d"))),
                    vec![
                        if_stmt(
                            binary(
                                BinOp::And,
                                unary_not(ident("found")),
                                op(BinOp::Eq, index(ident("d"), ident("i")), ident("value")),
                            ),
                            vec![assign(ident("found"), bool_lit(true))],
                        ),
                        if_stmt(
                            binary(
                                BinOp::And,
                                ident("found"),
                                op(
                                    BinOp::Lt,
                                    op(BinOp::Add, ident("i"), i(1)),
                                    len_of(ident("d")),
                                ),
                            ),
                            vec![assign(
                                index(ident("d"), ident("i")),
                                index(ident("d"), op(BinOp::Add, ident("i"), i(1))),
                            )],
                        ),
                        assign(ident("i"), op(BinOp::Add, ident("i"), i(1))),
                    ],
                ),
                if_stmt(
                    binary(
                        BinOp::And,
                        ident("found"),
                        op(BinOp::Gt, len_of(ident("d")), i(0)),
                    ),
                    vec![expr_stmt(call(member(ident("d"), "pop"), vec![]))],
                ),
                ret(null()),
            ],
        ),
    ]
}

/// A ChainMap is `{"maps": [...]}` — a plain dict, so every read goes through
/// the ordinary subscript path.
pub(super) fn chainmap_functions() -> Vec<Statement> {
    let maps_of = |e: Expr| index(e, str_lit("maps"));
    vec![
        function(
            "__py_chainmap_new",
            vec![rest_param("maps")],
            vec![ret(Expression_map(call_global(
                "list",
                vec![ident("maps")],
            )))],
        ),
        function(
            "__py_chainmap_get",
            vec![param("cm", Some(null())), param("key", Some(null()))],
            vec![
                for_in(
                    "m",
                    maps_of(ident("cm")),
                    vec![if_stmt(
                        // ⛔ NOT `key in m`. `BinOp::In` is lowered by the
                        // WALKER, and a spliced declaration is never walked, so
                        // the test answered False for every map and
                        // `ChainMap(...)["b"]` came back None.
                        call_global("__py_contains__", vec![ident("m"), ident("key")]),
                        vec![ret(index(ident("m"), ident("key")))],
                    )],
                ),
                ret(null()),
            ],
        ),
        function(
            "__py_chainmap_set",
            vec![
                param("cm", Some(null())),
                param("key", Some(null())),
                param("value", Some(null())),
            ],
            vec![
                assign(
                    index(index(maps_of(ident("cm")), i(0)), ident("key")),
                    ident("value"),
                ),
                ret(null()),
            ],
        ),
        function(
            "__py_chainmap_new_child",
            vec![param("cm", Some(null())), param("child", Some(null()))],
            vec![
                if_stmt(
                    is_none(ident("child")),
                    vec![assign(ident("child"), Expression_dict())],
                ),
                assign(ident("maps"), list_of(vec![ident("child")])),
                for_in(
                    "m",
                    maps_of(ident("cm")),
                    vec![push(ident("maps"), ident("m"))],
                ),
                ret(Expression_map(ident("maps"))),
            ],
        ),
        function(
            "__py_chainmap_parents",
            vec![param("cm", Some(null()))],
            vec![
                assign(ident("maps"), list_of(vec![])),
                assign(ident("i"), i(1)),
                while_stmt(
                    op(BinOp::Lt, ident("i"), len_of(maps_of(ident("cm")))),
                    vec![
                        push(ident("maps"), index(maps_of(ident("cm")), ident("i"))),
                        assign(ident("i"), op(BinOp::Add, ident("i"), i(1))),
                    ],
                ),
                ret(Expression_map(ident("maps"))),
            ],
        ),
        function(
            "__py_chainmap_maps",
            vec![param("cm", Some(null()))],
            vec![ret(maps_of(ident("cm")))],
        ),
    ]
}

/// `{"maps": <expr>}`
#[allow(non_snake_case)]
fn Expression_map(maps: Expr) -> Expr {
    Expression::with_span(
        vybe_ast::ExprKind::Object(vec![vybe_ast::ObjectProperty::KeyValue {
            key: str_lit("maps"),
            value: maps,
        }]),
        vybe_ast::Span::default(),
    )
}

#[allow(non_snake_case)]
fn Expression_dict() -> Expr {
    Expression::with_span(
        vybe_ast::ExprKind::Map(Vec::new()),
        vybe_ast::Span::default(),
    )
}

use vybe_ast::Expression;

pub(super) fn ordereddict_functions() -> Vec<Statement> {
    let rebuild = |first: bool| {
        let mut body = vec![assign(ident("out"), Expression_dict())];
        let copy_others = for_in(
            "k",
            ident("keys"),
            vec![if_stmt(
                op(BinOp::NotEq, ident("k"), ident("key")),
                vec![assign(
                    index(ident("out"), ident("k")),
                    index(ident("d"), ident("k")),
                )],
            )],
        );
        let put_moved = assign(index(ident("out"), ident("key")), ident("moved"));
        if first {
            body.push(put_moved);
            body.push(copy_others);
        } else {
            body.push(copy_others);
            body.push(put_moved);
        }
        body.push(ret(ident("out")));
        body
    };
    vec![function(
        "__py_ordereddict_move_to_end",
        vec![
            param("d", Some(null())),
            param("key", Some(null())),
            param("last", Some(bool_lit(true))),
        ],
        vec![
            if_stmt(
                unary_not(call_global(
                    "__py_contains__",
                    vec![ident("d"), ident("key")],
                )),
                vec![ret(ident("d"))],
            ),
            assign(ident("moved"), index(ident("d"), ident("key"))),
            assign(
                ident("keys"),
                call_global("list", vec![call(member(ident("d"), "keys"), vec![])]),
            ),
            if_stmt(ident("last"), rebuild(false)),
            ret(build_front(rebuild(true))),
        ],
    )]
}

/// The `last=False` arm, as an expression the outer `return` can take.
fn build_front(_stmts: Vec<Statement>) -> Expr {
    call_global("__py_ordereddict_front", vec![ident("d"), ident("key")])
}

pub(super) fn ordereddict_front() -> Statement {
    function(
        "__py_ordereddict_front",
        vec![param("d", Some(null())), param("key", Some(null()))],
        vec![
            assign(ident("out"), Expression_dict()),
            assign(
                index(ident("out"), ident("key")),
                index(ident("d"), ident("key")),
            ),
            for_in(
                "k",
                call_global("list", vec![call(member(ident("d"), "keys"), vec![])]),
                vec![if_stmt(
                    op(BinOp::NotEq, ident("k"), ident("key")),
                    vec![assign(
                        index(ident("out"), ident("k")),
                        index(ident("d"), ident("k")),
                    )],
                )],
            ),
            ret(ident("out")),
        ],
    )
}

pub(super) fn defaultdict_functions() -> Vec<Statement> {
    vec![
        // ⛔ The `f is int` / `f is list` arms the prelude carried are gone: a
        // spliced declaration never has `int` resolved, and calling the factory
        // answers the same thing anyway — `int()` is 0, `list()` is `[]`.
        function(
            "__py_default_factory_value",
            vec![param("f", Some(null()))],
            vec![
                if_stmt(op(BinOp::Eq, ident("f"), str_lit("int")), vec![ret(i(0))]),
                if_stmt(
                    op(BinOp::Eq, ident("f"), str_lit("list")),
                    vec![ret(list_of(vec![]))],
                ),
                if_stmt(
                    op(BinOp::Eq, ident("f"), str_lit("set")),
                    vec![ret(call_global("set", vec![]))],
                ),
                if_stmt(
                    op(BinOp::Eq, ident("f"), str_lit("dict")),
                    vec![ret(Expression_dict())],
                ),
                if_stmt(is_none(ident("f")), vec![ret(null())]),
                ret(call(ident("f"), vec![])),
            ],
        ),
        function(
            "__py_defaultdict",
            vec![
                param("factory", Some(null())),
                param("initial", Some(null())),
            ],
            vec![
                assign(ident("d"), Expression_dict()),
                if_stmt(
                    is_not_none(ident("initial")),
                    vec![for_in(
                        "k",
                        call_global("__py_iter_array__", vec![ident("initial")]),
                        vec![assign(
                            index(ident("d"), ident("k")),
                            index(ident("initial"), ident("k")),
                        )],
                    )],
                ),
                ret(ident("d")),
            ],
        ),
        function(
            "__py_defaultdict_get",
            vec![
                param("d", Some(null())),
                param("factory", Some(null())),
                param("k", Some(null())),
            ],
            vec![
                for_in(
                    "existing",
                    ident("d"),
                    vec![if_stmt(
                        op(BinOp::Eq, ident("existing"), ident("k")),
                        vec![ret(index(ident("d"), ident("existing")))],
                    )],
                ),
                assign(
                    index(ident("d"), ident("k")),
                    call_global("__py_default_factory_value", vec![ident("factory")]),
                ),
                ret(index(ident("d"), ident("k"))),
            ],
        ),
        function(
            "__py_defaultdict_append",
            vec![
                param("d", Some(null())),
                param("factory", Some(null())),
                param("k", Some(null())),
                param("value", Some(null())),
            ],
            vec![
                assign(
                    ident("arr"),
                    call_global(
                        "__py_defaultdict_get",
                        vec![ident("d"), ident("factory"), ident("k")],
                    ),
                ),
                push(ident("arr"), ident("value")),
                ret(null()),
            ],
        ),
        function(
            "__py_defaultdict_add",
            vec![
                param("d", Some(null())),
                param("factory", Some(null())),
                param("k", Some(null())),
                param("value", Some(null())),
            ],
            vec![
                assign(
                    ident("s"),
                    call_global(
                        "__py_defaultdict_get",
                        vec![ident("d"), ident("factory"), ident("k")],
                    ),
                ),
                expr_stmt(call(member(ident("s"), "add"), vec![ident("value")])),
                ret(null()),
            ],
        ),
        function(
            "__py_defaultdict_iadd",
            vec![
                param("d", Some(null())),
                param("factory", Some(null())),
                param("k", Some(null())),
                param("value", Some(null())),
            ],
            vec![
                assign(
                    index(ident("d"), ident("k")),
                    op(
                        BinOp::Add,
                        call_global(
                            "__py_defaultdict_get",
                            vec![ident("d"), ident("factory"), ident("k")],
                        ),
                        ident("value"),
                    ),
                ),
                ret(null()),
            ],
        ),
    ]
}

pub(super) fn user_dict() -> Statement {
    class(
        "UserDict",
        vec![
            init(
                vec![param("initial", Some(null()))],
                vec![
                    set_this("data", Expression_dict()),
                    if_stmt(
                        is_not_none(ident("initial")),
                        vec![for_in(
                            "k",
                            call_global("__py_iter_array__", vec![ident("initial")]),
                            vec![assign(
                                index(this_field("data"), ident("k")),
                                index(ident("initial"), ident("k")),
                            )],
                        )],
                    ),
                ],
            ),
            method(
                "__getitem__",
                vec![param("k", Some(null()))],
                vec![ret(index(this_field("data"), ident("k")))],
            ),
            method(
                "__setitem__",
                vec![param("k", Some(null())), param("v", Some(null()))],
                vec![assign(index(this_field("data"), ident("k")), ident("v"))],
            ),
            method(
                "__repr__",
                vec![],
                vec![ret(call_global("repr", vec![this_field("data")]))],
            ),
        ],
    )
}

pub(super) fn user_list() -> Statement {
    class(
        "UserList",
        vec![
            init(
                vec![param("initial", Some(null()))],
                vec![
                    set_this("data", list_of(vec![])),
                    if_stmt(
                        is_not_none(ident("initial")),
                        vec![set_this(
                            "data",
                            call_global("list", vec![ident("initial")]),
                        )],
                    ),
                ],
            ),
            method(
                "append",
                vec![param("v", Some(null()))],
                vec![push(this_field("data"), ident("v"))],
            ),
            method(
                "extend",
                vec![param("values", Some(null()))],
                vec![for_in(
                    "v",
                    ident("values"),
                    vec![push(this_field("data"), ident("v"))],
                )],
            ),
            method(
                "__repr__",
                vec![],
                vec![ret(call_global("repr", vec![this_field("data")]))],
            ),
        ],
    )
}

pub(super) fn user_string() -> Statement {
    class(
        "UserString",
        vec![
            init(
                vec![param("value", Some(str_lit("")))],
                vec![set_this("data", call_global("str", vec![ident("value")]))],
            ),
            method("__str__", vec![], vec![ret(this_field("data"))]),
            method("__repr__", vec![], vec![ret(this_field("data"))]),
            method(
                "upper",
                vec![],
                vec![ret(call(member(this_field("data"), "upper"), vec![]))],
            ),
        ],
    )
}

pub(super) fn user_functions() -> Vec<Statement> {
    vec![
        function(
            "__py_userdict",
            vec![param("initial", Some(null()))],
            vec![
                assign(ident("d"), Expression_dict()),
                if_stmt(
                    is_not_none(ident("initial")),
                    vec![for_in(
                        "k",
                        call_global("__py_iter_array__", vec![ident("initial")]),
                        vec![assign(
                            index(ident("d"), ident("k")),
                            index(ident("initial"), ident("k")),
                        )],
                    )],
                ),
                ret(ident("d")),
            ],
        ),
        function(
            "__py_userlist",
            vec![param("initial", Some(null()))],
            vec![
                if_stmt(is_none(ident("initial")), vec![ret(list_of(vec![]))]),
                ret(ident("initial")),
            ],
        ),
        function(
            "__py_userstring",
            vec![param("value", Some(str_lit("")))],
            vec![ret(call_global("str", vec![ident("value")]))],
        ),
    ]
}
