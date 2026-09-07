//! `contextvars` — contextual variables as declared AST classes.
//!
//! This is the Python adapter surface for context-local state. It uses ordinary
//! class declarations and helper functions, so tokens, type names and methods
//! flow through the same class/call/error machinery as source-declared Python.

use super::builders::*;
use vybe_ast::{Expression, Statement, StmtKind};

type Expr = Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn is_expr(l: Expr, r: Expr) -> Expr {
    call_global("__py_is__", vec![l, r])
}

fn is_not_expr(l: Expr, r: Expr) -> Expr {
    call_global("__py_is_not__", vec![l, r])
}

fn token_missing() -> Expr {
    ident("__py_context_missing")
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

pub(super) fn token() -> Statement {
    class(
        "Token",
        vec![
            static_field("MISSING", token_missing()),
            init(
                vec![
                    param("var", Some(null())),
                    param("old_value", Some(token_missing())),
                ],
                vec![
                    set_this("var", ident("var")),
                    set_this("old_value", ident("old_value")),
                    set_this("_used", bool_lit(false)),
                ],
            ),
        ],
    )
}

pub(super) fn context_var() -> Statement {
    class(
        "ContextVar",
        vec![
            init(
                vec![param("name", None), param("default", Some(token_missing()))],
                vec![
                    set_this("name", ident("name")),
                    set_this("_default", ident("default")),
                    set_this("_has_value", bool_lit(false)),
                    set_this("_value", null()),
                    expr_stmt(call(
                        member(ident("__py_context_registry"), "append"),
                        vec![ident("self")],
                    )),
                ],
            ),
            method(
                "get",
                vec![param("default", Some(token_missing()))],
                vec![
                    if_stmt(this_field("_has_value"), vec![ret(this_field("_value"))]),
                    if_stmt(
                        is_not_expr(ident("default"), token_missing()),
                        vec![ret(ident("default"))],
                    ),
                    if_stmt(
                        is_not_expr(this_field("_default"), token_missing()),
                        vec![ret(this_field("_default"))],
                    ),
                    raise_call("LookupError", vec![]),
                ],
            ),
            method(
                "set",
                vec![param("value", None)],
                vec![
                    assign(ident("__old"), token_missing()),
                    if_stmt(
                        this_field("_has_value"),
                        vec![assign(ident("__old"), this_field("_value"))],
                    ),
                    assign(
                        ident("__tok"),
                        new("Token", vec![ident("self"), ident("__old")]),
                    ),
                    set_this("_value", ident("value")),
                    set_this("_has_value", bool_lit(true)),
                    ret(ident("__tok")),
                ],
            ),
            method(
                "reset",
                vec![param("token", None)],
                vec![
                    if_stmt(
                        is_not_expr(field_of(ident("token"), "var"), ident("self")),
                        vec![raise_call(
                            "ValueError",
                            vec![str_lit("token was created by a different ContextVar")],
                        )],
                    ),
                    if_stmt(
                        field_of(ident("token"), "_used"),
                        vec![raise_call(
                            "RuntimeError",
                            vec![str_lit("token has already been used once")],
                        )],
                    ),
                    assign(index(ident("token"), str_lit("_used")), bool_lit(true)),
                    if_else(
                        is_expr(field_of(ident("token"), "old_value"), token_missing()),
                        vec![
                            set_this("_value", null()),
                            set_this("_has_value", bool_lit(false)),
                        ],
                        vec![
                            set_this("_value", field_of(ident("token"), "old_value")),
                            set_this("_has_value", bool_lit(true)),
                        ],
                    ),
                    ret(null()),
                ],
            ),
        ],
    )
}

pub(super) fn context() -> Statement {
    class(
        "Context",
        vec![
            init(vec![], vec![set_this("_items", list_of(vec![]))]),
            method(
                "__len__",
                vec![],
                vec![ret(call_global("len", vec![this_field("_items")]))],
            ),
            method(
                "items",
                vec![],
                vec![ret(call_global("list", vec![this_field("_items")]))],
            ),
            method(
                "keys",
                vec![],
                vec![
                    assign(ident("__out"), list_of(vec![])),
                    for_in(
                        "__item",
                        this_field("_items"),
                        vec![expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![index(ident("__item"), i(0))],
                        ))],
                    ),
                    ret(ident("__out")),
                ],
            ),
            method(
                "values",
                vec![],
                vec![
                    assign(ident("__out"), list_of(vec![])),
                    for_in(
                        "__item",
                        this_field("_items"),
                        vec![expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![index(ident("__item"), i(1))],
                        ))],
                    ),
                    ret(ident("__out")),
                ],
            ),
            method(
                "get",
                vec![param("var", None), param("default", Some(null()))],
                vec![
                    for_in(
                        "__item",
                        this_field("_items"),
                        vec![if_stmt(
                            is_expr(index(ident("__item"), i(0)), ident("var")),
                            vec![ret(index(ident("__item"), i(1)))],
                        )],
                    ),
                    ret(ident("default")),
                ],
            ),
            method(
                "run",
                vec![
                    param("func", None),
                    param("arg0", Some(token_missing())),
                    param("arg1", Some(token_missing())),
                    param("extra", Some(token_missing())),
                ],
                vec![
                    assign(
                        ident("__saved"),
                        call_global("__py_context_snapshot", vec![]),
                    ),
                    expr_stmt(call_global(
                        "__py_context_apply_items",
                        vec![this_field("_items")],
                    )),
                    Statement::with_span(
                        StmtKind::Try {
                            body: vec![
                                if_else(
                                    is_expr(ident("arg0"), token_missing()),
                                    vec![assign(ident("__result"), call(ident("func"), vec![]))],
                                    vec![if_else(
                                        is_expr(ident("arg1"), token_missing()),
                                        vec![assign(
                                            ident("__result"),
                                            call(ident("func"), vec![ident("arg0")]),
                                        )],
                                        vec![if_else(
                                            is_expr(ident("extra"), token_missing()),
                                            vec![assign(
                                                ident("__result"),
                                                call(
                                                    ident("func"),
                                                    vec![ident("arg0"), ident("arg1")],
                                                ),
                                            )],
                                            vec![assign(
                                                ident("__result"),
                                                call(
                                                    ident("func"),
                                                    vec![
                                                        ident("arg0"),
                                                        ident("arg1"),
                                                        ident("extra"),
                                                    ],
                                                ),
                                            )],
                                        )],
                                    )],
                                ),
                                assign(
                                    this_slot("_items"),
                                    call_global("__py_context_current_items", vec![]),
                                ),
                            ],
                            catches: vec![],
                            else_body: None,
                            finally: Some(vec![expr_stmt(call_global(
                                "__py_context_restore",
                                vec![ident("__saved")],
                            ))]),
                        },
                        span(),
                    ),
                    ret(ident("__result")),
                ],
            ),
            method(
                "copy",
                vec![],
                vec![
                    assign(ident("__ctx"), new("Context", vec![])),
                    assign(
                        index(ident("__ctx"), str_lit("_items")),
                        call_global("list", vec![this_field("_items")]),
                    ),
                    ret(ident("__ctx")),
                ],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("__py_context_missing", dict_of(vec![])),
        global_assign("__py_context_registry", list_of(vec![])),
        function(
            "__py_context_current_items",
            vec![],
            vec![
                assign(ident("__items"), list_of(vec![])),
                for_in(
                    "__var",
                    ident("__py_context_registry"),
                    vec![if_stmt(
                        field_of(ident("__var"), "_has_value"),
                        vec![expr_stmt(call(
                            member(ident("__items"), "append"),
                            vec![tuple_of(vec![
                                ident("__var"),
                                field_of(ident("__var"), "_value"),
                            ])],
                        ))],
                    )],
                ),
                ret(ident("__items")),
            ],
        ),
        function(
            "__py_context_snapshot",
            vec![],
            vec![
                assign(ident("__items"), list_of(vec![])),
                for_in(
                    "__var",
                    ident("__py_context_registry"),
                    vec![expr_stmt(call(
                        member(ident("__items"), "append"),
                        vec![tuple_of(vec![
                            ident("__var"),
                            field_of(ident("__var"), "_has_value"),
                            field_of(ident("__var"), "_value"),
                        ])],
                    ))],
                ),
                ret(ident("__items")),
            ],
        ),
        function(
            "__py_context_restore",
            vec![param("snapshot", None)],
            vec![
                for_in(
                    "__item",
                    ident("snapshot"),
                    vec![
                        assign(ident("__var"), index(ident("__item"), i(0))),
                        if_else(
                            index(ident("__item"), i(1)),
                            vec![
                                assign(
                                    index(ident("__var"), str_lit("_value")),
                                    index(ident("__item"), i(2)),
                                ),
                                assign(
                                    index(ident("__var"), str_lit("_has_value")),
                                    bool_lit(true),
                                ),
                            ],
                            vec![
                                assign(index(ident("__var"), str_lit("_value")), null()),
                                assign(
                                    index(ident("__var"), str_lit("_has_value")),
                                    bool_lit(false),
                                ),
                            ],
                        ),
                    ],
                ),
                ret(null()),
            ],
        ),
        function(
            "__py_context_apply_items",
            vec![param("items", None)],
            vec![
                for_in(
                    "__var",
                    ident("__py_context_registry"),
                    vec![
                        assign(index(ident("__var"), str_lit("_value")), null()),
                        assign(
                            index(ident("__var"), str_lit("_has_value")),
                            bool_lit(false),
                        ),
                    ],
                ),
                for_in(
                    "__item",
                    ident("items"),
                    vec![
                        assign(ident("__var"), index(ident("__item"), i(0))),
                        assign(
                            index(ident("__var"), str_lit("_value")),
                            index(ident("__item"), i(1)),
                        ),
                        assign(index(ident("__var"), str_lit("_has_value")), bool_lit(true)),
                    ],
                ),
                ret(null()),
            ],
        ),
        function(
            "copy_context",
            vec![],
            vec![
                assign(ident("__ctx"), new("Context", vec![])),
                assign(
                    index(ident("__ctx"), str_lit("_items")),
                    call_global("__py_context_current_items", vec![]),
                ),
                ret(ident("__ctx")),
            ],
        ),
        function(
            "__py_thread_context_call",
            vec![
                param("target", None),
                param("args", None),
                param("thread", Some(null())),
            ],
            vec![
                assign(
                    ident("__saved"),
                    call_global("__py_context_snapshot", vec![]),
                ),
                Statement::with_span(
                    StmtKind::Try {
                        body: vec![expr_stmt(call_global(
                            "__py_thread_call",
                            vec![ident("target"), ident("args"), ident("thread")],
                        ))],
                        catches: vec![],
                        else_body: None,
                        finally: Some(vec![expr_stmt(call_global(
                            "__py_context_restore",
                            vec![ident("__saved")],
                        ))]),
                    },
                    span(),
                ),
                ret(null()),
            ],
        ),
        function(
            "__py_thread_start",
            vec![param("thread", None)],
            vec![if_stmt(
                unary_not(field_of(ident("thread"), "_started")),
                vec![assign(
                    index(ident("thread"), str_lit("_started")),
                    bool_lit(true),
                )],
            )],
        ),
        function(
            "__py_thread_join",
            vec![param("thread", None), param("timeout", Some(null()))],
            vec![
                expr_stmt(call_global("__py_thread_run", vec![ident("thread")])),
                assign(index(ident("thread"), str_lit("_done")), bool_lit(true)),
                ret(null()),
            ],
        ),
    ]
}
