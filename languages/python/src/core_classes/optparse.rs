//! `optparse` — a compact option parser over Python objects.
//!
//! This implements the behavior exercised by the stdlib surface tests:
//! option registration, grouped options, defaults, value conversion, choices,
//! append/count/const actions, positional leftovers, and `SystemExit` on parse
//! errors.

use super::builders::*;
use vybe_ast::{BinOp, ExprKind, Statement};

type Expr = vybe_ast::Expression;

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn not_contains(haystack: Expr, needle: Expr) -> Expr {
    unary_not(contains(haystack, needle))
}

fn eq(left: Expr, right: Expr) -> Expr {
    binary(BinOp::Eq, left, right)
}

fn and(left: Expr, right: Expr) -> Expr {
    binary(BinOp::And, left, right)
}

fn idx(object: Expr, key: &str) -> Expr {
    index(object, str_lit(key))
}

fn empty_object() -> Expr {
    Expr::new(ExprKind::Object(vec![]))
}

fn type_of(expr: Expr) -> Expr {
    Expr::new(ExprKind::TypeOf(Box::new(expr)))
}

fn is_undefined(expr: Expr) -> Expr {
    binary(BinOp::StrictEq, type_of(expr), str_lit("undefined"))
}

fn is_missing(expr: Expr) -> Expr {
    binary(
        BinOp::Or,
        binary(BinOp::Or, is_none(expr.clone()), is_undefined(expr.clone())),
        eq(expr, str_lit("undefined")),
    )
}

fn opts() -> Expr {
    this_field("_options")
}

fn option_params() -> Vec<vybe_ast::Param> {
    vec![param("flags", None), kwargs_param("kwargs")]
}

fn option_spec_append_body(target: Expr) -> Vec<Statement> {
    vec![
        assign(
            ident("__flags"),
            ternary(
                eq(
                    call_global("__py_type_name", vec![ident("flags")]),
                    str_lit("str"),
                ),
                list_of(vec![ident("flags")]),
                ident("flags"),
            ),
        ),
        assign(
            ident("__action"),
            dict_get(ident("kwargs"), "action", str_lit("store")),
        ),
        assign(ident("__dest"), dict_get(ident("kwargs"), "dest", null())),
        if_stmt(
            is_missing(ident("__dest")),
            vec![assign(
                ident("__dest"),
                call(
                    member(
                        index(
                            ident("__flags"),
                            binary(BinOp::Sub, call_global("len", vec![ident("__flags")]), num(1.0)),
                        ),
                        "replace",
                    ),
                    vec![str_lit("-"), str_lit("")],
                ),
            )],
        ),
        append(
            target,
            dict_str(vec![
                ("flags", ident("__flags")),
                ("action", ident("__action")),
                ("dest", ident("__dest")),
                ("typ", dict_get(ident("kwargs"), "type", null())),
                ("default", dict_get(ident("kwargs"), "default", null())),
                ("const", dict_get(ident("kwargs"), "const", null())),
                (
                    "choices",
                    dict_get(ident("kwargs"), "choices", list_of(vec![])),
                ),
            ]),
        ),
        ret(null()),
    ]
}

fn append(list: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(list, "append"), vec![value]))
}

fn system_exit() -> Statement {
    expr_stmt(call_global("__py_raise_SystemExit", vec![num(2.0)]))
}

fn dict_get(dict: Expr, key: &str, default: Expr) -> Expr {
    let value = index(dict.clone(), str_lit(key));
    ternary(is_missing(value.clone()), default, value)
}

fn convert_value_body() -> Vec<Statement> {
    vec![
        if_stmt(is_missing(ident("typ")), vec![ret(ident("value"))]),
        if_stmt(
            eq(ident("typ"), str_lit("choice")),
            vec![
                if_stmt(
                    not_contains(ident("choices"), ident("value")),
                    vec![system_exit()],
                ),
                ret(ident("value")),
            ],
        ),
        if_stmt(
            eq(ident("typ"), str_lit("int")),
            vec![try_except(
                vec![ret(call_global("int", vec![ident("value")]))],
                "ValueError",
                vec![system_exit()],
            )],
        ),
        if_stmt(
            eq(ident("typ"), str_lit("float")),
            vec![try_except(
                vec![ret(call_global("float", vec![ident("value")]))],
                "ValueError",
                vec![system_exit()],
            )],
        ),
        ret(ident("value")),
    ]
}

pub(super) fn values() -> Statement {
    class("__PyOptValues", vec![])
}

pub(super) fn option_group() -> Statement {
    class(
        "OptionGroup",
        vec![
            init(
                vec![
                    param("parser", None),
                    param("title", Some(str_lit(""))),
                    param("description", Some(null())),
                ],
                vec![
                    set_this("parser", ident("parser")),
                    set_this("title", ident("title")),
                    set_this("description", ident("description")),
                    set_this("_options", list_of(vec![])),
                ],
            ),
            method(
                "add_option",
                option_params(),
                option_spec_append_body(opts()),
            ),
        ],
    )
}

pub(super) fn option_parser() -> Statement {
    class(
        "OptionParser",
        vec![
            init(any_args(), vec![set_this("_options", list_of(vec![]))]),
            method(
                "add_option",
                option_params(),
                option_spec_append_body(opts()),
            ),
            method(
                "add_option_group",
                vec![param("group", None)],
                vec![
                    for_in(
                        "__spec",
                        field_of(ident("group"), "_options"),
                        vec![append(opts(), ident("__spec"))],
                    ),
                    ret(ident("group")),
                ],
            ),
            method(
                "parse_args",
                vec![param("args", Some(null())), param("values", Some(null()))],
                parse_args_body(),
            ),
        ],
    )
}

fn parse_args_body() -> Vec<Statement> {
    vec![
        if_stmt(
            is_none(ident("args")),
            vec![assign(ident("args"), list_of(vec![]))],
        ),
        assign(
            ident("__opts"),
            ternary(is_none(ident("values")), empty_object(), ident("values")),
        ),
        assign(ident("__remaining"), list_of(vec![])),
        for_in(
            "__spec",
            opts(),
            vec![
                assign(ident("__dest"), idx(ident("__spec"), "dest")),
                assign(ident("__action"), idx(ident("__spec"), "action")),
                assign(ident("__default"), idx(ident("__spec"), "default")),
                assign(
                    ident("__initial"),
                    ternary(is_missing(ident("__default")), null(), ident("__default")),
                ),
                if_stmt(
                    and(
                        eq(ident("__action"), str_lit("count")),
                        is_missing(ident("__default")),
                    ),
                    vec![assign(ident("__initial"), num(0.0))],
                ),
                assign(index(ident("__opts"), ident("__dest")), ident("__initial")),
            ],
        ),
        assign(ident("__i"), num(0.0)),
        while_stmt(
            binary(
                BinOp::Lt,
                ident("__i"),
                call_global("len", vec![ident("args")]),
            ),
            vec![
                assign(ident("__arg"), index(ident("args"), ident("__i"))),
                assign(ident("__handled"), bool_lit(false)),
                for_in("__spec", opts(), parse_one_option_body()),
                if_stmt(
                    unary_not(ident("__handled")),
                    vec![
                        if_stmt(
                            call(member(ident("__arg"), "startswith"), vec![str_lit("-")]),
                            vec![system_exit()],
                        ),
                        append(ident("__remaining"), ident("__arg")),
                    ],
                ),
                assign(ident("__i"), binary(BinOp::Add, ident("__i"), num(1.0))),
            ],
        ),
        ret(tuple_of(vec![ident("__opts"), ident("__remaining")])),
    ]
}

fn parse_one_option_body() -> Vec<Statement> {
    vec![if_stmt(
        and(
            unary_not(ident("__handled")),
            contains(idx(ident("__spec"), "flags"), ident("__arg")),
        ),
        vec![
            assign(ident("__handled"), bool_lit(true)),
            assign(ident("__action"), idx(ident("__spec"), "action")),
            assign(ident("__dest"), idx(ident("__spec"), "dest")),
            if_stmt(
                eq(ident("__action"), str_lit("store_true")),
                vec![assign(
                    index(ident("__opts"), ident("__dest")),
                    bool_lit(true),
                )],
            ),
            if_stmt(
                eq(ident("__action"), str_lit("store_false")),
                vec![assign(
                    index(ident("__opts"), ident("__dest")),
                    bool_lit(false),
                )],
            ),
            if_stmt(
                eq(ident("__action"), str_lit("store_const")),
                vec![assign(
                    index(ident("__opts"), ident("__dest")),
                    idx(ident("__spec"), "const"),
                )],
            ),
            if_stmt(
                eq(ident("__action"), str_lit("append_const")),
                vec![
                    if_stmt(
                        is_missing(index(ident("__opts"), ident("__dest"))),
                        vec![assign(
                            index(ident("__opts"), ident("__dest")),
                            list_of(vec![]),
                        )],
                    ),
                    append(
                        index(ident("__opts"), ident("__dest")),
                        idx(ident("__spec"), "const"),
                    ),
                ],
            ),
            if_stmt(
                eq(ident("__action"), str_lit("count")),
                vec![assign(
                    index(ident("__opts"), ident("__dest")),
                    binary(
                        BinOp::Add,
                        index(ident("__opts"), ident("__dest")),
                        num(1.0),
                    ),
                )],
            ),
            if_stmt(
                eq(ident("__action"), str_lit("append")),
                vec![
                    assign(ident("__i"), binary(BinOp::Add, ident("__i"), num(1.0))),
                    if_stmt(
                        binary(
                            BinOp::GtEq,
                            ident("__i"),
                            call_global("len", vec![ident("args")]),
                        ),
                        vec![system_exit()],
                    ),
                    if_stmt(
                        is_missing(index(ident("__opts"), ident("__dest"))),
                        vec![assign(
                            index(ident("__opts"), ident("__dest")),
                            list_of(vec![]),
                        )],
                    ),
                    append(
                        index(ident("__opts"), ident("__dest")),
                        call_global(
                            "__py_optparse_convert",
                            vec![
                                index(ident("args"), ident("__i")),
                                idx(ident("__spec"), "typ"),
                                idx(ident("__spec"), "choices"),
                            ],
                        ),
                    ),
                ],
            ),
            if_stmt(
                eq(ident("__action"), str_lit("store")),
                vec![
                    assign(ident("__i"), binary(BinOp::Add, ident("__i"), num(1.0))),
                    if_stmt(
                        binary(
                            BinOp::GtEq,
                            ident("__i"),
                            call_global("len", vec![ident("args")]),
                        ),
                        vec![system_exit()],
                    ),
                    assign(
                        index(ident("__opts"), ident("__dest")),
                        call_global(
                            "__py_optparse_convert",
                            vec![
                                index(ident("args"), ident("__i")),
                                idx(ident("__spec"), "typ"),
                                idx(ident("__spec"), "choices"),
                            ],
                        ),
                    ),
                ],
            ),
        ],
    )]
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![function(
        "__py_optparse_convert",
        vec![
            param("value", None),
            param("typ", Some(null())),
            param("choices", Some(list_of(vec![]))),
        ],
        convert_value_body(),
    )]
}
