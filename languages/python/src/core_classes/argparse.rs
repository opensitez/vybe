//! `argparse` declared as core classes.
//!
//! This is intentionally AST, not a source prelude. The parser state is class
//! state; parsing itself uses ordinary maps, arrays, loops and exceptions.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn eq(l: Expr, r: Expr) -> Expr {
    op(BinOp::StrictEq, l, r)
}

fn len_of(e: Expr) -> Expr {
    call_global("len", vec![e])
}

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn missing(e: Expr) -> Expr {
    op(BinOp::Or, is_none(e.clone()), op(BinOp::StrictEq, e, str_lit("undefined")))
}

fn dict_get(dict: Expr, key: &str, default: Expr) -> Expr {
    let key_expr = str_lit(key);
    ternary(
        contains(dict.clone(), key_expr.clone()),
        index(dict, key_expr),
        default,
    )
}

fn spec_field(name: &str) -> Expr {
    index(ident("__spec"), str_lit(name))
}

fn ns_get(key: Expr) -> Expr {
    call_global("getattr", vec![ident("__ns"), key, null()])
}

fn ns_set(key: Expr, value: Expr) -> Statement {
    expr_stmt(call_global("setattr", vec![ident("__ns"), key, value]))
}

fn starts_with(value: Expr, prefix: Expr) -> Expr {
    call(member(value, "startswith"), vec![prefix])
}

fn append(list: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(list, "append"), vec![value]))
}

fn system_exit() -> Statement {
    expr_stmt(call_global("__py_raise_SystemExit", vec![num(2.0)]))
}

fn dest_from_flags(flags: Expr) -> Expr {
    call(
        member(
            ternary(
                starts_with(
                    index(flags.clone(), op(BinOp::Sub, len_of(flags.clone()), num(1.0))),
                    str_lit("--"),
                ),
                slice_from(index(flags.clone(), op(BinOp::Sub, len_of(flags.clone()), num(1.0))), num(2.0)),
                ternary(
                    starts_with(
                        index(flags.clone(), op(BinOp::Sub, len_of(flags.clone()), num(1.0))),
                        str_lit("-"),
                    ),
                    slice_from(index(flags.clone(), op(BinOp::Sub, len_of(flags.clone()), num(1.0))), num(1.0)),
                    index(flags.clone(), op(BinOp::Sub, len_of(flags), num(1.0))),
                ),
            ),
            "replace",
        ),
        vec![str_lit("-"), str_lit("_")],
    )
}

fn make_spec_body() -> Vec<Statement> {
    vec![
        assign(
            ident("__flags"),
            ternary(
                call_global("isinstance", vec![ident("flags"), ident("str")]),
                list_of(vec![ident("flags")]),
                ident("flags"),
            ),
        ),
        assign(ident("__first"), index(ident("__flags"), num(0.0))),
        assign(
            ident("__positional"),
            unary_not(starts_with(ident("__first"), str_lit("-"))),
        ),
        assign(ident("__dest"), dict_get(ident("kwargs"), "dest", null())),
        if_stmt(
            missing(ident("__dest")),
            vec![assign(
                ident("__dest"),
                ternary(
                    ident("__positional"),
                    ident("__first"),
                    dest_from_flags(ident("__flags")),
                ),
            )],
        ),
        ret(dict_str(vec![
            ("flags", ident("__flags")),
            ("dest", ident("__dest")),
            ("positional", ident("__positional")),
            (
                "action",
                dict_get(ident("kwargs"), "action", str_lit("store")),
            ),
            ("default", dict_get(ident("kwargs"), "default", null())),
            (
                "required",
                dict_get(ident("kwargs"), "required", bool_lit(false)),
            ),
            ("const", dict_get(ident("kwargs"), "const", null())),
            ("typ", dict_get(ident("kwargs"), "typ", null())),
            (
                "choices",
                dict_get(ident("kwargs"), "choices", list_of(vec![])),
            ),
            ("nargs", dict_get(ident("kwargs"), "nargs", null())),
        ])),
    ]
}

fn convert_value_body() -> Vec<Statement> {
    let is_int = eq(ident("typ"), str_lit("__argparse_int"));
    let is_float = eq(ident("typ"), str_lit("__argparse_float"));
    let is_str = eq(ident("typ"), str_lit("__argparse_str"));
    let is_custom = op(
        BinOp::And,
        is_not_none(ident("typ")),
        op(
            BinOp::And,
            op(BinOp::NotEq, ident("typ"), str_lit("__argparse_int")),
            op(
                BinOp::And,
                op(BinOp::NotEq, ident("typ"), str_lit("__argparse_float")),
                op(BinOp::NotEq, ident("typ"), str_lit("__argparse_str")),
            ),
        ),
    );
    vec![
        assign(ident("__converted"), ident("value")),
        if_stmt(is_int, vec![assign(ident("__converted"), call_global("int", vec![ident("value")]))]),
        if_stmt(is_float, vec![assign(ident("__converted"), call_global("float", vec![ident("value")]))]),
        if_stmt(is_str, vec![assign(ident("__converted"), call_global("str", vec![ident("value")]))]),
        if_stmt(
            is_custom,
            vec![try_except(
                vec![assign(ident("__converted"), call(ident("typ"), vec![ident("value")]))],
                "ValueError",
                vec![system_exit()],
            )],
        ),
        if_stmt(
            op(BinOp::Gt, len_of(ident("choices")), num(0.0)),
            vec![if_stmt(
                unary_not(contains(ident("choices"), ident("__converted"))),
                vec![system_exit()],
            )],
        ),
        ret(ident("__converted")),
    ]
}

fn add_spec_to(target: Expr, also: Option<Expr>) -> Vec<Statement> {
    let mut body = vec![
        assign(ident("__flags"), list_of(vec![ident("flag")])),
        for_in("__alias", ident("aliases"), vec![append(ident("__flags"), ident("__alias"))]),
        assign(
            ident("__spec"),
            call_global("__py_argparse_make_spec", vec![ident("__flags"), ident("kwargs")]),
        ),
        append(target.clone(), ident("__spec")),
    ];
    if let Some(extra) = also {
        body.push(append(extra, ident("__spec")));
    }
    body.push(if_stmt(
        spec_field("positional"),
        vec![append(this_field("_positionals"), ident("__spec"))],
    ));
    body.push(ret(ident("__spec")));
    body
}

fn add_argument_params() -> Vec<vybe_ast::Param> {
    vec![
        param("flag", None),
        param("aliases", Some(list_of(vec![]))),
        param("kwargs", Some(dict_str(vec![]))),
    ]
}

fn initialize_defaults_body() -> Vec<Statement> {
    vec![
        assign(ident("__ns"), ternary(is_none(ident("namespace")), new("Namespace", vec![]), ident("namespace"))),
        for_in(
            "__key",
            this_field("_defaults"),
            vec![ns_set(
                ident("__key"),
                index(this_field("_defaults"), ident("__key")),
            )],
        ),
        for_in(
            "__spec",
            this_field("_options"),
            vec![
                assign(ident("__dest"), spec_field("dest")),
                assign(ident("__action"), spec_field("action")),
                assign(ident("__default"), spec_field("default")),
                if_stmt(
                    missing(ns_get(ident("__dest"))),
                    vec![ns_set(
                        ident("__dest"),
                        ternary(
                            eq(ident("__action"), str_lit("store_true")),
                            bool_lit(false),
                            ternary(
                                eq(ident("__action"), str_lit("store_false")),
                                bool_lit(true),
                                ternary(
                                    eq(ident("__action"), str_lit("count")),
                                    num(0.0),
                                    ternary(missing(ident("__default")), null(), ident("__default")),
                                ),
                            ),
                        ),
                    )],
                ),
            ],
        ),
        for_in(
            "__spec",
            this_field("_positionals"),
            vec![
                assign(ident("__dest"), spec_field("dest")),
                assign(ident("__default"), spec_field("default")),
                if_stmt(
                    op(
                        BinOp::And,
                        eq(spec_field("nargs"), str_lit("?")),
                        op(
                            BinOp::And,
                            missing(ns_get(ident("__dest"))),
                            unary_not(missing(ident("__default"))),
                        ),
                    ),
                    vec![ns_set(ident("__dest"), ident("__default"))],
                ),
            ],
        ),
    ]
}

fn option_action_body() -> Vec<Statement> {
    vec![
        if_stmt(
            eq(ident("__action"), str_lit("store_true")),
            vec![ns_set(ident("__dest"), bool_lit(true))],
        ),
        if_stmt(
            eq(ident("__action"), str_lit("store_false")),
            vec![ns_set(ident("__dest"), bool_lit(false))],
        ),
        if_stmt(
            eq(ident("__action"), str_lit("store_const")),
            vec![ns_set(ident("__dest"), spec_field("const"))],
        ),
        if_stmt(
            eq(ident("__action"), str_lit("count")),
            vec![ns_set(
                ident("__dest"),
                op(
                    BinOp::Add,
                    ns_get(ident("__dest")),
                    ternary(
                        op(BinOp::Gt, len_of(ident("__arg")), num(2.0)),
                        op(BinOp::Sub, len_of(ident("__arg")), num(1.0)),
                        num(1.0),
                    ),
                ),
            )],
        ),
        if_stmt(
            eq(ident("__action"), str_lit("append")),
            vec![
                assign(ident("__i"), op(BinOp::Add, ident("__i"), num(1.0))),
                if_stmt(op(BinOp::GtEq, ident("__i"), len_of(ident("args"))), vec![system_exit()]),
                if_stmt(missing(ns_get(ident("__dest"))), vec![ns_set(ident("__dest"), list_of(vec![]))]),
                append(
                    ns_get(ident("__dest")),
                    call_global("__py_argparse_convert", vec![
                        index(ident("args"), ident("__i")),
                        spec_field("typ"),
                        spec_field("choices"),
                    ]),
                ),
            ],
        ),
        if_stmt(
            op(
                BinOp::And,
                eq(ident("__action"), str_lit("store")),
                eq(spec_field("nargs"), str_lit("?")),
            ),
            vec![
                if_stmt(
                    op(
                        BinOp::And,
                        op(BinOp::Lt, op(BinOp::Add, ident("__i"), num(1.0)), len_of(ident("args"))),
                        unary_not(starts_with(
                            index(ident("args"), op(BinOp::Add, ident("__i"), num(1.0))),
                            str_lit("-"),
                        )),
                    ),
                    vec![
                        assign(ident("__i"), op(BinOp::Add, ident("__i"), num(1.0))),
                        ns_set(
                            ident("__dest"),
                            call_global("__py_argparse_convert", vec![
                                index(ident("args"), ident("__i")),
                                spec_field("typ"),
                                spec_field("choices"),
                            ]),
                        ),
                    ],
                ),
            ],
        ),
        if_stmt(
                op(BinOp::And, eq(ident("__action"), str_lit("store")), is_none(spec_field("nargs"))),
            vec![
                assign(ident("__i"), op(BinOp::Add, ident("__i"), num(1.0))),
                if_stmt(op(BinOp::GtEq, ident("__i"), len_of(ident("args"))), vec![system_exit()]),
                ns_set(
                    ident("__dest"),
                    call_global("__py_argparse_convert", vec![
                        index(ident("args"), ident("__i")),
                        spec_field("typ"),
                        spec_field("choices"),
                    ]),
                ),
            ],
        ),
        if_stmt(
            op(
                BinOp::And,
                eq(ident("__action"), str_lit("store")),
                op(
                    BinOp::And,
                    is_not_none(spec_field("nargs")),
                    op(BinOp::NotEq, spec_field("nargs"), str_lit("?")),
                ),
            ),
            vec![
                assign(ident("__vals"), list_of(vec![])),
                while_stmt(
                    op(
                        BinOp::And,
                        op(BinOp::Lt, op(BinOp::Add, ident("__i"), num(1.0)), len_of(ident("args"))),
                        unary_not(starts_with(index(ident("args"), op(BinOp::Add, ident("__i"), num(1.0))), str_lit("-"))),
                    ),
                    vec![
                        assign(ident("__i"), op(BinOp::Add, ident("__i"), num(1.0))),
                        append(
                            ident("__vals"),
                            call_global("__py_argparse_convert", vec![
                                index(ident("args"), ident("__i")),
                                spec_field("typ"),
                                spec_field("choices"),
                            ]),
                        ),
                    ],
                ),
                if_stmt(
                    op(BinOp::And, eq(spec_field("nargs"), str_lit("+")), eq(len_of(ident("__vals")), num(0.0))),
                    vec![system_exit()],
                ),
                ns_set(ident("__dest"), ident("__vals")),
            ],
        ),
    ]
}

fn parse_args_body() -> Vec<Statement> {
    let mut body = initialize_defaults_body();
    body.extend(vec![
        if_stmt(is_none(ident("args")), vec![assign(ident("args"), list_of(vec![]))]),
        assign(ident("__i"), num(0.0)),
        assign(ident("__pos"), num(0.0)),
        while_stmt(
            op(BinOp::Lt, ident("__i"), len_of(ident("args"))),
            vec![
                assign(ident("__arg"), index(ident("args"), ident("__i"))),
                assign(ident("__handled"), bool_lit(false)),
                if_stmt(
                    starts_with(ident("__arg"), str_lit("-")),
                    vec![
                        for_in(
                            "__spec",
                            this_field("_options"),
                            vec![if_stmt(
                                op(BinOp::And, unary_not(ident("__handled")), unary_not(spec_field("positional"))),
                                vec![
                                    assign(ident("__action"), spec_field("action")),
                                    assign(ident("__dest"), spec_field("dest")),
                                    for_in(
                                        "__flag",
                                        spec_field("flags"),
                                        vec![if_stmt(
                                            op(
                                                BinOp::Or,
                                                eq(ident("__arg"), ident("__flag")),
                                                op(
                                                    BinOp::And,
                                                    eq(ident("__action"), str_lit("count")),
                                                    starts_with(ident("__arg"), ident("__flag")),
                                                ),
                                            ),
                                            {
                                                let mut stmts = vec![assign(ident("__handled"), bool_lit(true))];
                                                stmts.extend(option_action_body());
                                                stmts
                                            },
                                        )],
                                    ),
                                ],
                            )],
                        ),
                        if_stmt(unary_not(ident("__handled")), vec![system_exit()]),
                    ],
                ),
                if_stmt(
                    unary_not(starts_with(ident("__arg"), str_lit("-"))),
                    vec![
                        if_stmt(
                            op(BinOp::And, op(BinOp::Gt, len_of(this_field("_subparsers")), num(0.0)), contains(this_field("_subparsers"), ident("__arg"))),
                            vec![
                                ns_set(this_field("_subparser_dest"), ident("__arg")),
                                assign(
                                    ident("__child_parser"),
                                    index(this_field("_subparsers"), ident("__arg")),
                                ),
                                assign(ident("__parent_ns"), ident("__ns")),
                                assign(
                                    ident("__child_ns"),
                                    call(
                                        member(ident("__child_parser"), "parse_args"),
                                        vec![slice_from(ident("args"), op(BinOp::Add, ident("__i"), num(1.0)))],
                                    ),
                                ),
                                assign(ident("__ns"), ident("__parent_ns")),
                                for_in(
                                    "__child_spec",
                                    field_of(ident("__child_parser"), "_options"),
                                    vec![
                                        assign(ident("__child_dest"), index(ident("__child_spec"), str_lit("dest"))),
                                        if_stmt(
                                            unary_not(missing(call_global(
                                                "getattr",
                                                vec![ident("__child_ns"), ident("__child_dest"), null()],
                                            ))),
                                            vec![ns_set(
                                                ident("__child_dest"),
                                                call_global(
                                                    "getattr",
                                                    vec![ident("__child_ns"), ident("__child_dest"), null()],
                                                ),
                                            )],
                                        ),
                                    ],
                                ),
                                for_in(
                                    "__child_spec",
                                    field_of(ident("__child_parser"), "_positionals"),
                                    vec![
                                        assign(ident("__child_dest"), index(ident("__child_spec"), str_lit("dest"))),
                                        if_stmt(
                                            unary_not(missing(call_global(
                                                "getattr",
                                                vec![ident("__child_ns"), ident("__child_dest"), null()],
                                            ))),
                                            vec![ns_set(
                                                ident("__child_dest"),
                                                call_global(
                                                    "getattr",
                                                    vec![ident("__child_ns"), ident("__child_dest"), null()],
                                                ),
                                            )],
                                        ),
                                    ],
                                ),
                                assign(ident("__i"), len_of(ident("args"))),
                                assign(ident("__handled"), bool_lit(true)),
                            ],
                        ),
                        if_stmt(
                            unary_not(ident("__handled")),
                            vec![
                                assign(ident("__spec"), index(this_field("_positionals"), ident("__pos"))),
                                if_stmt(missing(ident("__spec")), vec![system_exit()]),
                                assign(ident("__dest"), spec_field("dest")),
                                if_stmt(
                                    eq(spec_field("nargs"), str_lit("?")),
                                    vec![ns_set(
                                        ident("__dest"),
                                        call_global("__py_argparse_convert", vec![
                                            ident("__arg"),
                                            spec_field("typ"),
                                            spec_field("choices"),
                                        ]),
                                    )],
                                ),
                                if_stmt(
                                    op(
                                        BinOp::And,
                                        is_not_none(spec_field("nargs")),
                                        op(BinOp::NotEq, spec_field("nargs"), str_lit("?")),
                                    ),
                                    vec![
                                        assign(ident("__vals"), list_of(vec![call_global(
                                            "__py_argparse_convert",
                                            vec![ident("__arg"), spec_field("typ"), spec_field("choices")],
                                        )])),
                                        while_stmt(
                                            op(
                                                BinOp::And,
                                                op(BinOp::Lt, op(BinOp::Add, ident("__i"), num(1.0)), len_of(ident("args"))),
                                                unary_not(starts_with(index(ident("args"), op(BinOp::Add, ident("__i"), num(1.0))), str_lit("-"))),
                                            ),
                                            vec![
                                                assign(ident("__i"), op(BinOp::Add, ident("__i"), num(1.0))),
                                                append(
                                                    ident("__vals"),
                                                    call_global("__py_argparse_convert", vec![
                                                        index(ident("args"), ident("__i")),
                                                        spec_field("typ"),
                                                        spec_field("choices"),
                                                    ]),
                                                ),
                                            ],
                                        ),
                                        ns_set(ident("__dest"), ident("__vals")),
                                    ],
                                ),
                                if_stmt(
                                    op(
                                        BinOp::Or,
                                        is_none(spec_field("nargs")),
                                        eq(spec_field("nargs"), str_lit("?")),
                                    ),
                                    vec![ns_set(
                                        ident("__dest"),
                                        call_global("__py_argparse_convert", vec![
                                            ident("__arg"),
                                            spec_field("typ"),
                                            spec_field("choices"),
                                        ]),
                                    )],
                                ),
                                assign(ident("__pos"), op(BinOp::Add, ident("__pos"), num(1.0))),
                            ],
                        ),
                    ],
                ),
                assign(ident("__i"), op(BinOp::Add, ident("__i"), num(1.0))),
            ],
        ),
        for_in(
            "__spec",
            this_field("_options"),
            vec![if_stmt(
                op(BinOp::And, spec_field("required"), missing(ns_get(spec_field("dest")))),
                vec![system_exit()],
            )],
        ),
        for_in(
            "__group",
            this_field("_mutex_groups"),
            vec![
                assign(ident("__seen"), num(0.0)),
                for_in(
                    "__spec",
                    field_of(ident("__group"), "_members"),
                    vec![if_stmt(
                        ns_get(spec_field("dest")),
                        vec![assign(ident("__seen"), op(BinOp::Add, ident("__seen"), num(1.0)))],
                    )],
                ),
                if_stmt(op(BinOp::Gt, ident("__seen"), num(1.0)), vec![system_exit()]),
                if_stmt(op(BinOp::And, field_of(ident("__group"), "required"), eq(ident("__seen"), num(0.0))), vec![system_exit()]),
            ],
        ),
        ret(ident("__ns")),
    ]);
    body
}

pub(super) fn namespace() -> Statement {
    class("Namespace", vec![init(any_args(), vec![])])
}

pub(super) fn argument_parser() -> Statement {
    class(
        "ArgumentParser",
        vec![
            init(
                any_args(),
                vec![
                    set_this("_options", list_of(vec![])),
                    set_this("_positionals", list_of(vec![])),
                    set_this("_defaults", dict_of(vec![])),
                    set_this("_mutex_groups", list_of(vec![])),
                    set_this("_subparsers", dict_of(vec![])),
                    set_this("_subparser_dest", str_lit("subcommand")),
                ],
            ),
            method(
                "add_argument",
                add_argument_params(),
                add_spec_to(this_field("_options"), None),
            ),
            method(
                "add_argument_group",
                any_args(),
                vec![ret(new("__ArgparseGroup", vec![ident("self"), bool_lit(false)]))],
            ),
            method(
                "add_mutually_exclusive_group",
                vec![kwargs_param("k")],
                vec![
                    assign(ident("__g"), new("__ArgparseGroup", vec![ident("self"), dict_get(ident("k"), "required", bool_lit(false))])),
                    append(this_field("_mutex_groups"), ident("__g")),
                    ret(ident("__g")),
                ],
            ),
            method(
                "add_subparsers",
                vec![kwargs_param("k")],
                vec![
                    assign(ident("__dest"), dict_get(ident("k"), "dest", str_lit("subcommand"))),
                    assign(this_slot("_subparser_dest"), ident("__dest")),
                    ret(new("__ArgparseSubparsers", vec![ident("self"), ident("__dest")])),
                ],
            ),
            method(
                "set_defaults",
                vec![kwargs_param("k")],
                vec![
                    for_in("__key", ident("k"), vec![assign(index(this_field("_defaults"), ident("__key")), index(ident("k"), ident("__key")))]),
                    ret(null()),
                ],
            ),
            method(
                "parse_args",
                vec![param("args", Some(null())), param("namespace", Some(null()))],
                parse_args_body(),
            ),
        ],
    )
}

pub(super) fn argparse_group() -> Statement {
    class(
        "__ArgparseGroup",
        vec![
            init(
                vec![param("parser", None), param("required", Some(bool_lit(false)))],
                vec![
                    set_this("parser", ident("parser")),
                    set_this("required", ident("required")),
                    set_this("_members", list_of(vec![])),
                ],
            ),
            method(
                "add_argument",
                add_argument_params(),
                add_spec_to(field_of(this_field("parser"), "_options"), Some(this_field("_members"))),
            ),
        ],
    )
}

pub(super) fn argparse_subparsers() -> Statement {
    class(
        "__ArgparseSubparsers",
        vec![
            init(
                vec![param("parser", None), param("dest", Some(str_lit("subcommand")))],
                vec![set_this("parser", ident("parser")), set_this("dest", ident("dest"))],
            ),
            method(
                "add_parser",
                vec![param("name", None)],
                vec![
                    assign(ident("__argparse_child_parser"), new("ArgumentParser", vec![])),
                    assign(index(field_of(this_field("parser"), "_subparsers"), ident("name")), ident("__argparse_child_parser")),
                    ret(ident("__argparse_child_parser")),
                ],
            ),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "__py_argparse_make_spec",
            vec![param("flags", None), param("kwargs", None)],
            make_spec_body(),
        ),
        function(
            "__py_argparse_convert",
            vec![
                param("value", None),
                param("typ", Some(null())),
                param("choices", Some(list_of(vec![]))),
            ],
            convert_value_body(),
        ),
    ]
}
