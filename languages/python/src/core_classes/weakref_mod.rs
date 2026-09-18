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

fn pair_key() -> Expr {
    index(ident("p"), i(0))
}

fn pair_value() -> Expr {
    index(ident("p"), i(1))
}

fn key_matches(identity: bool) -> Expr {
    if identity {
        op(BinOp::Is, pair_key(), ident("key"))
    } else {
        op(BinOp::Eq, pair_key(), ident("key"))
    }
}

fn append(target: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(target, "append"), vec![value]))
}

fn mapping_members(identity_keys: bool) -> Vec<vybe_ast::ClassMember> {
    vec![
        init(vec![], vec![set_this("_pairs", list_of(vec![]))]),
        method(
            "__setitem__",
            vec![param("key", None), param("value", None)],
            vec![
                for_in(
                    "p",
                    this_field("_pairs"),
                    vec![if_stmt(
                        key_matches(identity_keys),
                        vec![assign(pair_value(), ident("value")), ret(null())],
                    )],
                ),
                append(
                    this_field("_pairs"),
                    list_of(vec![ident("key"), ident("value")]),
                ),
                ret(null()),
            ],
        ),
        method(
            "__getitem__",
            vec![param("key", None)],
            vec![
                for_in(
                    "p",
                    this_field("_pairs"),
                    vec![if_stmt(key_matches(identity_keys), vec![ret(pair_value())])],
                ),
                raise_call("KeyError", vec![ident("key")]),
            ],
        ),
        method(
            "__contains__",
            vec![param("key", None)],
            vec![
                for_in(
                    "p",
                    this_field("_pairs"),
                    vec![if_stmt(
                        key_matches(identity_keys),
                        vec![ret(bool_lit(true))],
                    )],
                ),
                ret(bool_lit(false)),
            ],
        ),
        method(
            "__delitem__",
            vec![param("key", None)],
            vec![
                assign(ident("next"), list_of(vec![])),
                for_in(
                    "p",
                    this_field("_pairs"),
                    vec![if_else(
                        key_matches(identity_keys),
                        vec![],
                        vec![append(ident("next"), ident("p"))],
                    )],
                ),
                assign(this_slot("_pairs"), ident("next")),
                ret(null()),
            ],
        ),
        method("__len__", vec![], vec![ret(len_of(this_field("_pairs")))]),
        method(
            "__iter__",
            vec![],
            vec![ret(call(member(ident("self"), "keys"), vec![]))],
        ),
        method(
            "keys",
            vec![],
            vec![
                assign(ident("out"), list_of(vec![])),
                for_in(
                    "p",
                    this_field("_pairs"),
                    vec![append(ident("out"), pair_key())],
                ),
                ret(ident("out")),
            ],
        ),
        method(
            "values",
            vec![],
            vec![
                assign(ident("out"), list_of(vec![])),
                for_in(
                    "p",
                    this_field("_pairs"),
                    vec![append(ident("out"), pair_value())],
                ),
                ret(ident("out")),
            ],
        ),
        method(
            "items",
            vec![],
            vec![
                assign(ident("out"), list_of(vec![])),
                for_in(
                    "p",
                    this_field("_pairs"),
                    vec![append(
                        ident("out"),
                        tuple_of(vec![pair_key(), pair_value()]),
                    )],
                ),
                ret(ident("out")),
            ],
        ),
        method("copy", vec![], vec![ret(ident("self"))]),
    ]
}

pub(super) fn weak_key_dictionary() -> Statement {
    class("WeakKeyDictionary", mapping_members(true))
}

pub(super) fn weak_value_dictionary() -> Statement {
    class("WeakValueDictionary", mapping_members(false))
}

pub(super) fn weak_set() -> Statement {
    class(
        "WeakSet",
        vec![
            init(vec![], vec![set_this("_items", list_of(vec![]))]),
            method(
                "add",
                vec![param("item", None)],
                vec![
                    for_in(
                        "v",
                        this_field("_items"),
                        vec![if_stmt(
                            op(BinOp::Is, ident("v"), ident("item")),
                            vec![ret(null())],
                        )],
                    ),
                    append(this_field("_items"), ident("item")),
                    ret(null()),
                ],
            ),
            method(
                "discard",
                vec![param("item", None)],
                vec![
                    assign(ident("next"), list_of(vec![])),
                    for_in(
                        "v",
                        this_field("_items"),
                        vec![if_else(
                            op(BinOp::Is, ident("v"), ident("item")),
                            vec![],
                            vec![append(ident("next"), ident("v"))],
                        )],
                    ),
                    assign(this_slot("_items"), ident("next")),
                    ret(null()),
                ],
            ),
            method(
                "__contains__",
                vec![param("item", None)],
                vec![
                    for_in(
                        "v",
                        this_field("_items"),
                        vec![if_stmt(
                            op(BinOp::Is, ident("v"), ident("item")),
                            vec![ret(bool_lit(true))],
                        )],
                    ),
                    ret(bool_lit(false)),
                ],
            ),
            method("__len__", vec![], vec![ret(len_of(this_field("_items")))]),
            method("__iter__", vec![], vec![ret(this_field("_items"))]),
        ],
    )
}
