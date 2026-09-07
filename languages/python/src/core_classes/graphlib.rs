//! `graphlib.TopologicalSorter` — Kahn's algorithm over two adjacency dicts.
//!
//! ⛔ Constraints that apply to every declared class, learned the hard way:
//! membership goes through `__py_contains__` (`BinOp::In` is a walker rewrite),
//! a set is materialised with `list()` before iterating, and every temporary is
//! a PARAMETER — an `assign` to a fresh name lands in a global, so a recursive
//! or nested call would otherwise clobber the caller's.

use super::builders::*;
use vybe_ast::{BinOp, Param, Statement};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn empty_dict() -> Expr {
    Expression::with_span(
        vybe_ast::ExprKind::Map(Vec::new()),
        vybe_ast::Span::default(),
    )
}

use vybe_ast::Expression;

fn preds() -> Expr {
    this_field("_preds")
}

fn succs() -> Expr {
    this_field("_succs")
}

fn contains(hay: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![hay, needle])
}

fn len_of(e: Expr) -> Expr {
    call_global("len", vec![e])
}

fn push(target: Expr, value: Expr) -> Statement {
    expr_stmt(call(member(target, "append"), vec![value]))
}

/// `self._preds[n]` / `self._succs[n]`, creating both when the node is new.
fn ensure_node(node: Expr) -> Statement {
    if_stmt(
        unary_not(contains(preds(), node.clone())),
        vec![
            assign(index(preds(), node.clone()), list_of(vec![])),
            assign(index(succs(), node), list_of(vec![])),
        ],
    )
}

pub(super) const EXCEPTIONS: &[(&str, &str)] = &[("CycleError", "ValueError")];

pub(super) fn exception(name: &str, parent: &str) -> Statement {
    class_extending(name, &[parent], vec![])
}

pub(super) fn topological_sorter() -> Statement {
    class(
        "TopologicalSorter",
        vec![
            init(
                vec![param("graph", Some(null()))],
                vec![
                    set_this("_preds", empty_dict()),
                    set_this("_succs", empty_dict()),
                    set_this("_prepared", bool_lit(false)),
                    set_this("_ready", list_of(vec![])),
                    set_this("_out", i(0)),
                    set_this("_left", empty_dict()),
                    if_stmt(
                        is_not_none(ident("graph")),
                        vec![for_in(
                            "__n",
                            ident("graph"),
                            vec![expr_stmt(call(
                                member(ident("self"), "_add_edges"),
                                vec![ident("__n"), index(ident("graph"), ident("__n"))],
                            ))],
                        )],
                    ),
                ],
            ),
            // `graph[node]` is a SET of predecessors; a set does not iterate
            // directly inside a declaration, so it is materialised first.
            method(
                "_add_edges",
                vec![param("node", Some(null())), param("pred_set", Some(null()))],
                vec![
                    ensure_node(ident("node")),
                    if_stmt(
                        is_not_none(ident("pred_set")),
                        vec![for_in(
                            "__p",
                            call_global("list", vec![ident("pred_set")]),
                            vec![
                                ensure_node(ident("__p")),
                                push(index(preds(), ident("node")), ident("__p")),
                                push(index(succs(), ident("__p")), ident("node")),
                            ],
                        )],
                    ),
                ],
            ),
            method(
                "add",
                vec![param("node", Some(null())), rest_param("predecessors")],
                vec![
                    if_stmt(
                        this_field("_prepared"),
                        vec![raise_call(
                            "__py_exc_ValueError",
                            vec![str_lit("Nodes cannot be added after a call to prepare()")],
                        )],
                    ),
                    expr_stmt(call(
                        member(ident("self"), "_add_edges"),
                        vec![ident("node"), ident("predecessors")],
                    )),
                ],
            ),
            // Kahn, in insertion order so the result is deterministic.
            method(
                "_order",
                vec![
                    param("counts", Some(null())),
                    param("queue", Some(null())),
                    param("out", Some(null())),
                ],
                vec![
                    assign(ident("counts"), empty_dict()),
                    assign(ident("queue"), list_of(vec![])),
                    assign(ident("out"), list_of(vec![])),
                    for_in(
                        "__n",
                        preds(),
                        vec![assign(
                            index(ident("counts"), ident("__n")),
                            len_of(index(preds(), ident("__n"))),
                        )],
                    ),
                    for_in(
                        "__n",
                        preds(),
                        vec![if_stmt(
                            op(BinOp::Eq, index(ident("counts"), ident("__n")), i(0)),
                            vec![push(ident("queue"), ident("__n"))],
                        )],
                    ),
                    while_stmt(
                        op(BinOp::Gt, len_of(ident("queue")), i(0)),
                        vec![
                            assign(ident("__cur"), index(ident("queue"), i(0))),
                            expr_stmt(call(member(ident("queue"), "pop"), vec![i(0)])),
                            push(ident("out"), ident("__cur")),
                            for_in(
                                "__s",
                                index(succs(), ident("__cur")),
                                vec![
                                    assign(
                                        index(ident("counts"), ident("__s")),
                                        op(BinOp::Sub, index(ident("counts"), ident("__s")), i(1)),
                                    ),
                                    if_stmt(
                                        op(BinOp::Eq, index(ident("counts"), ident("__s")), i(0)),
                                        vec![push(ident("queue"), ident("__s"))],
                                    ),
                                ],
                            ),
                        ],
                    ),
                    ret(ident("out")),
                ],
            ),
            // The nodes left over after a full Kahn pass are exactly a cycle.
            method(
                "_cycle_nodes",
                vec![param("order", Some(null())), param("left", Some(null()))],
                vec![
                    assign(
                        ident("order"),
                        call(member(ident("self"), "_order"), vec![]),
                    ),
                    assign(ident("left"), list_of(vec![])),
                    for_in(
                        "__n",
                        preds(),
                        vec![if_stmt(
                            unary_not(contains(ident("order"), ident("__n"))),
                            vec![push(ident("left"), ident("__n"))],
                        )],
                    ),
                    ret(ident("left")),
                ],
            ),
            method(
                "prepare",
                vec![param("cyc", Some(null()))],
                vec![
                    // CPython does NOT raise on a second `prepare()`.
                    assign(
                        ident("cyc"),
                        call(member(ident("self"), "_cycle_nodes"), vec![]),
                    ),
                    if_stmt(
                        op(BinOp::Gt, len_of(ident("cyc")), i(0)),
                        vec![raise_new(
                            "CycleError",
                            vec![str_lit("nodes are in a cycle"), ident("cyc")],
                        )],
                    ),
                    set_this("_prepared", bool_lit(true)),
                    set_this("_left", empty_dict()),
                    for_in(
                        "__n",
                        preds(),
                        vec![assign(
                            index(this_field("_left"), ident("__n")),
                            len_of(index(preds(), ident("__n"))),
                        )],
                    ),
                    set_this("_ready", list_of(vec![])),
                    for_in(
                        "__n",
                        preds(),
                        vec![if_stmt(
                            op(BinOp::Eq, index(this_field("_left"), ident("__n")), i(0)),
                            vec![push(this_field("_ready"), ident("__n"))],
                        )],
                    ),
                ],
            ),
            method(
                "get_ready",
                vec![param("out", Some(null()))],
                vec![
                    assign(
                        ident("out"),
                        call_global("tuple", vec![this_field("_ready")]),
                    ),
                    set_this(
                        "_out",
                        op(BinOp::Add, this_field("_out"), len_of(this_field("_ready"))),
                    ),
                    set_this("_ready", list_of(vec![])),
                    ret(ident("out")),
                ],
            ),
            method(
                "done",
                vec![rest_param("nodes")],
                vec![for_in(
                    "__group",
                    ident("nodes"),
                    vec![for_in(
                        "__n",
                        // ⛔ `done(*nodes)` hands the whole tuple over as a single
                        // argument, so each element may itself be a container.
                        ternary(
                            op(
                                BinOp::GtEq,
                                call_global("__py_container_kind", vec![ident("__group")]),
                                i(2),
                            ),
                            call_global("list", vec![ident("__group")]),
                            list_of(vec![ident("__group")]),
                        ),
                        vec![
                            if_stmt(
                                unary_not(contains(this_field("_left"), ident("__n"))),
                                vec![raise_call(
                                    "__py_exc_ValueError",
                                    vec![str_lit("node was not passed out")],
                                )],
                            ),
                            if_stmt(
                                op(BinOp::NotEq, index(this_field("_left"), ident("__n")), i(0)),
                                vec![raise_call(
                                    "__py_exc_ValueError",
                                    vec![str_lit("node was not passed out")],
                                )],
                            ),
                            set_this("_out", op(BinOp::Sub, this_field("_out"), i(1))),
                            assign(index(this_field("_left"), ident("__n")), i(-1)),
                            for_in(
                                "__s",
                                index(succs(), ident("__n")),
                                vec![
                                    assign(
                                        index(this_field("_left"), ident("__s")),
                                        op(
                                            BinOp::Sub,
                                            index(this_field("_left"), ident("__s")),
                                            i(1),
                                        ),
                                    ),
                                    if_stmt(
                                        op(
                                            BinOp::Eq,
                                            index(this_field("_left"), ident("__s")),
                                            i(0),
                                        ),
                                        vec![push(this_field("_ready"), ident("__s"))],
                                    ),
                                ],
                            ),
                        ],
                    )],
                )],
            ),
            method(
                "is_active",
                vec![],
                vec![ret(binary(
                    BinOp::Or,
                    op(BinOp::Gt, len_of(this_field("_ready")), i(0)),
                    op(BinOp::Gt, this_field("_out"), i(0)),
                ))],
            ),
            method(
                "static_order",
                vec![param("cyc", Some(null())), param("order", Some(null()))],
                vec![
                    assign(
                        ident("cyc"),
                        call(member(ident("self"), "_cycle_nodes"), vec![]),
                    ),
                    if_stmt(
                        op(BinOp::Gt, len_of(ident("cyc")), i(0)),
                        vec![raise_new(
                            "CycleError",
                            vec![str_lit("nodes are in a cycle"), ident("cyc")],
                        )],
                    ),
                    set_this("_prepared", bool_lit(true)),
                    assign(
                        ident("order"),
                        call(member(ident("self"), "_order"), vec![]),
                    ),
                    ret(ident("order")),
                ],
            ),
        ],
    )
}

fn _unused(_p: Param) {}
