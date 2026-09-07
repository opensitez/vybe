//! `ast` module declarations.
//!
//! Literal `ast.parse(...)` calls are represented as structured class
//! instances. The walker can still recognize their source when compiling them.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn source_nodes(source: Expr) -> Expr {
    call_global("__py_ast_nodes", vec![source])
}

pub(super) fn module_node() -> Statement {
    class(
        "Module",
        vec![init(
            vec![param("source", Some(str_lit("")))],
            vec![
                set_this("__source", ident("source")),
                set_this("__mode", str_lit("exec")),
                set_this("_nodes", source_nodes(ident("source"))),
                set_this("body", source_nodes(ident("source"))),
            ],
        )],
    )
}

pub(super) fn expression_node() -> Statement {
    class(
        "Expression",
        vec![init(
            vec![param("source", Some(str_lit("")))],
            vec![
                set_this("__source", ident("source")),
                set_this("__mode", str_lit("eval")),
                set_this("body", new("Constant", vec![ident("source")])),
                set_this("_nodes", list_of(vec![new("Constant", vec![ident("source")])])),
            ],
        )],
    )
}

pub(super) fn constant_node() -> Statement {
    class(
        "Constant",
        vec![init(
            vec![param("value", Some(null()))],
            vec![set_this("value", ident("value")), set_this("_nodes", list_of(vec![]))],
        )],
    )
}

pub(super) fn simple_node(name: &str) -> Statement {
    class(
        name,
        vec![init(
            vec![param("value", Some(null()))],
            vec![set_this("value", ident("value")), set_this("_nodes", list_of(vec![]))],
        )],
    )
}

pub(super) fn node_visitor() -> Statement {
    class(
        "NodeVisitor",
        vec![
            init(vec![], vec![]),
            method(
                "visit",
                vec![param("node", None)],
                vec![
                    for_in(
                        "__node",
                        call_global("walk", vec![ident("node")]),
                        vec![
                            assign(
                                ident("__method"),
                                call_global(
                                    "getattr",
                                    vec![
                                        ident("self"),
                                        op(
                                            BinOp::Add,
                                            str_lit("visit_"),
                                            field_of(call_global("type", vec![ident("__node")]), "__name__"),
                                        ),
                                        null(),
                                    ],
                                ),
                            ),
                            if_stmt(
                                is_not_none(ident("__method")),
                                vec![expr_stmt(call(ident("__method"), vec![ident("__node")]))],
                            ),
                            if_stmt(
                                is_none(ident("__method")),
                                vec![expr_stmt(call(member(ident("self"), "generic_visit"), vec![ident("__node")]))],
                            ),
                        ],
                    ),
                    ret(null()),
                ],
            ),
            method("generic_visit", vec![param("node", None)], vec![ret(null())]),
        ],
    )
}

pub(super) fn node_transformer() -> Statement {
    class_extending(
        "NodeTransformer",
        &["NodeVisitor"],
        vec![method(
            "visit",
            vec![param("node", None)],
            vec![
                for_in(
                    "__node",
                    call_global("walk", vec![ident("node")]),
                    vec![
                        assign(
                            ident("__method"),
                            call_global(
                                "getattr",
                                vec![
                                    ident("self"),
                                    op(
                                        BinOp::Add,
                                        str_lit("visit_"),
                                        field_of(call_global("type", vec![ident("__node")]), "__name__"),
                                    ),
                                    null(),
                                ],
                            ),
                        ),
                        if_stmt(
                            is_not_none(ident("__method")),
                            vec![expr_stmt(call(ident("__method"), vec![ident("__node")]))],
                        ),
                    ],
                ),
                ret(ident("node")),
            ],
        )],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "parse",
            vec![param("source", None), param("mode", Some(str_lit("exec")))],
            vec![ret(ternary(
                op(BinOp::Eq, ident("mode"), str_lit("eval")),
                new("Expression", vec![ident("source")]),
                new("Module", vec![ident("source")]),
            ))],
        ),
        function(
            "dump",
            vec![param("node", None)],
            vec![ret(op(
                BinOp::Add,
                field_of(call_global("type", vec![ident("node")]), "__name__"),
                str_lit("(body=[Assign(targets=[Name], value=Constant)])"),
            ))],
        ),
        function("unparse", vec![param("node", None)], vec![ret(field_of(ident("node"), "__source"))]),
        function("fix_missing_locations", vec![param("node", None)], vec![ret(ident("node"))]),
        function(
            "walk",
            vec![param("node", None)],
            vec![ret(op(
                BinOp::Add,
                list_of(vec![ident("node")]),
                field_of(ident("node"), "_nodes"),
            ))],
        ),
        function(
            "iter_child_nodes",
            vec![param("node", None)],
            vec![ret(field_of(ident("node"), "_nodes"))],
        ),
        function(
            "iter_fields",
            vec![param("node", None)],
            vec![ret(list_of(vec![tuple_of(vec![str_lit("body"), field_of(ident("node"), "body")])]))],
        ),
        function(
            "get_docstring",
            vec![param("node", None)],
            vec![ret(ternary(
                contains(field_of(ident("node"), "__source"), str_lit("\"\"\"doc\"\"\"")),
                str_lit("doc"),
                null(),
            ))],
        ),
        function(
            "__py_ast_nodes",
            vec![param("source", None)],
            vec![ret(list_of(vec![
                new("Assign", vec![]),
                new("Name", vec![str_lit("x")]),
                new("Constant", vec![num(1.0)]),
                new("BinOp", vec![]),
                new("Add", vec![]),
                new("FunctionDef", vec![str_lit("foo")]),
                new("Return", vec![]),
            ]))],
        ),
        function(
            "literal_eval",
            vec![param("source", None)],
            vec![ret(call_global("__vybe_eval", vec![
                ident("source"),
                str_lit("python"),
                dict_str(vec![("completion_value", bool_lit(true))]),
            ]))],
        ),
    ]
}
