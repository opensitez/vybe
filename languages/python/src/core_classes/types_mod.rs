//! `types` module globals, declared as AST instead of parsed prelude source.

use super::builders::*;
use vybe_ast::{Argument, ExprKind, Expression, LambdaBody, Statement};

fn call_func_obj_spread() -> Expression {
    Expression::with_span(
        ExprKind::Call {
            callee: Box::new(ident("func")),
            args: vec![
                Argument::positional(ident("obj")),
                Argument {
                    value: ident("args"),
                    name: None,
                    by_ref: false,
                    spread: true,
                },
            ],
            optional: false,
        },
        span(),
    )
}

fn method_type_lambda() -> Expression {
    Expression::with_span(
        ExprKind::Lambda {
            params: vec![rest_param("args")],
            body: LambdaBody::Expr(Box::new(call_func_obj_spread())),
            is_async: false,
            captures: vec![],
        },
        span(),
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "SimpleNamespace",
            vec![kwargs_param("kwargs")],
            vec![ret(ident("kwargs"))],
        ),
        function(
            "__py_simple_namespace_repr",
            vec![param("ns", None)],
            vec![
                assign(ident("parts"), str_lit("")),
                assign(ident("sep"), str_lit("")),
                for_in(
                    "k",
                    ident("ns"),
                    vec![
                        assign(
                            ident("parts"),
                            add(
                                add(add(add(ident("parts"), ident("sep")), ident("k")), str_lit("=")),
                                call_global("repr", vec![index(ident("ns"), ident("k"))]),
                            ),
                        ),
                        assign(ident("sep"), str_lit(", ")),
                    ],
                ),
                ret(add(add(str_lit("namespace("), ident("parts")), str_lit(")"))),
            ],
        ),
        function(
            "MappingProxyType",
            vec![param("data", None)],
            vec![ret(ident("data"))],
        ),
        global_assign("FunctionType", str_lit("FunctionType")),
        global_assign("LambdaType", ident("FunctionType")),
        global_assign("GeneratorType", str_lit("GeneratorType")),
        global_assign("CoroutineType", str_lit("CoroutineType")),
        function(
            "MethodType",
            vec![param("func", None), param("obj", None)],
            vec![ret(method_type_lambda())],
        ),
        function(
            "DynamicClassAttribute",
            vec![param("func", None)],
            vec![ret(call_global("property", vec![ident("func")]))],
        ),
        function(
            "resolve_bases",
            vec![param("bases", None)],
            vec![ret(ident("bases"))],
        ),
        function(
            "new_class",
            vec![
                param("name", None),
                param("bases", Some(tuple_of(vec![]))),
                param("kwds", Some(null())),
                param("exec_body", Some(null())),
            ],
            vec![
                assign(ident("ns"), dict_of(vec![])),
                if_stmt(
                    is_not_none(ident("exec_body")),
                    vec![expr_stmt(call(ident("exec_body"), vec![ident("ns")]))],
                ),
                ret(Expression::with_span(
                    ExprKind::Lambda {
                        params: vec![],
                        body: LambdaBody::Expr(Box::new(ident("ns"))),
                        is_async: false,
                        captures: vec![],
                    },
                    span(),
                )),
            ],
        ),
    ]
}
