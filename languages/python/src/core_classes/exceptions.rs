//! Exception helpers declared as AST instead of source preludes.

use super::builders::*;
use vybe_ast::{BinOp, Statement, StmtKind};

type Expr = vybe_ast::Expression;

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

pub(super) fn exception_group_declarations() -> Vec<Statement> {
    vec![function(
        "__py_exception_group_split",
        vec![param("eg", None), param("typ", None)],
        vec![
            assign(ident("matched"), list_of(vec![])),
            assign(ident("rest"), list_of(vec![])),
            for_in(
                "e",
                read_attr(ident("eg"), "exceptions"),
                vec![if_else(
                    binary(
                        BinOp::Eq,
                        call_global("__py_type_name", vec![ident("e")]),
                        ident("typ"),
                    ),
                    vec![expr_stmt(call(
                        member(ident("matched"), "append"),
                        vec![ident("e")],
                    ))],
                    vec![expr_stmt(call(
                        member(ident("rest"), "append"),
                        vec![ident("e")],
                    ))],
                )],
            ),
            ret(tuple_of(vec![
                call_global(
                    "ExceptionGroup",
                    vec![read_attr(ident("eg"), "message"), ident("matched")],
                ),
                call_global(
                    "ExceptionGroup",
                    vec![read_attr(ident("eg"), "message"), ident("rest")],
                ),
            ])),
        ],
    )]
}
