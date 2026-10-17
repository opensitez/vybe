//! C-owned shutdown runs before terminal Exit, not in an unwound finalizer.

use vybe_ast::{BinOp, BitLane, ExprKind, Expression, IntOp, Statement, StmtKind};
use vybe_platform_libc::emitter::build::*;

pub const FINISH: &str = "__c_finish_shutdown";
pub const TERMINATE: &str = "__c_terminate";

fn binary_expr(op: BinOp, left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn effect(value: Expression) -> Statement {
    stmt(StmtKind::Expr(value))
}

fn eq(name: &str, value: i64) -> Expression {
    binary_expr(
        BinOp::Integer(IntOp::Eq, BitLane::W32),
        ident(name),
        int_lit(value),
    )
}

pub fn registration(name: &str, mut args: Vec<Expression>) -> Expression {
    let handler = if args.is_empty() {
        null_lit()
    } else {
        args.remove(0)
    };
    let argument = if args.is_empty() {
        null_lit()
    } else {
        args.remove(0)
    };
    let registry = if name == "at_quick_exit" {
        "__c_quick_callbacks"
    } else {
        "__c_exit_callbacks"
    };
    expr(ExprKind::Sequence(vec![
        call_member(
            ident(registry),
            "push",
            vec![expr(ExprKind::Array(
                vec![handler, int_lit(i64::from(name == "on_exit")), argument]
                    .into_iter()
                    .map(|value| vybe_ast::ArrayElement {
                        key: None,
                        value,
                        spread: false,
                        by_ref: false,
                    })
                    .collect(),
            ))],
        ),
        int_lit(0),
    ]))
}

fn drain(registry: &str) -> Statement {
    let entry = |slot| index_expr(ident("__c_callback"), int_lit(slot));
    stmt(StmtKind::While {
        cond: binary_expr(
            BinOp::Integer(IntOp::GtS, BitLane::W32),
            member(ident(registry), "length"),
            int_lit(0),
        ),
        body: vec![
            var_decl_stmt("__c_callback", call_member(ident(registry), "pop", vec![])),
            if_stmt(
                binary_expr(
                    BinOp::Integer(IntOp::Eq, BitLane::W32),
                    entry(1),
                    int_lit(1),
                ),
                vec![effect(call_expr(
                    entry(0),
                    vec![
                        ident("__c_exit_status"),
                        // This slot stores a pointer VALUE, not the pointee.
                        // Ordinary Index auto-dereferences managed references.
                        call_expr(
                            ident("__c_array_get"),
                            vec![ident("__c_callback"), int_lit(2)],
                        ),
                    ],
                ))],
                Some(vec![effect(call_expr(entry(0), vec![]))]),
            ),
        ],
        else_body: None,
    })
}

pub fn support(body: &mut Vec<Statement>) {
    body.extend([
        var_decl_stmt("__c_exit_callbacks", expr(ExprKind::Array(vec![]))),
        var_decl_stmt("__c_quick_callbacks", expr(ExprKind::Array(vec![]))),
        typed_var_decl_stmt("__c_shutdown_started", "int", int_lit(0)),
        typed_var_decl_stmt("__c_quick_shutdown", "int", int_lit(0)),
        function_stmt(
            FINISH,
            vec![],
            vec![
                if_stmt(
                    eq("__c_shutdown_started", 0),
                    vec![
                        // Mark before invoking any user code; callbacks may call exit.
                        effect(assign_expr(ident("__c_shutdown_started"), int_lit(1))),
                        if_stmt(
                            eq("__c_quick_shutdown", 1),
                            vec![drain("__c_quick_callbacks")],
                            Some(vec![if_stmt(
                                eq("__c_skip_atexit", 0),
                                vec![drain("__c_exit_callbacks")],
                                None,
                            )]),
                        ),
                    ],
                    None,
                ),
                // Reentrant exit still drains partial output, without adding a newline.
                if_stmt(
                    eq("__c_skip_flush", 0),
                    vec![
                        if_stmt(
                            binary_expr(
                                BinOp::Integer(IntOp::Ne, BitLane::W32),
                                call_expr(
                                    ident("__c_string_length"),
                                    vec![ident("__c_stdout_buffer")],
                                ),
                                int_lit(0),
                            ),
                            vec![effect(call_expr(
                                ident("__c_write_stdout"),
                                vec![ident("__c_stdout_buffer")],
                            ))],
                            None,
                        ),
                        effect(assign_expr(ident("__c_stdout_buffer"), str_lit(""))),
                    ],
                    None,
                ),
            ],
        ),
        function_stmt(
            TERMINATE,
            vec!["__c_status", "__c_mode"],
            vec![
                // Keep the existing inline-child emulation local to that child.
                if_stmt(
                    eq("__c_child_exec_done", 1),
                    vec![
                        effect(assign_expr(ident("__c_child_exec_done"), int_lit(0))),
                        stmt(StmtKind::Return(Some(int_lit(0)))),
                    ],
                    None,
                ),
                if_stmt(
                    eq("__c_in_forked_child", 1),
                    vec![
                        effect(assign_expr(ident("__c_child_status"), ident("__c_status"))),
                        effect(assign_expr(ident("__c_in_forked_child"), int_lit(0))),
                        stmt(StmtKind::Return(Some(int_lit(0)))),
                    ],
                    None,
                ),
                effect(assign_expr(ident("__c_exit_status"), ident("__c_status"))),
                if_stmt(
                    eq("__c_mode", 1),
                    vec![
                        effect(assign_expr(ident("__c_skip_atexit"), int_lit(1))),
                        effect(assign_expr(ident("__c_skip_flush"), int_lit(1))),
                    ],
                    None,
                ),
                if_stmt(
                    eq("__c_mode", 2),
                    vec![
                        effect(assign_expr(ident("__c_quick_shutdown"), int_lit(1))),
                        effect(assign_expr(ident("__c_skip_flush"), int_lit(1))),
                    ],
                    None,
                ),
                effect(call_expr(ident(FINISH), vec![])),
                stmt(StmtKind::Exit {
                    status: Some(ident("__c_exit_status")),
                }),
            ],
        ),
    ]);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shutdown_comparisons_are_explicit_w32_integer_operations() {
        let mut body = Vec::new();
        support(&mut body);
        let mut comparisons = 0;
        for statement in &body {
            statement.walk_exprs(&mut |expression| {
                if let ExprKind::Binary { op, .. } = &expression.kind {
                    assert!(matches!(
                        op,
                        BinOp::Integer(IntOp::Eq | IntOp::Ne | IntOp::GtS, BitLane::W32)
                    ));
                    comparisons += 1;
                }
            });
        }
        assert!(comparisons >= 8);
    }

    #[test]
    fn on_exit_argument_is_loaded_as_a_raw_pointer_value() {
        let mut body = Vec::new();
        support(&mut body);
        let mut on_exit_calls = 0;
        for statement in &body {
            statement.walk_exprs(&mut |expression| {
                let ExprKind::Call { callee, args, .. } = &expression.kind else {
                    return;
                };
                if !matches!(&callee.kind, ExprKind::Index { object, .. }
                    if matches!(&object.kind, ExprKind::Ident(name) if name == "__c_callback"))
                    || args.len() != 2
                {
                    return;
                }
                assert!(
                    matches!(&args[1].value.kind, ExprKind::Call { callee, args, .. }
                    if matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_array_get")
                        && args.len() == 2
                        && matches!(&args[1].value.kind, ExprKind::Lit(vybe_ast::Literal::Int(2))))
                );
                on_exit_calls += 1;
            });
        }
        assert_eq!(on_exit_calls, 2);
    }
}
