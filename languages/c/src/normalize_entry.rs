//! C program completion is explicit control flow, not module initialization.

use vybe_ast::{BindingPattern, ExprKind, Statement, StmtKind};
use vybe_platform_libc::emitter::build::*;

pub const ENTRY: &str = "__c_start";

fn process_field(name: &str) -> Option<&'static str> {
    match name {
        "__c_exit_status" => Some("exit_status"),
        "__c_exit_callbacks" => Some("exit_callbacks"),
        "__c_quick_callbacks" => Some("quick_callbacks"),
        "__c_shutdown_started" => Some("shutdown_started"),
        "__c_quick_shutdown" => Some("quick_shutdown"),
        "__c_skip_atexit" => Some("skip_atexit"),
        "__c_skip_flush" => Some("skip_flush"),
        "__c_stdout_buffer" => Some("stdout_buffer"),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sdl_parent_has_no_shutdown_and_guest_state_is_process_owned() {
        let mut body = vec![
            typed_var_decl_stmt("__c_exit_status", "int", int_lit(0)),
            typed_var_decl_stmt("__c_skip_atexit", "int", int_lit(0)),
            typed_var_decl_stmt("__c_skip_flush", "int", int_lit(0)),
            var_decl_stmt("__c_stdout_buffer", str_lit("")),
            function_stmt(
                "main",
                vec![],
                vec![stmt(StmtKind::Return(Some(int_lit(7))))],
            ),
        ];
        crate::normalize_shutdown::support(&mut body);
        normalize(&mut body, true, true);
        let mut process_declarations = 0;
        for statement in &body {
            if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
                for declaration in declarations {
                    if let BindingPattern::Ident(name) = &declaration.pattern {
                        assert!(
                            process_field(name).is_none(),
                            "unshared declaration: {name}"
                        );
                        process_declarations += usize::from(name == "__c_process");
                    }
                }
            }
            statement.walk_exprs(&mut |expression| {
                if let ExprKind::Ident(name) = &expression.kind {
                    assert!(
                        process_field(name).is_none(),
                        "unshared state reference: {name}"
                    );
                }
            });
        }
        assert_eq!(process_declarations, 1);
        let parent = body
            .iter()
            .find_map(|statement| match &statement.kind {
                StmtKind::FunctionDecl { name, body, .. } if name == ENTRY => Some(body),
                _ => None,
            })
            .unwrap();
        assert!(
            !parent
                .iter()
                .any(|statement| matches!(statement.kind, StmtKind::Exit { .. }))
        );
        assert!(!parent.iter().any(|statement| matches!(&statement.kind,
            StmtKind::Expr(expression) if matches!(&expression.kind,
                ExprKind::Call { callee, .. } if matches!(&callee.kind,
                    ExprKind::Ident(name) if name == crate::normalize_shutdown::FINISH)))));
    }
}

pub fn normalize(body: &mut Vec<Statement>, has_sdl: bool, windowed: bool) {
    if has_sdl {
        vybe_platform_libc::emitter::sdl_app::initialize_input(body);
    }
    let main_arguments = body.iter_mut().find_map(|statement| {
        let StmtKind::FunctionDecl { name, params, .. } = &mut statement.kind else {
            return None;
        };
        (name == "main").then(|| {
            params
                .iter_mut()
                .map(|param| param.default.take().unwrap_or_else(|| int_lit(0)))
                .collect::<Vec<_>>()
        })
    });

    // An explicit exit request belongs to the C process, not to the lifetime
    // of its startup module. In particular, yielding startup to a UI callback
    // must not invoke the legacy module-end __c_exit_status convention.
    let mut normalized = vec![var_decl_stmt("__c_process", expr(ExprKind::Object(vec![])))];
    for mut statement in body.drain(..) {
        if let StmtKind::VarDecl { declarations, .. } = &mut statement.kind {
            let mut initializers = Vec::new();
            declarations.retain_mut(|declaration| {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    return true;
                };
                let Some(field) = process_field(name) else {
                    return true;
                };
                // No exit request exists until main returns or exit is called.
                if field != "exit_status" {
                    if let Some(value) = declaration.init.take() {
                        initializers.push(stmt(StmtKind::Expr(assign_expr(
                            member(ident("__c_process"), field),
                            value,
                        ))));
                    }
                }
                false
            });
            normalized.extend(initializers);
            if declarations.is_empty() {
                continue;
            }
        }
        statement.walk_exprs_mut(&mut |expression| {
            if let ExprKind::Ident(name) = &expression.kind {
                if let Some(field) = process_field(name) {
                    *expression = member(ident("__c_process"), field);
                }
            }
        });
        normalized.push(statement);
    }
    *body = normalized;
    let Some(main_arguments) = main_arguments else {
        return;
    };

    let status = expr(ExprKind::NullCoalesce {
        left: Box::new(member(ident("__c_process"), "exit_status")),
        right: Box::new(expr(ExprKind::NullCoalesce {
            left: Box::new(ident("__c_main_result")),
            right: Box::new(int_lit(0)),
        })),
    });
    body.push(function_stmt(
        ENTRY,
        vec![],
        vec![
            var_decl_stmt("__c_main_result", call_expr(ident("main"), main_arguments)),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("__c_process"), "exit_status"),
                status.clone(),
            ))),
            stmt(StmtKind::Expr(call_expr(
                ident(crate::normalize_shutdown::FINISH),
                vec![],
            ))),
            stmt(StmtKind::Exit {
                status: Some(member(ident("__c_process"), "exit_status")),
            }),
        ],
    ));
    if windowed {
        vybe_platform_libc::emitter::sdl_app::normalize_entry(body, ENTRY);
    }
}
