//! C program completion is explicit control flow, not module initialization.

use vybe_ast::{BindingPattern, ExprKind, Statement, StmtKind};
use vybe_platform_libc::emitter::build::*;

pub const ENTRY: &str = "__c_start";

pub fn normalize(body: &mut Vec<Statement>, has_sdl: bool, windowed: bool) {
    if has_sdl {
        vybe_platform_libc::emitter::sdl_app::initialize_input(body);
    }
    let has_main = body.iter().any(|statement| {
        matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "main")
    });

    // An explicit exit request belongs to the C process, not to the lifetime
    // of its startup module. In particular, yielding startup to a UI callback
    // must not invoke the legacy module-end __c_exit_status convention.
    for statement in body.iter_mut() {
        if let StmtKind::VarDecl { declarations, .. } = &mut statement.kind {
            for declaration in declarations {
                if matches!(&declaration.pattern, BindingPattern::Ident(name) if name == "__c_exit_status")
                {
                    declaration.pattern = BindingPattern::Ident("__c_process".into());
                    declaration.type_hint = None;
                    declaration.init = Some(expr(ExprKind::Object(vec![])));
                }
            }
        }
        statement.walk_exprs_mut(&mut |expression| {
            if matches!(&expression.kind, ExprKind::Ident(name) if name == "__c_exit_status") {
                *expression = member(ident("__c_process"), "exit_status");
            }
        });
    }
    if !has_main {
        return;
    }

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
            var_decl_stmt("__c_main_result", call_expr(ident("main"), vec![])),
            stmt(StmtKind::Exit {
                status: Some(status),
            }),
        ],
    ));
    if windowed {
        vybe_platform_libc::emitter::sdl_app::normalize_entry(body, ENTRY);
    }
}
