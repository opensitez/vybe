use crate::goto_lowering::lower_function_gotos;
use std::collections::HashSet;
use vybe_ast::{BindingPattern, Statement, StmtKind};

// Test-only ordered-owner proof for this matrix: all source locals are plain
// integer scalars with distinct already-lowered names, no address escape/VLA.
// This is deliberately not a production inference for arbitrary C bodies.
fn certified_matrix_bindings(body: &[Statement], out: &mut HashSet<String>) {
    for statement in body {
        match &statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    if declaration.array_bounds.is_none() && !declaration.with_events
                        && declaration.type_hint.as_deref().is_some_and(|hint|
                            matches!(hint, "int" | "integer" | "int32"))
                        && let BindingPattern::Ident(name) = &declaration.pattern
                    { out.insert(name.clone()); }
                }
            }
            StmtKind::Block(body) | StmtKind::While { body, .. } | StmtKind::DoWhile { body, .. } =>
                certified_matrix_bindings(body, out),
            StmtKind::For { init, body, .. } => {
                if let Some(init) = init { certified_matrix_bindings(std::slice::from_ref(init.as_ref()), out); }
                certified_matrix_bindings(body, out);
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                certified_matrix_bindings(then_body, out);
                for (_, body) in elifs { certified_matrix_bindings(body, out); }
                if let Some(body) = else_body { certified_matrix_bindings(body, out); }
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases { certified_matrix_bindings(&case.body, out); }
                if let Some(body) = default { certified_matrix_bindings(body, out); }
            }
            _ => {}
        }
    }
}

fn parse_prepared(source: &str, parse: vybe_runtime::registry::ParseFn) -> Result<vybe_ast::Module, String> {
    // Supplied by the PRIVATE test-only patch, never by shared production code.
    let mut module = crate::walker::with_unlowered_gotos_for_tests(|| parse(source))?;
    for statement in &mut module.body {
        if let StmtKind::FunctionDecl { body, params, .. } = &mut statement.kind {
            let mut certified = HashSet::new();
            certified_matrix_bindings(body, &mut certified);
            let reserved = params.iter().map(|parameter| parameter.name.clone()).collect();
            *body = lower_function_gotos(body, &certified, &reserved)
                .map_err(|error| format!("{} at {:?}", error.reason, error.span))?;
        }
    }
    Ok(module)
}

pub(crate) fn parse_prepared_peg(source: &str) -> Result<vybe_ast::Module, String> {
    parse_prepared(source, crate::parse_with_peg)
}

pub(crate) fn parse_prepared_native(source: &str) -> Result<vybe_ast::Module, String> {
    parse_prepared(source, crate::parse_optimized)
}
