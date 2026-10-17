//! Default-off C policy: prefix-preserving graph or proven legacy top targets.
//! No caller may use this adapter on bodies with already-privatized labels.
use std::collections::{HashMap, HashSet};
use vybe_ast::{BindingPattern, ExprKind, LambdaBody, Span, Statement, StmtKind};
use crate::goto_lowering::{contains_unprepared_control, lower_function_gotos, LowerError};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Path { Unchanged, PrefixGraph, ProvenTopTargets }

pub(crate) struct Prepared {
    pub(crate) body: Vec<Statement>,
    pub(crate) path: Path,
}

pub(crate) struct Rejected {
    pub(crate) body: Vec<Statement>,
    pub(crate) error: LowerError,
}

/// Complete-function integration entry: unchanged functions move through with
/// no full AST clone. Rejection returns the original owned body for diagnostics.
pub(crate) fn prepare_owned(body: Vec<Statement>, reserved_names: &HashSet<String>) -> Result<Prepared, Rejected> {
    if !contains_unprepared_control(&body, false) {
        return match crate::goto_lowering::validate_callable_control(&body) {
            Ok(()) => Ok(Prepared { body, path: Path::Unchanged }),
            Err(error) => Err(Rejected { body, error }),
        };
    }
    match prepare(&body, reserved_names) {
        Ok(value) => Ok(value),
        Err(error) => Err(Rejected { body, error }),
    }
}

#[derive(Default)]
struct Flow {
    labels: HashMap<String, Span>,
    top_labels: HashSet<String>,
    targets: HashSet<String>,
    bindings: HashSet<String>,
}

fn fail(span: &Span, reason: impl Into<String>) -> LowerError {
    LowerError { span: span.clone(), reason: reason.into() }
}

impl Flow {
    fn scan(&mut self, body: &[Statement], top: bool) -> Result<(), LowerError> {
        for statement in body {
            match &statement.kind {
                StmtKind::Label(name) => {
                    if self.labels.insert(name.clone(), statement.span.clone()).is_some() {
                        return Err(fail(&statement.span, format!("duplicate C label: {name}")));
                    }
                    if top { self.top_labels.insert(name.clone()); }
                }
                StmtKind::GoTo(name) => { self.targets.insert(name.clone()); }
                StmtKind::Block(body) | StmtKind::DoWhile { body, .. }
                    | StmtKind::While { body, else_body: None, .. } => self.scan(body, false)?,
                StmtKind::For { init, body, .. } => {
                    if let Some(init) = init { self.scan(std::slice::from_ref(init.as_ref()), false)?; }
                    self.scan(body, false)?;
                }
                StmtKind::If { then_body, elifs, else_body, .. } => {
                    self.scan(then_body, false)?;
                    for (_, body) in elifs { self.scan(body, false)?; }
                    if let Some(body) = else_body { self.scan(body, false)?; }
                }
                StmtKind::Switch { cases, default, .. } => {
                    for case in cases { self.scan(&case.body, false)?; }
                    if let Some(body) = default { self.scan(body, false)?; }
                }
                StmtKind::FunctionDecl { .. } => continue,
                StmtKind::VarDecl { declarations, .. } => {
                    for declaration in declarations {
                        let BindingPattern::Ident(name) = &declaration.pattern else {
                            return Err(fail(&statement.span, "unproven binding pattern in legacy C goto fallback"));
                        };
                        self.bindings.insert(name.clone());
                    }
                }
                StmtKind::Expr(_) | StmtKind::Return(_)
                    | StmtKind::Assign { .. } | StmtKind::CompoundAssign { .. }
                    | StmtKind::Break(_) | StmtKind::Continue(_) | StmtKind::Directive { .. }
                    | StmtKind::Empty => {}
                _ => return Err(fail(&statement.span, "unproven control scope in legacy C goto fallback")),
            }
            // Independently callable/statement-expression control cannot be a
            // function-level destination. Do not infer ownership from a name.
            let mut embedded = None;
            statement.walk_exprs(&mut |value| {
                let control = match &value.kind {
                    ExprKind::Lambda { body: LambdaBody::Block(body), .. } => contains_unprepared_control(body, true),
                    ExprKind::FunctionExpr(function) => match &function.kind {
                        StmtKind::FunctionDecl { body, .. } => contains_unprepared_control(body, true), _ => false,
                    },
                    _ => false,
                };
                if control { embedded = Some(value.span.clone()); }
            });
            if let Some(span) = embedded { return Err(fail(&span, "opaque callable control in C goto fallback")); }
        }
        Ok(())
    }
}

/// No binding splitting permissions: every non-prefix declaration declines the
/// graph. Legacy fallback is permitted only after an independent full-function
/// target proof, never because its output happened to erase an unknown goto.
pub(crate) fn prepare(body: &[Statement], reserved_names: &HashSet<String>) -> Result<Prepared, LowerError> {
    let has_goto = contains_unprepared_control(body, false);
    let graph_error = match lower_function_gotos(body, &HashSet::new(), reserved_names) {
        Ok(body) => return Ok(Prepared { body, path: if has_goto { Path::PrefixGraph } else { Path::Unchanged } }),
        Err(error) => error,
    };
    let mut flow = Flow::default();
    flow.scan(body, true)?;
    for target in &flow.targets {
        if !flow.top_labels.contains(target) {
            return Err(fail(&graph_error.span, format!(
                "{}; legacy fallback cannot prove top-level destination {target}", graph_error.reason)));
        }
    }
    if flow.targets.is_empty() {
        return Err(fail(&graph_error.span, "goto preparation failed without a proven legacy target set"));
    }
    // Common lowering uses fixed internal identifiers. Do not admit collisions
    // or an existing dispatcher from an earlier, incomplete lowering barrier.
    let mut used = reserved_names.clone();
    used.extend(flow.labels.keys().cloned());
    used.extend(flow.bindings.iter().cloned());
    for statement in body {
        statement.walk_exprs(&mut |value| {
            if let ExprKind::Ident(name) = &value.kind { used.insert(name.clone()); }
        });
    }
    if used.contains("__c_goto_pc") || used.contains("__c_goto_dispatch") {
        return Err(fail(&graph_error.span, "legacy C goto dispatcher name collides with owner bindings"));
    }
    let lowered = vybe_compiler::primitives::control_flow::lower_gotos(
        body.to_vec(), "__c_goto_pc", "__c_goto_dispatch", false);
    if contains_unprepared_control(&lowered, false) {
        return Err(fail(&graph_error.span, "legacy C goto lowering left unresolved goto statements"));
    }
    Ok(Prepared { body: lowered, path: Path::ProvenTopTargets })
}

#[cfg(test)]
#[path = "goto_prefix_policy_tests.rs"]
mod tests;
