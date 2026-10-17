//! Conservative AST eligibility proof. This does not select physical storage.
use std::collections::{BTreeMap, BTreeSet};
pub use vybe_ast::word_local_transport::WordInterpretation;
use vybe_ast::{
    BinOp, BindingPattern, BitLane, ExprKind, Expression, Literal, Statement, StmtKind,
    StorageAccess, TypeBinding, TypeHint, VarDeclKind,
};
use vybe_ast::{Directives, WordLocalTransport};
mod subset;
pub use subset::{RejectedBinding, Rejection, SubsetProof, prove_subset};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BindingProof {
    pub name: String,
    pub interpretation: WordInterpretation,
}

/// Opt-in eligibility only. The caller resolves declared widths, never names.
/// None means retain the complete region's existing boxed implementation.
pub fn prove(
    directives: &Directives,
    body: &[Statement],
    resolve: impl Fn(&TypeHint) -> Option<WordInterpretation>,
) -> Option<Vec<BindingProof>> {
    if directives.word_local_transport() != WordLocalTransport::ProvenW32 {
        return None;
    }
    let mut bindings = BTreeMap::new();
    for statement in body {
        if let StmtKind::VarDecl { declarations, kind } = &statement.kind {
            if matches!(kind, VarDeclKind::Static | VarDeclKind::FunctionScoped) {
                return None;
            }
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    return None;
                };
                let hint = declaration.type_hint.as_ref()?;
                if hint.binding != TypeBinding::Checked
                    || hint.access != Some(StorageAccess::Ordinary)
                    || declaration.array_bounds.is_some()
                    || declaration.with_events
                    || bindings.insert(name.clone(), resolve(hint)?).is_some()
                {
                    return None;
                }
            }
        }
    }
    if bindings.is_empty() {
        return None;
    }
    let mut initialized = BTreeSet::new();
    let mut visible = BTreeMap::new();
    for statement in body {
        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    return None;
                };
                visible.insert(name.clone(), *bindings.get(name)?);
            }
        }
        statements(
            std::slice::from_ref(statement),
            &visible,
            &mut initialized,
            true,
        )?;
    }
    Some(
        bindings
            .into_iter()
            .map(|(name, interpretation)| BindingProof {
                name,
                interpretation,
            })
            .collect(),
    )
}

fn stored_word(
    value: &Expression,
    interpretation: WordInterpretation,
    bindings: &BTreeMap<String, WordInterpretation>,
    initialized: &BTreeSet<String>,
) -> Option<()> {
    word(value, initialized)?;
    let compatible = match &value.kind {
        ExprKind::Lit(Literal::Int(value)) => match interpretation {
            WordInterpretation::Signed => i32::try_from(*value).is_ok(),
            WordInterpretation::Unsigned => u32::try_from(*value).is_ok(),
        },
        ExprKind::Ident(name) => bindings.get(name) == Some(&interpretation),
        ExprKind::Binary {
            op: BinOp::Integer(op, BitLane::W32),
            ..
        } => !op.is_comparison() && interpretation == WordInterpretation::Signed,
        _ => false,
    };
    compatible.then_some(())
}

fn word(value: &Expression, initialized: &BTreeSet<String>) -> Option<()> {
    match &value.kind {
        ExprKind::Lit(Literal::Int(value))
            if (i64::from(i32::MIN)..=i64::from(u32::MAX)).contains(value) =>
        {
            Some(())
        }
        ExprKind::Ident(name) if initialized.contains(name) => Some(()),
        ExprKind::Binary {
            op: BinOp::Integer(_, BitLane::W32),
            left,
            right,
        } => {
            word(left, initialized)?;
            word(right, initialized)
        }
        _ => None,
    }
}

fn assignment(
    value: &Expression,
    bindings: &BTreeMap<String, WordInterpretation>,
    initialized: &mut BTreeSet<String>,
) -> Option<()> {
    let ExprKind::Assign { target, value } = &value.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &target.kind else {
        return None;
    };
    stored_word(value, *bindings.get(name)?, bindings, initialized)?;
    initialized.insert(name.clone());
    Some(())
}

fn statements(
    body: &[Statement],
    bindings: &BTreeMap<String, WordInterpretation>,
    initialized: &mut BTreeSet<String>,
    declarations_allowed: bool,
) -> Option<()> {
    for statement in body {
        match &statement.kind {
            StmtKind::VarDecl { declarations, .. } if declarations_allowed => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        return None;
                    };
                    if let Some(value) = &declaration.init {
                        stored_word(value, *bindings.get(name)?, bindings, initialized)?;
                        initialized.insert(name.clone());
                    }
                }
            }
            StmtKind::Expr(value) => assignment(value, bindings, initialized)?,
            StmtKind::Return(Some(value)) => word(value, initialized)?,
            StmtKind::Return(None) => {}
            StmtKind::Block(body) => statements(body, bindings, initialized, false)?,
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                word(cond, initialized)?;
                let before = initialized.clone();
                let mut joined = before.clone();
                statements(then_body, bindings, &mut joined, false)?;
                for (cond, body) in elifs {
                    word(cond, &before)?;
                    let mut arm = before.clone();
                    statements(body, bindings, &mut arm, false)?;
                    joined.retain(|name| arm.contains(name));
                }
                let mut other = before;
                if let Some(body) = else_body {
                    statements(body, bindings, &mut other, false)?;
                }
                joined.retain(|name| other.contains(name));
                *initialized = joined;
            }
            StmtKind::While {
                cond,
                body,
                else_body: None,
            }
            | StmtKind::DoWhile {
                body,
                cond,
                until: false,
            } => {
                // Require header inputs before entering, even for do/while.
                word(cond, initialized)?;
                let mut iteration = initialized.clone();
                statements(body, bindings, &mut iteration, false)?;
                // Zero-iteration paths cannot establish a later binding.
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                if let Some(init) = init {
                    statements(
                        std::slice::from_ref(init.as_ref()),
                        bindings,
                        initialized,
                        false,
                    )?;
                }
                if let Some(cond) = cond {
                    word(cond, initialized)?;
                }
                let mut iteration = initialized.clone();
                statements(body, bindings, &mut iteration, false)?;
                if let Some(update) = update {
                    assignment(update, bindings, &mut iteration)?;
                }
            }
            // Calls, closures, addresses, fields, exceptions, jumps, directives
            // and nested declarations need independent storage/flow proofs.
            _ => return None,
        }
    }
    Some(())
}
