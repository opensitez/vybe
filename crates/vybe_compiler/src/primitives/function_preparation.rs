//! Ordered function headers sealed independently of source body storage.
use super::CallSignature;
use vybe_ast::{PassBy, Statement, StmtKind};

#[derive(Debug)]
pub(super) struct FunctionDeclaration {
    pub(super) name: String,
    pub(super) param_modes: Vec<PassBy>,
    pub(super) param_types: Vec<Option<String>>,
    pub(super) min_arity: usize,
    pub(super) signature: CallSignature,
    pub(super) return_type: Option<String>,
    pub(super) is_generator: bool,
}

#[derive(Default)]
pub(super) struct FunctionRegistrationPreparation {
    declarations: Vec<FunctionDeclaration>,
}

impl FunctionRegistrationPreparation {
    /// Keep every top-level declaration in source order. Canonicalization and
    /// first-wins registration remain at the existing compiler barrier.
    pub(super) fn consume(&mut self, statements: &[Statement]) {
        for statement in statements {
            let StmtKind::FunctionDecl { name, params, return_type, is_generator, .. }
                = &statement.kind else { continue; };
            self.declarations.push(FunctionDeclaration {
                name: name.clone(),
                param_modes: params.iter().map(|param| param.pass_by).collect(),
                param_types: params.iter().map(|param| param.type_hint.as_deref().map(str::to_owned)).collect(),
                // This table's historical kwargs behavior differs from the
                // named-argument CallSignature; keep both contracts intact.
                min_arity: params.iter().take_while(|param| param.default.is_none() && !param.is_rest).count(),
                signature: CallSignature::from_params(params), return_type: return_type.clone(),
                is_generator: *is_generator,
            });
        }
    }

    pub(super) fn finish(self) -> Vec<FunctionDeclaration> { self.declarations }
}

#[cfg(test)]
#[path = "function_preparation_tests.rs"]
mod tests;
