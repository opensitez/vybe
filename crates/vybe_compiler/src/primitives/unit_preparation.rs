//! One preparation transaction over finalized common-AST source units.
use super::atomic_preparation::ModuleAtomicPreparation;
use super::function_preparation::{FunctionDeclaration, FunctionRegistrationPreparation};
use super::reference_analysis::{ModuleReferenceBindings, ModuleReferencePreparation};
use super::source_feature_preparation::{SourceFeaturePreparation, SourceFeatures};
use std::borrow::Cow;
use std::collections::HashSet;
use vybe_ast::{CaseAlphabet, Statement};

/// Owns the original vector (or preserves the legacy borrow). Collecting units
/// never moves statements, so source address receipts retain their storage.
pub(super) struct TypedUnitPreparation<'a> {
    body: Cow<'a, [Statement]>,
    next: usize,
    references: ModuleReferencePreparation,
    atomic: ModuleAtomicPreparation,
    source_bodies: HashSet<(usize, usize)>,
    features: SourceFeaturePreparation,
    functions: FunctionRegistrationPreparation,
}

pub(super) struct PreparedTypedUnits<'a> {
    pub(super) body: Cow<'a, [Statement]>,
    pub(super) references: HashSet<String>,
    pub(super) atomic_names: HashSet<String>,
    pub(super) source_bodies: HashSet<(usize, usize)>,
    pub(super) features: SourceFeatures,
    pub(super) functions: Vec<FunctionDeclaration>,
}

impl<'a> TypedUnitPreparation<'a> {
    pub(super) fn new(
        body: Cow<'a, [Statement]>,
        fold: Option<CaseAlphabet>,
        atomic_names: HashSet<String>,
        source_bodies: HashSet<(usize, usize)>,
        features: SourceFeaturePreparation,
    ) -> Self {
        let mut bindings = ModuleReferenceBindings::new(fold);
        bindings.collect(&body);
        Self {
            body,
            next: 0,
            references: bindings.seal(),
            atomic: ModuleAtomicPreparation::new(atomic_names),
            source_bodies,
            features,
            functions: FunctionRegistrationPreparation::default(),
        }
    }

    pub(super) fn collect_next(&mut self) -> Option<usize> {
        let index = self.next;
        let statement = self.body.get(index)?;
        let unit = std::slice::from_ref(statement);
        self.references.consume(unit);
        self.atomic.consume(unit, Some(&mut self.source_bodies));
        self.features.consume(unit);
        self.functions.consume(unit);
        self.next += 1;
        Some(index)
    }

    /// No facts or source storage leave an incomplete transaction. Linking,
    /// declaration registration and emission remain after this source barrier.
    pub(super) fn seal(mut self) -> Result<PreparedTypedUnits<'a>, String> {
        if self.next != self.body.len() {
            return Err(format!(
                "typed-unit preparation incomplete: collected {} of {}",
                self.next,
                self.body.len()
            ));
        }
        let atomic_names = self.atomic.finish();
        if !atomic_names.is_empty() {
            self.source_bodies.clear();
        }
        Ok(PreparedTypedUnits {
            body: self.body,
            references: self.references.finish(),
            atomic_names,
            source_bodies: self.source_bodies,
            features: self.features.finish(),
            functions: self.functions.finish(),
        })
    }
}

#[cfg(test)]
#[path = "unit_preparation_tests.rs"]
mod tests;
