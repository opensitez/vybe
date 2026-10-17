//! Source-unit atomic facts, using the existing module/function visitor.
use std::collections::HashSet;
use vybe_ast::Statement;
use super::{AtomicModuleScan, collect_atomic_place_idents};

/// Owned semantic facts only. Source-address receipts remain with their storage
/// owner and are never retained by the staged accumulator.
pub(super) struct ModuleAtomicPreparation {
    names: HashSet<String>,
}

impl ModuleAtomicPreparation {
    pub(super) fn new(names: HashSet<String>) -> Self { Self { names } }

    /// Consume complete top-level statements. Only a caller retaining the
    /// original AST storage through emission may request source-body receipts;
    /// an independently released typed-unit producer must pass `None`.
    pub(super) fn consume(&mut self, statements: &[Statement],
        source_bodies: Option<&mut HashSet<(usize, usize)>>,
    ) {
        if let Some(bodies) = source_bodies {
            collect_atomic_place_idents(statements, &mut AtomicModuleScan {
                class_depth: 0, names: &mut self.names, bodies,
            });
        } else {
            collect_atomic_place_idents(statements, &mut self.names);
        }
    }

    pub(super) fn finish(self) -> HashSet<String> { self.names }
}

#[cfg(test)]
#[path = "atomic_preparation_tests.rs"]
mod tests;
