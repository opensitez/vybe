//! Source-derived capability facts sealed before linking and emission.
use vybe_ast::{ExprKind, Literal, Statement};
use super::stmt_uses_proxy;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct SourceFeatures {
    pub(super) bigint_enabled: bool,
    pub(super) uses_proxy: bool,
}

/// Policy admission is supplied by the existing preparation caller. The
/// accumulator retains no AST, source identities, or mutable compiler state.
pub(super) struct SourceFeaturePreparation {
    facts: SourceFeatures,
    inspect_proxy: bool,
}

impl SourceFeaturePreparation {
    pub(super) fn new(bigint_enabled: bool, inspect_proxy: bool, uses_proxy: bool) -> Self {
        Self { facts: SourceFeatures { bigint_enabled, uses_proxy }, inspect_proxy }
    }

    pub(super) fn consume(&mut self, statements: &[Statement]) {
        for statement in statements {
            if self.facts.bigint_enabled && (!self.inspect_proxy || self.facts.uses_proxy) { break; }
            if !self.facts.bigint_enabled {
                statement.walk_exprs(&mut |value| {
                    self.facts.bigint_enabled |= matches!(value.kind, ExprKind::Lit(Literal::BigInt(_)));
                });
            }
            if self.inspect_proxy && !self.facts.uses_proxy {
                self.facts.uses_proxy = stmt_uses_proxy(statement);
            }
        }
    }

    pub(super) fn finish(self) -> SourceFeatures { self.facts }
}

#[cfg(test)]
#[path = "source_feature_preparation_tests.rs"]
mod tests;
