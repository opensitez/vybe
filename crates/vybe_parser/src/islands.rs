//! Explicit source-language Pratt bindings over existing grammar rule names.
//! These are optimization/semantic contracts, never inferred from PEG order.
use crate::{
    CompiledGrammar, Diagnostic, Span,
    pratt::Fixity,
    program::{BoundOperator, OwnedPratt, Program},
};

#[derive(Debug, Clone, Copy)]
pub struct Operator<'a> {
    pub rule: &'a str,
    pub precedence: u16,
    pub fixity: Fixity,
}
/// Match the original expression rule's trailing trivia behavior explicitly.
/// A normal `atom ~ operator_tail*` consumes trivia before its optional tail;
/// a bare atom or an atomic expression generally preserves it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrailingTrivia {
    Preserve,
    Consume,
}
impl CompiledGrammar {
    /// Opt into Pratt for recognition and cooperating direct-AST builders.
    /// The grammar author certifies the table describes the target rule's
    /// expression language. Capture mode stays on the original PEG grammar.
    /// Operator rules must be lexical: their semantic hooks are suppressed.
    pub fn bind_pratt(
        &mut self,
        target: &str,
        atom: &str,
        operators: &[Operator<'_>],
        trailing_trivia: TrailingTrivia,
    ) -> Result<(), Vec<Diagnostic>> {
        let span = self
            .rule_id(target)
            .map_or(Span { start: 0, end: 0 }, |id| {
                self.syntax().rules[id].name_span
            });
        let error = |message: String| vec![Diagnostic::new("G014", message, span)];
        let target = self
            .rule_id(target)
            .ok_or_else(|| error(format!("unknown Pratt target '{target}'")))?;
        let atom = self
            .rule_id(atom)
            .ok_or_else(|| error(format!("unknown Pratt atom '{atom}'")))?;
        if target == atom || operators.len() > 128 {
            return Err(error(
                "Pratt atom must differ from target; at most 128 operator bindings are allowed"
                    .into(),
            ));
        }
        let mut bound = Vec::new();
        let mut positions = std::collections::HashSet::new();
        for op in operators {
            let rule = self
                .rule_id(op.rule)
                .ok_or_else(|| error(format!("unknown Pratt operator '{}'", op.rule)))?;
            let prefix = op.fixity == Fixity::Prefix;
            if !positions.insert((rule, prefix)) {
                return Err(error(format!("ambiguous Pratt operator '{}'", op.rule)));
            }
            bound.push(BoundOperator {
                rule,
                precedence: op.precedence,
                fixity: op.fixity,
            });
        }
        for rule in std::iter::once(target)
            .chain(std::iter::once(atom))
            .chain(bound.iter().map(|op| op.rule))
        {
            let facts = self.analysis().rules[rule];
            if facts.may_succeed_empty || facts.effects.is_stack_dependent() {
                return Err(error(format!(
                    "Pratt rule '{}' is nullable/uncertain or stack-dependent; use the ordinary backend",
                    self.rule(rule).name
                )));
            }
        }
        self.pratt[target] = Some(OwnedPratt {
            atom,
            operators: bound,
            trailing_trivia,
        });
        Ok(())
    }
}
