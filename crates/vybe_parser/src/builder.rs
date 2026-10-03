//! Transactional semantic construction during matching, without a capture tree
//! or a second normalization walk. A builder owns its AST arena/value stack and
//! journals mutations so speculative alternatives can restore exact state.
use crate::{Span, grammar::RuleId};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BuildError {
    pub span: Span,
    pub message: String,
}

/// Rule hooks run even inside atomic rules, but never in lookahead or implicit
/// trivia skipping. They are grammar bindings, independent of capture modes.
///
/// `checkpoint` is a cheap journal cursor. `rollback(mark)` must restore every
/// mutation since that mark, including popped/replaced semantic values and
/// unfinished frames. Marks may be restored repeatedly. External side effects
/// cannot be rolled back and must not be performed by speculative hooks.
pub trait Builder {
    /// Checked hooks can report semantic construction errors (for example an
    /// out-of-range literal). Errors are fatal for this build and restore its
    /// initial checkpoint; they do not select another syntax alternative.
    /// Defaults preserve existing infallible adapters.
    fn try_begin(&mut self, rule: RuleId, offset: usize) -> Result<(), BuildError> {
        self.begin(rule, offset);
        Ok(())
    }
    fn try_finish(&mut self, rule: RuleId, span: Span, source: &str) -> Result<(), BuildError> {
        self.finish(rule, span, source);
        Ok(())
    }
    fn try_reduce(
        &mut self,
        rule: RuleId,
        fixity: crate::pratt::Fixity,
        span: Span,
    ) -> Result<(), BuildError> {
        self.reduce(rule, fixity, span);
        Ok(())
    }
    /// Opt into a configured island whose atom hooks push one semantic value.
    /// Operator rule hooks are suppressed; `reduce` owns their AST semantics.
    /// Other rule wrappers may disappear on this path. Capture parsing always
    /// retains the ordinary grammar backend.
    fn supports_pratt(&self, _: RuleId) -> bool {
        false
    }
    fn reduce(&mut self, _: RuleId, _: crate::pratt::Fixity, _: Span) {}
    fn checkpoint(&self) -> usize;
    fn rollback(&mut self, mark: usize);
    fn begin(&mut self, rule: RuleId, offset: usize);
    fn finish(&mut self, rule: RuleId, span: Span, source: &str);
}

#[derive(Debug, Default)]
pub(crate) struct NoBuilder;
impl Builder for NoBuilder {
    fn supports_pratt(&self, _: RuleId) -> bool {
        true
    }
    fn checkpoint(&self) -> usize {
        0
    }
    fn rollback(&mut self, _: usize) {}
    fn begin(&mut self, _: RuleId, _: usize) {}
    fn finish(&mut self, _: RuleId, _: Span, _: &str) {}
}

impl<B: Builder> Builder for &mut B {
    fn try_begin(&mut self, rule: RuleId, offset: usize) -> Result<(), BuildError> {
        (**self).try_begin(rule, offset)
    }
    fn try_finish(&mut self, rule: RuleId, span: Span, source: &str) -> Result<(), BuildError> {
        (**self).try_finish(rule, span, source)
    }
    fn try_reduce(
        &mut self,
        rule: RuleId,
        fixity: crate::pratt::Fixity,
        span: Span,
    ) -> Result<(), BuildError> {
        (**self).try_reduce(rule, fixity, span)
    }
    fn supports_pratt(&self, rule: RuleId) -> bool {
        (**self).supports_pratt(rule)
    }
    fn reduce(&mut self, rule: RuleId, fixity: crate::pratt::Fixity, span: Span) {
        (**self).reduce(rule, fixity, span);
    }
    fn checkpoint(&self) -> usize {
        (**self).checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        (**self).rollback(mark);
    }
    fn begin(&mut self, rule: RuleId, offset: usize) {
        (**self).begin(rule, offset);
    }
    fn finish(&mut self, rule: RuleId, span: Span, source: &str) {
        (**self).finish(rule, span, source);
    }
}
