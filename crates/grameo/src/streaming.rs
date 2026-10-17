//! Opt-in typed semantic-unit emission during parsing.
//!
//! A completed rule is not an irreversible PEG commit: an enclosing choice can
//! still backtrack. Both the semantic producer and emission sink must restore
//! all state at checkpoints. A sink may emit VM instructions immediately when
//! it has a real journal, or queue typed units for a later sealed compiler
//! boundary. Neither may execute externally visible code before validation.
use crate::{
    Span,
    builder::{BuildError, Builder},
    grammar::RuleId,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Boundary {
    pub rule: RuleId,
    pub span: Span,
}

/// Existing AST builders opt in at complete statement/declaration boundaries.
/// The borrowed value stays owned by the builder, so speculative transfer does
/// not remove values that its existing rollback journal needs. The value must
/// contain the original scope/class/directive semantics, not a reduced token
/// substitute. Nested incomplete declarations should return None.
pub trait CompletedUnits: Builder {
    type Unit: ?Sized;
    fn is_emission_boundary(&self, rule: RuleId) -> bool;
    fn completed_unit(&self, boundary: Boundary) -> Result<Option<&Self::Unit>, BuildError>;
}

/// A transactional consumer of complete semantic units. Checkpoints must cover
/// every mutation (instructions, resolver/scope/type/import state and errors),
/// not merely code length. `finish` may validate/finalize private state but must
/// remain rollback-safe if it returns Err. Publish only after the parser returns
/// Ok. A compiler without such a journal must queue units until a sealed entry.
pub trait Sink<Unit: ?Sized> {
    type Checkpoint: Clone;
    fn checkpoint(&self) -> Self::Checkpoint;
    fn rollback(&mut self, mark: Self::Checkpoint);
    fn emit(&mut self, unit: &Unit, boundary: Boundary, source: &str) -> Result<(), BuildError>;
    fn finish(&mut self, _source: &str, _consumed: usize) -> Result<(), BuildError> {
        Ok(())
    }
}

impl<B: CompletedUnits> CompletedUnits for &mut B {
    type Unit = B::Unit;
    fn is_emission_boundary(&self, rule: RuleId) -> bool {
        (**self).is_emission_boundary(rule)
    }
    fn completed_unit(&self, boundary: Boundary) -> Result<Option<&Self::Unit>, BuildError> {
        (**self).completed_unit(boundary)
    }
}
impl<Unit: ?Sized, S: Sink<Unit>> Sink<Unit> for &mut S {
    type Checkpoint = S::Checkpoint;
    fn checkpoint(&self) -> Self::Checkpoint {
        (**self).checkpoint()
    }
    fn rollback(&mut self, mark: Self::Checkpoint) {
        (**self).rollback(mark)
    }
    fn emit(&mut self, unit: &Unit, boundary: Boundary, source: &str) -> Result<(), BuildError> {
        (**self).emit(unit, boundary, source)
    }
    fn finish(&mut self, source: &str, consumed: usize) -> Result<(), BuildError> {
        (**self).finish(source, consumed)
    }
}

/// Explicit adapter; ordinary Builders and PEG grammars never select it.
/// The adapter's journal combines otherwise unrelated builder/sink cursors.
/// Omission permissions that could hide newly observed emission boundaries
/// default back to reference matching; explicit Pratt semantics remain enabled.
/// Parser hooks use checked methods and return sink errors normally. Explicit
/// legacy infallible finish/commit calls panic on a sink error; use checked
/// parser entry points to retain transactional error recovery.
pub struct StreamingBuilder<B: CompletedUnits, S: Sink<B::Unit>> {
    builder: B,
    sink: S,
    undo: Vec<(usize, S::Checkpoint)>,
}
impl<B: CompletedUnits, S: Sink<B::Unit>> StreamingBuilder<B, S> {
    pub fn new(builder: B, sink: S) -> Self {
        Self {
            builder,
            sink,
            undo: Vec::new(),
        }
    }
    pub fn builder(&self) -> &B {
        &self.builder
    }
    pub fn sink(&self) -> &S {
        &self.sink
    }
    pub fn into_parts(self) -> (B, S) {
        (self.builder, self.sink)
    }
    fn record(&mut self) {
        self.undo
            .push((self.builder.checkpoint(), self.sink.checkpoint()));
    }
    fn emit_completed(&mut self, rule: RuleId, span: Span, source: &str) -> Result<(), BuildError> {
        if self.builder.is_emission_boundary(rule) {
            let boundary = Boundary { rule, span };
            if let Some(unit) = self.builder.completed_unit(boundary)? {
                self.sink.emit(unit, boundary, source)?;
            }
        }
        Ok(())
    }
}

impl<B: CompletedUnits, S: Sink<B::Unit>> Builder for StreamingBuilder<B, S> {
    fn supports_choice_heads(&self) -> bool {
        self.builder.supports_choice_heads()
    }
    fn supports_shared_wrapper_prefix_at(
        &self,
        _a: RuleId,
        _b: RuleId,
        _o: usize,
        _s: &str,
    ) -> bool {
        false
    }
    fn supports_shared_wrapper_prefix(&self, _a: RuleId, _b: RuleId) -> bool {
        false
    }
    fn failure_prefix_hooks_are_inert(&self, rule: RuleId) -> bool {
        !self.builder.is_emission_boundary(rule)
            && self.builder.failure_prefix_hooks_are_inert(rule)
    }
    fn supports_owned_capture_boundary(&self, _rule: RuleId) -> bool {
        false
    }
    fn owned_capture_end(&self, rule: RuleId, start: usize, source: &str) -> Option<usize> {
        self.builder.owned_capture_end(rule, start, source)
    }
    fn observes_capture(&self, rule: RuleId) -> bool {
        self.builder.observes_capture(rule)
    }
    fn supports_shared_inner_prefix(&self, _r: RuleId) -> bool {
        false
    }
    fn supports_shared_group_prefix(&self, _a: RuleId, _b: RuleId) -> bool {
        false
    }
    fn supports_shared_prefix(&self, _rule: RuleId) -> bool {
        false
    }
    fn empty_success_is_inert(&self, _r: usize) -> bool {
        false
    }
    fn empty_suffix_is_inert(&self, _r: usize) -> bool {
        false
    }
    fn supports_unobserved_failure_prefixes(&self) -> bool {
        self.builder.supports_unobserved_failure_prefixes()
    }
    fn supports_pratt(&self, rule: RuleId) -> bool {
        self.builder.supports_pratt(rule)
    }
    fn supports_pratt_at(&self, rule: RuleId, offset: usize, source: &str) -> bool {
        self.builder.supports_pratt_at(rule, offset, source)
    }
    fn accepts_pratt(&self, rule: RuleId, span: Span) -> bool {
        self.builder.accepts_pratt(rule, span)
    }
    fn suppress_pratt_captures(&self, rule: RuleId) -> bool {
        self.builder.suppress_pratt_captures(rule)
    }
    fn opaque_rule(&self, _rule: RuleId) -> bool {
        false
    }
    fn prune_capture_children(&self, rule: RuleId, span: Span) -> bool {
        self.builder.prune_capture_children(rule, span)
    }

    fn observes_rule(&self, rule: RuleId) -> bool {
        self.builder.observes_rule(rule) || self.builder.is_emission_boundary(rule)
    }
    fn checkpoint(&self) -> usize {
        self.undo.len()
    }
    fn rollback(&mut self, mark: usize) {
        assert!(mark <= self.undo.len(), "streaming checkpoint is not live");
        while self.undo.len() > mark {
            let (builder, sink) = self.undo.pop().unwrap();
            self.builder.rollback(builder);
            self.sink.rollback(sink);
        }
    }
    fn rollback_pratt(&mut self, mark: usize, rule: RuleId, span: Span) {
        assert!(mark <= self.undo.len(), "streaming checkpoint is not live");
        if let Some((builder, sink)) = self.undo.get(mark).cloned() {
            self.undo.truncate(mark);
            self.sink.rollback(sink.clone());
            self.builder.rollback_pratt(builder, rule, span);
            // A specialized builder can retain rollback-safe observational
            // receipts. Preserve their new journal cursor for outer rollback.
            if self.builder.checkpoint() != builder {
                self.undo.push((builder, sink));
            }
        }
    }
    fn begin(&mut self, rule: RuleId, offset: usize) {
        self.record();
        self.builder.begin(rule, offset);
    }
    fn finish(&mut self, rule: RuleId, span: Span, source: &str) {
        self.try_finish(rule, span, source)
            .expect("use checked streaming finish for fallible sinks");
    }
    fn try_begin(&mut self, rule: RuleId, offset: usize) -> Result<(), BuildError> {
        self.record();
        self.builder.try_begin(rule, offset)
    }
    fn try_finish(&mut self, rule: RuleId, span: Span, source: &str) -> Result<(), BuildError> {
        self.record();
        self.builder.try_finish(rule, span, source)?;
        self.emit_completed(rule, span, source)
    }
    fn wants_completed_capture(&self, rule: RuleId) -> bool { self.builder.wants_completed_capture(rule) }
    fn try_completed_capture(&mut self, rule: RuleId, span: Span,
        tree: crate::ParseTree<'_, '_>) -> Result<(), BuildError> {
        self.record();
        let source = tree.source();
        self.builder.try_completed_capture(rule, span, tree)?;
        // CompletedUnits may expose a typed statement at this richer boundary.
        self.emit_completed(rule, span, source)
    }
    fn try_finish_capture(
        &mut self,
        rule: RuleId,
        span: Span,
        source: &str,
        capture: Option<crate::tree::CaptureHandle>,
    ) -> Result<(), BuildError> {
        self.record();
        self.builder
            .try_finish_capture(rule, span, source, capture)?;
        if self.builder.wants_completed_capture(rule) {
            // The richer callback below constructs this boundary's value.
            // Do not emit the preceding statement's still-live value here.
            Ok(())
        } else {
            self.emit_completed(rule, span, source)
        }
    }
    fn reduce(&mut self, rule: RuleId, fixity: crate::pratt::Fixity, span: Span) {
        self.record();
        self.builder.reduce(rule, fixity, span);
    }
    fn try_reduce(
        &mut self,
        rule: RuleId,
        fixity: crate::pratt::Fixity,
        span: Span,
    ) -> Result<(), BuildError> {
        self.record();
        self.builder.try_reduce(rule, fixity, span)
    }
    fn try_pratt_boundary(&mut self, rule: RuleId, offset: usize) -> Result<(), BuildError> {
        self.record();
        self.builder.try_pratt_boundary(rule, offset)
    }
    fn committed(&mut self, source: &str, consumed: usize) {
        self.try_committed(source, consumed)
            .expect("use checked streaming commit for fallible sinks");
    }
    fn try_committed(&mut self, source: &str, consumed: usize) -> Result<(), BuildError> {
        self.record();
        self.sink.finish(source, consumed)?;
        self.builder.try_committed(source, consumed)?;
        // Old entry marks cease to be valid only after checked success.
        self.undo.clear();
        Ok(())
    }
}
