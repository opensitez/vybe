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
/// The cursor changes only through mutating hooks or rollback; selection queries
/// and `checkpoint` itself must not change semantic state. The matcher may reuse
/// the current cursor and skip rollback when no journaled mutation occurred.
pub trait Builder {
    /// Opt into certified non-consuming failed-head omission at explicitly
    /// selected ordered choices. Observed prefix rules retain original execution;
    /// captures, successful hooks and branch rollback still use original bodies.
    /// Selection/checkpoint queries must be pure. Default preserves PEG callbacks.
    fn supports_choice_heads(&self)->bool{false}

    /// Opt into retaining an equal prefix under unobserved normal wrappers.
    /// Captures keep their positions; only the inert wrapper's rule tag changes
    /// when the branch suffix fails. Selection must not mutate semantic state.
    fn supports_shared_wrapper_prefix(&self, _: RuleId, _: RuleId) -> bool { false }
    /// Optional conservative source admission. Unknown forms should decline;
    /// the original grammar still validates every accepted operand.
    fn supports_shared_wrapper_prefix_at(&self, branch: RuleId, fallback: RuleId,
        _: usize, _: &str,
    ) -> bool { self.supports_shared_wrapper_prefix(branch, fallback) }

    /// Opt into retaining deterministic successful-prefix semantic state instead
    /// of rolling it back and replaying it through an inert wrapper. Selection
    /// must not mutate state. Capture-token issuance prevents unsafe relocation.
    /// Default preserves ordinary PEG/callback behavior.
    fn supports_shared_inner_prefix(&self,_:RuleId)->bool{false}
    fn supports_shared_group_prefix(&self,_:RuleId,_:RuleId)->bool{false}
    fn supports_shared_prefix(&self, _: RuleId) -> bool { false }

    /// Called once after successful matching of the caller-selected entry rule,
    /// when speculative checkpoints are
    /// no longer used. Bindings may transfer owned values, discard undo storage,
    /// and normalize context-free AST nodes here. Failed matches never call this
    /// hook. Scope-dependent normalization must retain its contextual boundary.
    fn committed(&mut self, _: &str, _: usize) {}
    /// Checked finalization after the entry matches. A failure restores the
    /// initial build checkpoint just like other checked semantic errors.
    /// Keep state rollback-safe until this returns successfully. Existing
    /// builders retain their infallible committed callback by default.
    fn try_committed(&mut self, source: &str, consumed: usize) -> Result<(), BuildError> {
        self.committed(source, consumed);
        Ok(())
    }
    /// Bind begin/finish hooks only for rules that contribute semantic state.
    /// Unobserved wrappers retain matching/captures and do not suppress hooks
    /// in their children. The default preserves existing builder behavior.
    fn observes_rule(&self, _: RuleId) -> bool { true }
    /// Explicit permission to omit pure unobserved failed-rule checkpoints
    /// and internal callback queries when all descendant rules are unobserved.
    /// Default preserves custom builder callback policies.
    fn supports_unobserved_failure_prefixes(&self) -> bool { false }
    /// Allow this rule's callbacks to be omitted on a generated certified
    /// zero-consumption failure. All reached rules are admitted separately.
    /// Grant only when that prefix has no checked errors, semantic mutations
    /// or external observations to preserve. Successful matches and ordinary
    /// failures retain their hooks. Program observers have separate admission.
    fn failure_prefix_hooks_are_inert(&self, _: RuleId) -> bool { false }


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
    /// Finish a hybrid semantic value with its finalized parse-local capture.
    /// Default builders retain the exact existing finish policy. Opt-in users
    /// must journal tokens with values and retain any captures they later use.
    /// Explicitly request a finalized capture token only for rules whose
    /// semantic consumer needs it. Ordinary builders construct no tokens.
    /// Explicit complete-owned rule boundary. The selected finish owns the
    /// original matched source span and preserves its retained root anchor.
    /// Descendant semantic hooks must be rollback-safe; no capture hooks may
    /// observe omitted nodes. Defaults preserve PEG and existing builders.
    fn supports_owned_capture_boundary(&self,_:RuleId)->bool{false}
    /// Predict the exact original rule end without semantic mutation. A mismatch
    /// replays the original body with captures from its post-begin checkpoint.
    fn owned_capture_end(&self,_:RuleId,_:usize,_:&str)->Option<usize>{None}
    fn observes_capture(&self, _: RuleId) -> bool { false }
    /// Request an owned snapshot of this completed capture subtree while the
    /// parser is still matching. Default builders incur no copy or callback.
    /// Mutations remain speculative and obey checkpoint/rollback; consumers
    /// must not publish externally executable code before source-unit commit.
    fn wants_completed_capture(&self, _: RuleId) -> bool { false }
    fn try_completed_capture(&mut self, _: RuleId, _: Span,
        _: crate::ParseTree<'_, '_>) -> Result<(), BuildError> { Ok(()) }

    fn try_finish_capture(&mut self, rule: RuleId, span: Span, source: &str,
        _: Option<crate::tree::CaptureHandle>) -> Result<(), BuildError> {
        self.try_finish(rule, span, source)
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
    /// Reports an operand boundary after the island's implicit trivia has
    /// matched, before the next operator candidate. A binding can extend its
    /// top value's span without rescanning trivia. The edit must be journaled:
    /// failed operators and Preserve-trivia islands can restore this boundary.
    /// Suppressed in predicates, trivia and nonsemantic recognition.
    fn try_pratt_boundary(&mut self, _: RuleId, _: usize) -> Result<(), BuildError> {
        Ok(())
    }
    /// Opt into a configured island whose atom hooks push one semantic value.
    /// Operator rule hooks are suppressed; `reduce` owns their AST semantics.
    /// Other rule wrappers may disappear on this path. Ordinary capture parsing
    /// retains the PEG backend; `build_capture_program` explicitly opts into
    /// semantic islands alongside captures for contextual lowering.
    fn supports_pratt(&self, _: RuleId) -> bool {
        false
    }
    /// Refine island selection for the current input position. A conservative
    /// binding can keep contextual syntax on PEG before speculative AST work.
    /// Matching still runs the configured grammar; this is only backend selection.
    fn supports_pratt_at(&self, rule: RuleId, _: usize, _: &str) -> bool {
        self.supports_pratt(rule)
    }
    /// After successful reductions, decide whether a hybrid island produced a
    /// complete semantic value. Returning false restores its input, captures,
    /// grammar stack and builder journal, then matches the original PEG body
    /// without semantic hooks. Its ordinary captures remain available for
    /// contextual lowering. This is not a syntax alternative or error recovery.
    /// Only hybrid capture builds consult this hook; AST-only builds cannot
    /// fall back to a contextual capture walker.
    fn accepts_pratt(&self, _: RuleId, _: Span) -> bool { true }
    /// Opt into omitting descendant captures while a selected hybrid Pratt
    /// island constructs an owned value. Its own capture, when declared,
    /// remains the contextual anchor. If `accepts_pratt` returns false, the
    /// matcher restores its journals and recreates the original PEG captures.
    /// Return true only when a successful owned value replaces all descendant
    /// capture use; ordinary PEG regions and AST-only matching are unchanged.
    fn suppress_pratt_captures(&self, _: RuleId) -> bool { false }
    /// Permit a certified atomic rule to match without inner rule callbacks.
    /// The rule's own begin/finish hooks still run. Return true only when all
    /// inner hooks are semantically irrelevant and the complete source span
    /// suffices to construct its value (for example a number or identifier).
    /// Uncertified bodies retain ordinary matching and callbacks.
    fn opaque_rule(&self, _: RuleId) -> bool {
        false
    }
    /// After a semantic finish, discard a rule's child captures when the
    /// builder already owns its complete value. The rule capture remains as
    /// an anchor for a contextual frontend. Ordinary capture parsing is unchanged.
    fn prune_capture_children(&self, _: RuleId, _: Span) -> bool { false }
    fn reduce(&mut self, _: RuleId, _: crate::pratt::Fixity, _: Span) {}
    fn checkpoint(&self) -> usize;
    fn rollback(&mut self, mark: usize);
    /// Restore a rejected, successfully matched hybrid Pratt island before its
    /// original PEG body is recaptured without semantic hooks. Default is the
    /// ordinary rollback policy. Opt-in owners may retain only certified
    /// observational receipts for this exact island, reallocated in their
    /// existing journal after rollback. Retained entries must remain removable
    /// by an enclosing rollback or failed replay; never retain AST values,
    /// capture handles, unfinished frames, execution effects or scope mutations.
    fn rollback_pratt(&mut self, mark: usize, _: RuleId, _: Span) {
        self.rollback(mark);
    }
    fn begin(&mut self, rule: RuleId, offset: usize);
    fn finish(&mut self, rule: RuleId, span: Span, source: &str);
    /// Pure opt-in: omitting this unobserved successful-empty rule, its pure
    /// failed literal attempts and checkpoint/rollback queries preserves every
    /// semantic state change and external observation. The default denies it.
    fn empty_success_is_inert(&self,_:usize)->bool{false}
    /// Pure opt-in: omitting the selected failed optional prefix and its checkpoint/rollback/query work preserves all observations/state.
    fn empty_suffix_is_inert(&self,_:usize)->bool{false}
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
    fn supports_choice_heads(&self)->bool{(**self).supports_choice_heads()}

 fn supports_shared_wrapper_prefix_at(&self,a:RuleId,b:RuleId,o:usize,s:&str)->bool{(**self).supports_shared_wrapper_prefix_at(a,b,o,s)}
 fn supports_shared_wrapper_prefix(&self,a:RuleId,b:RuleId)->bool{(**self).supports_shared_wrapper_prefix(a,b)}

    fn failure_prefix_hooks_are_inert(&self, rule: RuleId) -> bool { (**self).failure_prefix_hooks_are_inert(rule) }

    fn supports_owned_capture_boundary(&self,rule:RuleId)->bool{(**self).supports_owned_capture_boundary(rule)}
    fn owned_capture_end(&self,rule:RuleId,start:usize,source:&str)->Option<usize>{(**self).owned_capture_end(rule,start,source)}
    fn observes_capture(&self, rule: RuleId) -> bool { (**self).observes_capture(rule) }
    fn wants_completed_capture(&self, rule: RuleId) -> bool { (**self).wants_completed_capture(rule) }
    fn try_completed_capture(&mut self, rule: RuleId, span: Span, tree: crate::ParseTree<'_, '_>) -> Result<(), BuildError> {
        (**self).try_completed_capture(rule, span, tree)
    }


    fn try_finish_capture(&mut self, rule: RuleId, span: Span, source: &str,
        capture: Option<crate::tree::CaptureHandle>) -> Result<(), BuildError> {
        (**self).try_finish_capture(rule, span, source, capture)
    }

    fn committed(&mut self, source: &str, consumed: usize) {
        (**self).committed(source, consumed);
    }
    fn try_committed(&mut self, source: &str, consumed: usize) -> Result<(), BuildError> {
        (**self).try_committed(source, consumed)
    }
    fn observes_rule(&self, rule: RuleId) -> bool { (**self).observes_rule(rule) }
    fn supports_shared_inner_prefix(&self,r:RuleId)->bool{(**self).supports_shared_inner_prefix(r)}
    fn supports_shared_group_prefix(&self,a:RuleId,b:RuleId)->bool{(**self).supports_shared_group_prefix(a,b)}
    fn supports_shared_prefix(&self, rule: RuleId) -> bool { (**self).supports_shared_prefix(rule) }
    fn empty_success_is_inert(&self,r:usize)->bool{(**self).empty_success_is_inert(r)}
    fn empty_suffix_is_inert(&self,r:usize)->bool{(**self).empty_suffix_is_inert(r)}
    fn supports_unobserved_failure_prefixes(&self) -> bool { (**self).supports_unobserved_failure_prefixes() }

    fn try_pratt_boundary(&mut self, rule: RuleId, offset: usize) -> Result<(), BuildError> {
        (**self).try_pratt_boundary(rule, offset)
    }
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
    fn supports_pratt_at(&self, rule: RuleId, offset: usize, source: &str) -> bool {
        (**self).supports_pratt_at(rule, offset, source)
    }
    fn accepts_pratt(&self, rule: RuleId, span: Span) -> bool {
        (**self).accepts_pratt(rule, span)
    }
    fn suppress_pratt_captures(&self, rule: RuleId) -> bool {
        (**self).suppress_pratt_captures(rule)
    }
    fn opaque_rule(&self, rule: RuleId) -> bool {
        (**self).opaque_rule(rule)
    }
    fn reduce(&mut self, rule: RuleId, fixity: crate::pratt::Fixity, span: Span) {
        (**self).reduce(rule, fixity, span);
    }
    fn prune_capture_children(&self, rule: RuleId, span: Span) -> bool {
        (**self).prune_capture_children(rule, span)
    }
    fn checkpoint(&self) -> usize {
        (**self).checkpoint()
    }
    fn rollback(&mut self, mark: usize) {
        (**self).rollback(mark);
    }
    fn rollback_pratt(&mut self, mark: usize, rule: RuleId, span: Span) {
        (**self).rollback_pratt(mark, rule, span);
    }
    fn begin(&mut self, rule: RuleId, offset: usize) {
        (**self).begin(rule, offset);
    }
    fn finish(&mut self, rule: RuleId, span: Span, source: &str) {
        (**self).finish(rule, span, source);
    }
}
