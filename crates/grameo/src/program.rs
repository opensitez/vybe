//! Runtime contract for resolved IR and build-time generated parser programs.
//! Names are metadata; source matching uses dense rule/expression IDs.
use crate::grammar::{ExprId, RuleId};
use crate::{Builtin, CompiledGrammar, ExprKind, Reference, RuleMode};

#[derive(Debug, Clone, Copy)]
pub struct RuleSpec<'a> {
    pub name: &'a str,
    pub mode: RuleMode,
    pub expression: ExprId,
    pub scope: usize,
}

#[derive(Debug, Clone, Copy, Default)]
pub struct Scope {
    pub whitespace: Option<RuleId>,
    pub comment: Option<RuleId>,
    pub whitespace_class: Option<crate::lexical::ByteClass>,
    pub whitespace_prefix_class: Option<crate::lexical::ByteClass>,
}

#[derive(Debug, Clone, Copy)]
pub struct BoundOperator {
    pub rule: RuleId,
    pub precedence: u16,
    pub fixity: crate::pratt::Fixity,
}
#[derive(Debug, Clone, Copy)]
pub struct SourcePratt<'a> {
    pub atom: RuleId,
    pub operators: &'a [BoundOperator],
    pub trailing_trivia: crate::islands::TrailingTrivia,
}
#[derive(Debug, Clone)]
pub(crate) struct OwnedPratt {
    pub operator_rhs_atoms: Vec<(RuleId, RuleId)>,
    pub rhs_atom: Option<RuleId>,
    pub rhs_operator: Option<RuleId>,
    pub operator_trivia: Option<RuleId>,
    pub atom: RuleId,
    pub operators: Vec<BoundOperator>,
    pub trailing_trivia: crate::islands::TrailingTrivia,
}

#[derive(Debug, Clone, Copy)]
pub enum FragmentExpected<'a> {
    Builtin(Builtin),
    Literal(&'a str),
    Range(char, char),
}
#[derive(Debug, Clone, Copy)]
pub struct FragmentLiteral<'a> {
    pub text: &'a str,
    pub insensitive: bool,
    pub success_steps: usize,
}
#[derive(Debug, Clone, Copy)]
pub struct FragmentGuard<'a> {
    pub rule: RuleId,
    pub alternatives: &'a [FragmentLiteral<'a>],
    pub candidates: &'a [&'a [u16]; 256],
    pub failure_steps: usize,
    pub boundary: &'a [u16; 256],
    pub choice_failure_fixed: usize,
    pub boundary_fixed: usize,
    pub boundary_failure_steps: usize,
}
#[derive(Debug, Clone, Copy)]
pub struct FragmentSuffix<'a> {
    pub alternatives: &'a [FragmentLiteral<'a>],
    pub failure_steps: usize,
    pub expected: &'a [FragmentExpected<'a>],
}
#[derive(Debug, Clone, Copy)]
pub struct AtomicLiteral<'a> {
    pub text: &'a str,
    pub insensitive: bool,
    pub boundary: &'a [u16; 256],
    pub literal_failure_steps: usize,
    pub boundary_fixed_steps: usize,
    pub boundary_failure_steps: usize,
}
#[derive(Debug, Clone, Copy)]
pub struct AtomicFragment<'a> {
    pub literal: Option<&'a AtomicLiteral<'a>>,
    pub guard: Option<&'a FragmentGuard<'a>>,
    pub suffix: Option<&'a FragmentSuffix<'a>>,
    pub first: &'a [u16; 256],
    pub tail: &'a [u16; 256],
    pub fixed_steps: usize,
    pub failure_guard_fixed: usize,
    pub failure_head_fixed: usize,
    pub head_failure_steps: usize,
    pub head_failure_safe: bool,
    pub head_expected: &'a [FragmentExpected<'a>],
    pub expected: &'a [FragmentExpected<'a>],
}
#[derive(Debug, Clone, Copy)]
pub enum Instruction<'a> {
    Literal {
        text: &'a str,
        insensitive: bool,
    },
    Range {
        start: char,
        end: char,
    },
    Call(RuleId),
    Builtin(Builtin),
    Sequence {
        left: ExprId,
        right: ExprId,
    },
    Choice {
        left: ExprId,
        right: ExprId,
    },
    Repeat {
        expression: ExprId,
        min: u32,
        max: Option<u32>,
    },
    Predicate {
        expression: ExprId,
        positive: bool,
    },
    Group(ExprId),
    Push(ExprId),
    PushLiteral(&'a str),
    PeekSlice {
        start: i32,
        end: Option<i32>,
    },
}

/// Outcome of a precomputed implicit trivia first-byte test.
#[derive(Debug, Clone, Copy)]
pub enum TriviaFailure {
    Unknown,
    Possible,
    Miss { steps: usize, depth: usize },
}

/// Immutable parser program. Generated implementations provide static Rust
/// dispatch/literals and need no grammar parsing, allocation, or name resolution
/// at source-parse startup. Implementations must provide valid arena IDs.
/// A scope word whose low six bits are reserved for private parser flags.
/// Generated code constructs these in constants, so unrepresentable scopes
/// are rejected during compilation rather than truncated during parsing.
#[doc(hidden)]
#[derive(Clone, Copy)]
pub struct CompactScope(usize);
impl CompactScope {
    pub const fn new(scope: usize) -> Self {
        assert!(
            scope <= usize::MAX >> 6,
            "scope does not fit compact capture environment"
        );
        Self(scope << 6)
    }
    pub(crate) fn word(self) -> usize {
        self.0
    }
}

/// Generated proof that a choice's left subtree fails without state changes
/// for bytes outside `bytes`. Only certified rule calls, groups, choices and left-failing sequences are used.
#[derive(Debug, Clone, Copy)]
pub struct NativeChoicePrefix<'a> {
    pub bytes: [u64; 4],
    /// All left-subtree tasks; caller adds its Choice continuation.
    pub steps: usize,
    pub depth: usize,
    pub diagnostic_complete: bool,
    pub events: &'a [NativeChoiceEvent],
}
/// Generated left-sequence spine ending in a rule call. Engine preserves frames.
#[derive(Clone, Copy, Debug)]
pub struct NativeRuleHead<'a> {
    pub sequences: &'a [(ExprId, ExprId)],
    pub call_expression: ExprId,
    pub callee: RuleId,
}

/// Original observer order for a certified first-byte miss.
#[derive(Debug, Clone, Copy)]
pub enum NativeChoiceEvent {
    Expression(ExprId),
    Call { expression: ExprId, rule: RuleId },
}

/// Generated opt-in zero-consumption failed rule path. Caller Eval is excluded.
/// Work includes EnterRule + FinishRule; expected terminals retain source order.
#[derive(Debug)]
pub struct NullableFailurePrefix<'a> {
    pub steps: usize,
    pub depth: usize,
    pub expected: &'a [FragmentExpected<'a>],
    pub unobserved_rules: &'a [RuleId],
}

/// Checked PEG shape for a shared successful expression prefix. Providers must
/// retain its rule/instruction geometry; runtime revalidates direct override edges.
#[derive(Debug, Clone, Copy)]
pub struct SharedPrefixChoice {
    pub target: RuleId, pub target_expression: ExprId,
    pub branch_expression: ExprId, pub branch_sequence: ExprId,
    pub left_choice: ExprId, pub wrapper_call: ExprId, pub right: ExprId,
    pub scope: usize, pub prefix_rule: RuleId, pub suffix_rule: Option<RuleId>,
    pub prior: ExprId, pub wrapper: RuleId, pub prefix: ExprId, pub suffix: ExprId,
}

/// Checked `branch(prefix ~ suffix) | fallback(prefix)` capture-preserving choice.
/// The matcher rechecks the current IR, modes, scopes and observer policies.
#[derive(Debug, Clone, Copy)]
pub struct SharedWrapperPrefix {
    pub choice: ExprId,
    pub branch: RuleId,
    pub fallback_rule: RuleId,
    pub prefix: ExprId,
    pub suffix: ExprId,
    pub fallback: ExprId,
}
/// Opted-in delimited generator/group choice; original PEG geometry stays authoritative.
#[derive(Debug,Clone,Copy)]
pub struct SharedGroupPrefix {
 pub target: usize,
 pub target_expression: usize,
 pub choice: usize,
 pub left: usize,
 pub right: usize,
 pub left_inner: usize,
 pub open: usize,
 pub branch_call: usize,
 pub close: usize,
 pub right_inner: usize,
 pub fallback_open: usize,
 pub optional: usize,
 pub fallback_close: usize,
 pub branch: usize,
 pub branch_expression: usize,
 pub fallback_rule: usize,
 pub fallback_expression: usize,
 pub prefix_rule: usize,
 pub prefix: usize,
 pub branch_tail: usize,
 pub fallback_inner: usize,
 pub fallback_prefix: usize,
 pub commas: usize,
 pub trailing: usize,
}
/// Immutable generated geometry for an explicitly owned normal rule.
#[derive(Clone,Copy,Debug)]
pub struct OwnedCaptureBoundary<'a> {pub expression:ExprId,pub scope:usize,pub capture_rules:&'a[RuleId]}
/// Explicit capture-preserving shared prefix inside one rule.
#[derive(Debug,Clone,Copy,PartialEq,Eq)]
pub struct SharedInnerPrefix {
pub target:usize,
pub target_expression:usize,
pub choice:usize,
pub left:usize,
pub right:usize,
pub prefix_rule:usize,
pub prefix:usize,
pub first:usize,
pub second:usize,
pub third:usize,
pub kind:usize,
}
/// Opted-in silent literal repetition. An absent first byte proves zero-length
/// success without entering the repetition. Immutable geometry must be retained.
#[derive(Debug,Clone,Copy)]
pub struct EmptySuccessHead{pub expression:ExprId,pub scope:usize,pub first:[u64;4]}
pub trait Program: std::fmt::Debug {
    /// Opted-in generated geometry agrees with current rule/body/scope and
    /// descendant capture-observer rules. Forwarding wrappers default to false.
    fn owned_capture_boundary_metadata_is_current(&self)->bool{false}
    fn owned_capture_boundary(&self,_:RuleId)->Option<OwnedCaptureBoundary<'_>>{None}
    /// Opt-in embedding storage policy. Unmodified/forwarding programs retain
    /// independent units unless they explicitly carry this contract.
    fn code_space(&self) -> crate::source_units::CodeSpace { crate::source_units::CodeSpace::Unit }
    /// Explicit source-load policy for an embedding's transactional AST hook.
    /// Metadata only: the parser never resolves paths or executes a load.
    fn source_unit_binding(&self, _: RuleId, _: Option<RuleId>) -> Option<crate::source_units::Binding<'_>> { None }
    /// Generated immutable choice metadata agrees with every current instruction,
    /// literal, rule/scope/mode and SourcePratt setting. Forwarding wrappers must
    /// independently establish this contract; forwarding metadata alone is insufficient.
    fn choice_heads_metadata_is_current(&self)->bool{false}
    /// Cheap generated selected-root admission. Default PEG programs have no targets.
    fn choice_head_branch_metadata_is_current(&self) -> bool { false }
    /// Optional per-alternative conservative hook proof; None retains union behavior.
    fn choice_head_branch_rules(&self, _expression: ExprId, _branch: usize) -> Option<&[RuleId]> { None }
    fn choice_heads_candidate(&self,_:ExprId)->bool{false}
    fn choice_heads(&self,_:ExprId)->Option<&ChoiceHeads>{None}

 fn shared_inner_prefix(&self,_:ExprId)->Option<SharedInnerPrefix>{None}
 fn shared_group_prefix(&self,_:ExprId)->Option<SharedGroupPrefix>{None}
    /// Immutable resolved/generated wrapper geometry agrees with current instructions.
    /// Custom/forwarding programs must independently certify this; default revalidates.
    fn shared_wrapper_prefix_metadata_is_current(&self)->bool{false}
 fn shared_wrapper_prefix(&self,_:ExprId)->Option<SharedWrapperPrefix>{None}

    /// Explicit permission to omit diagnostic matcher observations in opted-in
    /// certificates. Both observer callbacks must be no-ops. Forwarding wrappers
    /// retain their events by default, even when they forward certificate metadata.
    fn match_observers_are_noop(&self) -> bool { false }
    /// Disabled for ordinary generation and custom observational wrappers.
    /// Metadata alone does not permit omitting matcher observations. Unicode and
    /// unknown contexts return None.
    fn nullable_failure_prefix(&self, _: RuleId, _: u8, _: Option<u8>) -> Option<&NullableFailurePrefix<'_>> { None }

    /// Explicit generated-native opt-in. Ordinary capture parsing never uses
    /// this path; custom programs and standard generation leave it disabled.
    fn shared_prefix_choice(&self, _: ExprId) -> Option<SharedPrefixChoice> { None }
    fn native_hybrid_choices(&self) -> bool {
        false
    }
    fn native_hybrid_choice_prefix(&self, _: ExprId) -> Option<&NativeChoicePrefix<'_>> {
        None
    }
    /// Explicit native left-sequence spine. All semantic hooks and frames remain.
    fn native_hybrid_rule_head(&self, _: RuleId) -> Option<&NativeRuleHead<'_>> {
        None
    }

    #[inline(always)]
    fn delimited_fragment(&self, _: RuleId) -> Option<crate::quoted::DelimitedFragment<'_>> { None }
    #[inline(always)]
    fn quoted_fragment(&self, _: RuleId) -> Option<crate::quoted::Fragment> { None }
    /// Certified silent ASCII terminal repetition, without rule calls or effects.
    fn whitespace_run_class(&self, _: usize) -> Option<crate::lexical::ByteClass> { None }

    /// Forwarding wrappers must certify scanner geometry against current bodies.
    fn native_token_metadata_is_current(&self) -> bool { false }
    /// Generated opt-in success-only lexical recognizer. It must preserve PEG greediness.
    fn native_token(&self, _rule: RuleId, _source: &str, _offset: usize, _remaining_depth: usize) -> Option<usize> { None }
    /// Certified lexical result: None retains PEG; Some(None) proves failure.
    /// Legacy positive-only providers never claim a negative result.
    fn native_token_result(&self, rule: RuleId, source: &str, offset: usize, remaining_depth: usize) -> Option<Option<usize>> {
        self.native_token(rule, source, offset, remaining_depth).map(Some)
    }

    fn atomic_fragments(&self) -> Option<()> {
        None
    }
    /// Explicit partial ASCII table policy; absent by default so tracing wrappers
    /// retain original intrinsic observations. Unsupported lanes replay whole rules.
    fn partial_atomic_fragment(&self, _rule: RuleId) -> Option<&AtomicFragment<'_>> { None }
    fn atomic_fragment(&self, _rule: RuleId) -> Option<&AtomicFragment<'_>> {
        None
    }
    /// Internal capture dispatch. Custom programs retain the wide environment.
    #[doc(hidden)]
    fn capture_parse<'g, 'i>(
        &'g self,
        entry: &str,
        source: &'i str,
        options: crate::MatchOptions,
    ) -> Result<crate::ParseTree<'g, 'i>, crate::ParseError>
    where
        Self: Sized,
    {
        crate::engine::parse_wide_program(self, entry, source, options)
    }
    /// Generated-only statically checked scope words indexed by scope ID. Default custom capture
    /// dispatch never calls this method; no arbitrary custom scope is packed.
    #[doc(hidden)]
    fn compact_capture_scope(&self, _: RuleId) -> CompactScope {
        panic!("program has not certified compact capture scopes")
    }
    /// Optional observational profiling; default programs compile these out.
    /// Flags encode atomicity, predicate/trivia suppression and capture mode.
    #[doc(hidden)]
    fn observe_rule_attempt(&self, _: RuleId, _: usize, _: u8) {}
    #[doc(hidden)]
    fn observe_expression(&self, _: ExprId, _: usize) {}

    /// First-priority single-byte repeat paths with exact successful task costs.
    fn leading_bytes(&self, _: ExprId) -> Option<crate::lexical::LeadingBytes> {
        None
    }
    /// Compact token execution for input-only atomic rule bodies. Unknown,
    /// recursive, stackful and capture-enabling paths retain normal dispatch.
    fn atomic_expression(&self, _: ExprId) -> bool {
        false
    }
    /// Ordered indices into a pure literal choice for this input first byte.
    /// Used only when diagnostics are suppressed. Indices retain grammar order;
    /// an empty slice certifies no match, while None retains full dispatch.
    fn terminal_candidates(&self, _: ExprId, _: Option<u8>) -> Option<&[usize]> {
        None
    }

    /// Optional generated eligibility guard. The default preserves every
    /// terminal-choice metadata query made by custom programs.
    #[inline]
    fn has_terminal_choice(&self, _: ExprId) -> bool {
        true
    }

    /// Pure terminal-choice dispatch, with original alternative priority,
    /// diagnostics and logical work costs. Unknown choices return None.
    fn terminal_choice(&self, _: ExprId) -> Option<crate::lexical::TerminalChoice<'_>> {
        None
    }
    /// Reject certified first-byte rule misses in capture mode, preserving
    /// terminal diagnostics and exact logical work/depth accounting.
    fn fast_rule_failures(&self) -> bool {
        true
    }
    /// Ordered terminal expression IDs for a certified rule failure.
    /// An empty slice retains ordinary dispatch when diagnostics are required.
    fn failure_terminals(&self, _: RuleId) -> &[ExprId] {
        &[]
    }
    /// Avoid scratch-stack round trips for immediate atomic token transitions.
    fn atomic_tail_dispatch(&self) -> bool {
        true
    }

    /// Dispatch continuation successors without task-stack round trips.
    fn tail_continuations(&self) -> bool {
        true
    }
    /// Execute immediate transitions in the iterative dispatch loop instead
    /// of pushing/popping redundant tasks. Logical work accounting is unchanged.
    fn tail_dispatch(&self) -> bool {
        true
    }
    /// Collapse certified failed implicit trivia in capture parsing, retaining
    /// the scheduled path's exact work/depth limits. Semantic builders keep
    /// their ordinary checkpoint/rollback callbacks.
    fn fast_trivia_failures(&self) -> bool {
        true
    }
    /// Bypass certified first-byte misses in capture-mode lookahead. Tight
    /// resource limits retain the original scheduled execution.
    fn fast_predicate_failures(&self) -> bool {
        true
    }
    /// Return the original full task cost only when this pure literal-choice
    /// expression is proven to fail without consuming input or mutating state.
    /// The cost must be positive and include the root Eval task. The engine
    /// uses it only where diagnostics/hooks are suppressed.
    fn literal_choice_failure(&self, _: ExprId, _: &str) -> Option<usize> {
        None
    }
    /// Precomputed cost/depth for an implicit trivia miss. Includes its two
    /// continuations (one if whitespace already finished), but not caller tasks.
    /// Unknown states or possible matches retain ordinary certificate dispatch.
    fn trivia_failure_cost(&self, _: usize, _: bool, _: Option<u8>) -> TriviaFailure {
        TriviaFailure::Unknown
    }
    /// Generated immutable prefixes can be read without a large by-value return.
    #[inline]
    fn borrowed_failure_prefixes(&self) -> bool {
        false
    }
    fn trivia_failure_prefix_ref(&self, _: RuleId) -> Option<&crate::lexical::FailurePrefix> {
        None
    }
    fn trivia_failure_prefix(&self, _: RuleId) -> Option<crate::lexical::FailurePrefix> {
        None
    }
    /// Enable certified atomic repetitions of one ASCII builtin. The scanner
    /// retains the ordinary execution path's work accounting and diagnostics.
    fn fast_ascii_repetitions(&self) -> bool {
        false
    }
    fn rule_id(&self, name: &str) -> Option<RuleId>;
    fn rule(&self, rule: RuleId) -> RuleSpec<'_>;
    fn instruction(&self, expression: ExprId) -> Instruction<'_>;
    fn source_pratt(&self, _: RuleId) -> Option<SourcePratt<'_>> {
        None
    }
    /// Explicit separator rule between an infix operator and its next atom.
    /// Default bindings use ordinary implicit trivia only.
    fn source_pratt_operator_trivia(&self, _: RuleId) -> Option<RuleId> { None }
    fn source_pratt_operator_rhs_atom(&self, _: RuleId, _: RuleId) -> Option<RuleId> { None }
    fn source_pratt_rhs_atom(&self, _: RuleId) -> Option<RuleId> { None }
    fn source_pratt_rhs_operator(&self, _: RuleId) -> Option<RuleId> { None }
    fn scope(&self, id: usize) -> Scope {
        assert_eq!(id, 0, "program has no such scope");
        Scope {
            whitespace: self.rule_id("WHITESPACE"),
            comment: self.rule_id("COMMENT"),
            whitespace_class: self.whitespace_class(),
            whitespace_prefix_class: self.whitespace_prefix_class(),
        }
    }
    /// Certified silent single-ASCII-scalar WHITESPACE, with no rule calls,
    /// captures, stack operations, predicates or implicit skipping.
    fn whitespace_class(&self) -> Option<crate::lexical::ByteClass> {
        None
    }
    /// Certified leading alternatives only; a miss must fall back to the full
    /// WHITESPACE rule. Later alternatives retain their original priority.
    fn whitespace_prefix_class(&self) -> Option<crate::lexical::ByteClass> {
        self.whitespace_class()
    }
    /// Cheap selected-rule lookup; metadata alone never admits a shortcut.
    fn empty_success_candidate(&self,_:RuleId)->bool{false}
    /// Grant only for unchanged certified rule instructions and modes/scopes.
    /// Accessors used by the certificate must be pure; their interpreter query
    /// counts/order may change. Forwarding or transformed programs must re-certify
    /// instead of forwarding this permission.
    fn empty_success_metadata_is_current(&self)->bool{false}
    fn empty_success_head(&self,_:RuleId)->Option<&EmptySuccessHead>{None}
    fn empty_suffix_candidate(&self,_:usize)->bool{false}
    /// Grant only for the unchanged certified instruction/mode/scope/Pratt geometry.
    /// All reached accessors must be pure: interpreter query counts/order may change.
    /// Forwarding/transformed programs must re-certify; forwarding metadata alone
    /// does not grant permission.
    fn empty_suffix_metadata_is_current(&self)->bool{false}
    fn empty_suffix_head(&self,_:usize)->Option<&EmptySuffixHead>{None}

}

impl Program for CompiledGrammar {
    fn code_space(&self) -> crate::source_units::CodeSpace { self.code_space.unwrap_or_default() }
    fn owned_capture_boundary_metadata_is_current(&self)->bool { !self.owned_capture_boundaries.is_empty() }
    fn owned_capture_boundary(&self,rule:RuleId)->Option<OwnedCaptureBoundary<'_>> {
        let capture_rules=self.owned_capture_boundaries.get(&rule)?;
        let definition=self.rule(rule);
        Some(OwnedCaptureBoundary{expression:definition.expression,scope:definition.scope,capture_rules})
    }

    fn source_unit_binding(&self, rule: RuleId, selector: Option<RuleId>) -> Option<crate::source_units::Binding<'_>> {
        self.source_units.iter().find(|binding| binding.rule == rule && binding.selector == selector).map(crate::source_units::OwnedBinding::borrowed)
    }
    fn shared_wrapper_prefix_metadata_is_current(&self)->bool{true}
    fn shared_wrapper_prefix(&self,id:ExprId)->Option<SharedWrapperPrefix>{self.shared_wrapper_prefixes.get(&id).copied()}
    fn match_observers_are_noop(&self) -> bool { true }
    fn empty_success_candidate(&self,r:RuleId)->bool{self.empty_success_heads.contains_key(&r)}
    fn empty_success_metadata_is_current(&self)->bool{true}
    fn empty_success_head(&self,r:RuleId)->Option<&EmptySuccessHead>{self.empty_success_heads.get(&r)}
    fn empty_suffix_candidate(&self,id:usize)->bool{self.empty_suffix_heads.contains_key(&id)}
    fn empty_suffix_metadata_is_current(&self)->bool{true}
    fn empty_suffix_head(&self,id:usize)->Option<&EmptySuffixHead>{self.empty_suffix_heads.get(&id)}

    fn shared_inner_prefix(&self,id:ExprId)->Option<SharedInnerPrefix>{self.shared_inner_prefixes.get(&id).copied()}
    fn shared_group_prefix(&self,id:ExprId)->Option<SharedGroupPrefix>{self.shared_group_prefixes.get(&id).copied()}
    fn shared_prefix_choice(&self, id: ExprId) -> Option<SharedPrefixChoice> { self.shared_prefixes.get(&id).copied() }
    fn leading_bytes(&self, expression: ExprId) -> Option<crate::lexical::LeadingBytes> {
        self.leading_bytes[expression]
    }
    fn atomic_expression(&self, expression: ExprId) -> bool {
        self.atomic_expressions[expression]
    }
    fn terminal_choice(&self, expression: ExprId) -> Option<crate::lexical::TerminalChoice<'_>> {
        let choice = self.terminal_choices[expression].as_ref()?;
        Some(crate::lexical::TerminalChoice {
            alternatives: &choice.alternatives,
            failure_steps: choice.failure_steps,
        })
    }
    fn failure_terminals(&self, rule: RuleId) -> &[ExprId] {
        &self.failure_terminals[rule]
    }
    fn literal_choice_failure(&self, expression: ExprId, source: &str) -> Option<usize> {
        let choice = self.literal_choices.get(&expression)?;
        for &id in &choice.alternatives {
            let Instruction::Literal { text, insensitive } = self.instruction(id) else {
                unreachable!()
            };
            if crate::lexical::literal_matches(source, text, insensitive) {
                return None;
            }
        }
        Some(choice.steps)
    }
    fn trivia_failure_prefix(&self, rule: RuleId) -> Option<crate::lexical::FailurePrefix> {
        self.trivia_failure[rule]
    }
    fn fast_ascii_repetitions(&self) -> bool {
        true
    }
    fn source_pratt_operator_rhs_atom(&self,target:RuleId,operator:RuleId)->Option<RuleId> { self.pratt.get(target)?.as_ref()?.operator_rhs_atoms.iter().find_map(|&(op,atom)|(op==operator).then_some(atom)) }
    fn source_pratt_rhs_operator(&self,rule:RuleId)->Option<RuleId> { self.pratt.get(rule)?.as_ref()?.rhs_operator }
    fn source_pratt_rhs_atom(&self,rule:RuleId)->Option<RuleId> {
        self.pratt.get(rule)?.as_ref()?.rhs_atom
    }
    fn source_pratt_operator_trivia(&self,rule:RuleId)->Option<RuleId> {
        self.pratt.get(rule)?.as_ref()?.operator_trivia
    }
    fn source_pratt(&self, rule: RuleId) -> Option<SourcePratt<'_>> {
        self.pratt.get(rule)?.as_ref().map(|spec| SourcePratt {
            atom: spec.atom,
            operators: &spec.operators,
            trailing_trivia: spec.trailing_trivia,
        })
    }
    fn scope(&self, id: usize) -> Scope {
        self.scopes[id]
    }
    fn whitespace_prefix_class(&self) -> Option<crate::lexical::ByteClass> {
        self.scopes[0].whitespace_prefix_class
    }
    fn whitespace_class(&self) -> Option<crate::lexical::ByteClass> {
        self.scopes[0].whitespace_class
    }
    fn rule_id(&self, name: &str) -> Option<RuleId> {
        self.rule_id(name)
    }
    fn rule(&self, rule: RuleId) -> RuleSpec<'_> {
        let scope = self.rule_scopes[rule];
        let rule = &self.syntax().rules[rule];
        RuleSpec {
            name: &rule.name,
            mode: rule.mode,
            expression: rule.expression,
            scope,
        }
    }
    fn instruction(&self, id: ExprId) -> Instruction<'_> {
        match &self.syntax().expressions[id].kind {
            ExprKind::Literal { text, insensitive } => Instruction::Literal {
                text,
                insensitive: *insensitive,
            },
            ExprKind::Range { start, end } => Instruction::Range {
                start: *start,
                end: *end,
            },
            ExprKind::Reference(_) => match self.reference(id).unwrap() {
                Reference::Rule(rule) => Instruction::Call(rule),
                Reference::Builtin(builtin) => Instruction::Builtin(builtin),
            },
            ExprKind::Sequence { left, right } => Instruction::Sequence {
                left: *left,
                right: *right,
            },
            ExprKind::Choice { left, right } => Instruction::Choice {
                left: *left,
                right: *right,
            },
            ExprKind::Repeat {
                expression,
                min,
                max,
            } => Instruction::Repeat {
                expression: *expression,
                min: *min,
                max: *max,
            },
            ExprKind::Predicate {
                expression,
                positive,
            } => Instruction::Predicate {
                expression: *expression,
                positive: *positive,
            },
            ExprKind::Group(child) => Instruction::Group(*child),
            ExprKind::Push(child) => Instruction::Push(*child),
            ExprKind::PushLiteral(text) => Instruction::PushLiteral(text),
            ExprKind::PeekSlice { start, end } => Instruction::PeekSlice {
                start: *start,
                end: *end,
            },
        }
    }
}

/// Conservative mandatory ASCII head: a byte class or complete literal word.
/// The boolean on each word selects ASCII case-insensitive matching.
#[derive(Debug)]
pub struct ChoiceHead {pub class:[u64;4],pub words:&'static[(&'static str,bool)]}
/// Original ordered-choice branch. Unknown heads must always be admitted.
#[derive(Debug)]
pub struct ChoiceHeadBranch {pub expression:usize,pub head:Option<&'static ChoiceHead>}
/// Immutable generated proof for a selected choice in its original lexical scope.
/// `branches` preserves original order and has at most 64 entries. Every mask
/// in `first` contains only valid branch bits; unknown heads are present in all
/// masks. `rules` covers all named rules inspected by known head proofs, and
/// `depth` bounds their maximum additional rule nesting before head failure.
/// Custom programs granting `choice_heads_metadata_is_current` must uphold
/// these invariants as well as agreement with their current program geometry.
#[derive(Debug)]
pub struct ChoiceHeads {pub first:&'static[u64;256],pub scope:usize,pub expression:usize,pub branches:&'static[ChoiceHeadBranch],pub rules:&'static[usize],pub depth:usize}

#[derive(Debug,Clone,Copy)]
pub struct EmptySuffixHead {pub owner:usize,pub owner_body:usize,pub repetition:usize,pub child:usize,pub scope:usize,pub first:[u64;4],pub calls:[usize;8],pub calls_len:u8,pub depth:usize}
