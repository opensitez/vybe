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

pub trait Program: std::fmt::Debug {
    fn atomic_fragments(&self) -> Option<()> {
        None
    }
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
}

impl Program for CompiledGrammar {
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
