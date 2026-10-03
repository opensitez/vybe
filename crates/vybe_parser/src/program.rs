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

/// Immutable parser program. Generated implementations provide static Rust
/// dispatch/literals and need no grammar parsing, allocation, or name resolution
/// at source-parse startup. Implementations must provide valid arena IDs.
pub trait Program: std::fmt::Debug {
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
