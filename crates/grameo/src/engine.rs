//! Iterative scannerless PEG reference engine. No host/Rust recursion is used
//! for source nesting. Generated recognition/direct AST backends follow later.
use crate::builder::{Builder, NoBuilder};
use crate::grammar::{ExprId, RuleId};
use crate::program::{Instruction, Program};
use crate::tree::{CaptureRule, Node, Nodes, ParseTree};
use crate::{Builtin, CompiledGrammar, Reference, RuleMode, Span};
use std::fmt;

#[derive(Debug, Clone, Copy)]
pub struct MatchOptions {
    /// Maximum parser operations. `usize::MAX` means no work cutoff.
    pub max_steps: usize,
    pub max_rule_depth: usize,
}
impl Default for MatchOptions {
    fn default() -> Self {
        Self {
            max_steps: usize::MAX,
            max_rule_depth: 4096,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParseErrorKind {
    Syntax,
    UnknownEntry,
    WorkLimit,
    DepthLimit,
    NonProgress,
    InvalidStack,
    BuildFailure,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParseError {
    pub kind: ParseErrorKind,
    pub offset: usize,
    pub expected: Vec<String>,
    pub message: String,
}
impl fmt::Display for ParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "{:?} at byte {}: {}",
            self.kind, self.offset, self.message
        )?;
        if !self.expected.is_empty() {
            write!(f, " (expected {})", self.expected.join(", "))?;
        }
        Ok(())
    }
}
impl std::error::Error for ParseError {}
impl From<crate::builder::BuildError> for ParseError {
    fn from(error: crate::builder::BuildError) -> Self {
        Self {
            kind: ParseErrorKind::BuildFailure,
            offset: error.span.start,
            expected: Vec::new(),
            message: error.message,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Recognition {
    pub consumed: usize,
    pub steps: usize,
}

pub fn parse_program<'g, 'i, P: Program>(
    program: &'g P,
    entry: &str,
    source: &'i str,
    options: MatchOptions,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    program.capture_parse(entry, source, options)
}
/// Parse one rule at an absolute UTF-8 boundary in the original input.
/// Captures retain the full input and absolute spans; `consumed` is the
/// absolute endpoint. SOI/EOI still refer to the full input boundaries.
/// This capture-only entry is useful for deferred contextual replay. It does
/// not change ordinary PEG parsing or enable semantic hooks/Pratt islands.
pub fn parse_program_at<'g, 'i, P: Program>(
    program: &'g P,
    entry: &str,
    source: &'i str,
    start_offset: usize,
    options: MatchOptions,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if start_offset > source.len() || !source.is_char_boundary(start_offset) {
        return Err(ParseError {
            kind: ParseErrorKind::Syntax,
            offset: start_offset,
            expected: Vec::new(),
            message: "parse start must be an in-bounds UTF-8 boundary".into(),
        });
    }
    if program.atomic_fragments().is_some() {
        parse_captures_at::<P, Env, true>(program, entry, source, start_offset, options)
    } else {
        parse_captures_at::<P, Env, false>(program, entry, source, start_offset, options)
    }
}
#[doc(hidden)]
pub fn parse_wide_program<'g, 'i, P: Program>(
    program: &'g P,
    entry: &str,
    source: &'i str,
    options: MatchOptions,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if program.atomic_fragments().is_some() {
        parse_captures::<P, Env, true>(program, entry, source, options)
    } else {
        parse_captures::<P, Env, false>(program, entry, source, options)
    }
}

#[doc(hidden)]
pub fn parse_compact_program<'g, 'i, P: Program>(
    program: &'g P,
    entry: &str,
    source: &'i str,
    options: MatchOptions,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if program.atomic_fragments().is_some() {
        parse_captures::<P, CompactEnv, true>(program, entry, source, options)
    } else {
        parse_captures::<P, CompactEnv, false>(program, entry, source, options)
    }
}

fn parse_captures<'g, 'i, P: Program, E: Environment, const FRAGMENT: bool>(
    program: &'g P,
    entry: &str,
    source: &'i str,
    options: MatchOptions,
) -> Result<ParseTree<'g, 'i>, ParseError>
where
    TaskMode<true>: PrattTaskMode<E>,
{
    parse_captures_at::<P, E, FRAGMENT>(program, entry, source, 0, options)
}

fn parse_captures_at<'g, 'i, P: Program, E: Environment, const FRAGMENT: bool>(
    program: &'g P,
    entry: &str,
    source: &'i str,
    start_offset: usize,
    options: MatchOptions,
) -> Result<ParseTree<'g, 'i>, ParseError>
where
    TaskMode<true>: PrattTaskMode<E>,
{
    // Successful capture parses expose no step count or failure expectations.
    // Keep explicit work budgets and all errors on the instrumented backend.
    if options.max_steps == usize::MAX {
        if let Ok(state) = State::<P, NoBuilder, true, E, FRAGMENT, false, false>::run_at(
            program, source, entry, start_offset, options, NoBuilder,
        ) {
            return Ok(ParseTree { grammar: program, source, nodes: state.nodes, consumed: state.offset });
        }
    }
    let state =
        State::<P, NoBuilder, true, E, FRAGMENT>::run_at(program, source, entry, start_offset, options, NoBuilder)?;
    Ok(ParseTree {
        grammar: program,
        source,
        nodes: state.nodes,
        consumed: state.offset,
    })
}

/// Capture entry for generated programs whose entire instruction arena is stack-free.
///
/// The caller must ensure that no instruction performs a stack push, read, or pop.
/// Generated callers select this entry only after the code generator checks that
/// whole-arena precondition. Programs requiring stack behavior must use the
/// ordinary compact or wide entry instead.
#[doc(hidden)]
pub fn parse_stackfree_compact_program<'g, 'i, P: Program>(
    program: &'g P,
    entry: &str,
    source: &'i str,
    options: MatchOptions,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if program.atomic_fragments().is_some() {
        parse_captures::<P, StackFreeEnv, true>(program, entry, source, options)
    } else {
        parse_captures::<P, StackFreeEnv, false>(program, entry, source, options)
    }
}

/// Build semantic values while retaining captures for rules that still need a
/// contextual walker. Configured Pratt islands use the builder's opt-in hooks;
/// their captures contain atoms/operators instead of the original PEG wrappers.
/// Ordinary `parse_program` capture layout is unchanged. Any failure restores
/// both the capture arena and the builder's initial transactional checkpoint.
pub fn build_capture_program<'g, 'i, P: Program, B: Builder>(
    program: &'g P, entry: &str, source: &'i str, options: MatchOptions, builder: &mut B,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if program.atomic_fragments().is_some() {
        hybrid_captures::<P, B, Env, true>(program, entry, source, options, builder)
    } else {
        hybrid_captures::<P, B, Env, false>(program, entry, source, options, builder)
    }
}

/// Build one explicitly selected semantic unit at an absolute UTF-8 boundary.
/// Hooks and opted-in Pratt islands run against the original full source, so
/// captures, diagnostics and semantic spans retain absolute offsets. Failure
/// restores only this call's builder checkpoint; earlier completed units remain.
/// Success calls `Builder::committed` once for this entry rule. The frontend must
/// choose a stable unit boundary and establish its declaration/scope barrier
/// before publishing or emitting that unit. This API does not infer boundaries
/// or make speculative rule finishes safe for external side effects.
pub fn build_capture_program_at<'g, 'i, P: Program, B: Builder>(
    program: &'g P, entry: &str, source: &'i str, start_offset: usize,
    options: MatchOptions, builder: &mut B,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if start_offset > source.len() || !source.is_char_boundary(start_offset) {
        return Err(ParseError {
            kind: ParseErrorKind::Syntax, offset: start_offset, expected: Vec::new(),
            message: "parse start must be an in-bounds UTF-8 boundary".into(),
        });
    }
    if program.atomic_fragments().is_some() {
        hybrid_captures_at::<P, B, Env, true>(program, entry, source, start_offset, options, builder)
    } else {
        hybrid_captures_at::<P, B, Env, false>(program, entry, source, start_offset, options, builder)
    }
}

/// Absolute-offset semantic capture entry with certified compact scopes.
#[doc(hidden)]
pub fn build_compact_capture_program_at<'g, 'i, P: Program, B: Builder>(
    program: &'g P, entry: &str, source: &'i str, start_offset: usize,
    options: MatchOptions, builder: &mut B,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if start_offset > source.len() || !source.is_char_boundary(start_offset) {
        return Err(ParseError {
            kind: ParseErrorKind::Syntax, offset: start_offset, expected: Vec::new(),
            message: "parse start must be an in-bounds UTF-8 boundary".into(),
        });
    }
    if program.atomic_fragments().is_some() {
        hybrid_captures_at::<P, B, CompactEnv, true>(program, entry, source, start_offset, options, builder)
    } else {
        hybrid_captures_at::<P, B, CompactEnv, false>(program, entry, source, start_offset, options, builder)
    }
}

/// Absolute-offset semantic capture entry for certified generated programs.
///
/// The entire instruction arena must be stack-free and all capture scopes must
/// satisfy the same compact-scope contract as the zero-offset stack-free entry.
/// Callers without those certificates must use `build_capture_program_at`.
/// UTF-8 boundaries, absolute captures/diagnostics, budgets and transactional
/// builder completion have the same contract as that general entry.
#[doc(hidden)]
pub fn build_stackfree_compact_capture_program_at<'g, 'i, P: Program, B: Builder>(
    program: &'g P, entry: &str, source: &'i str, start_offset: usize,
    options: MatchOptions, builder: &mut B,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if start_offset > source.len() || !source.is_char_boundary(start_offset) {
        return Err(ParseError {
            kind: ParseErrorKind::Syntax, offset: start_offset, expected: Vec::new(),
            message: "parse start must be an in-bounds UTF-8 boundary".into(),
        });
    }
    if program.atomic_fragments().is_some() {
        hybrid_captures_at::<P, B, StackFreeEnv, true>(program, entry, source, start_offset, options, builder)
    } else {
        hybrid_captures_at::<P, B, StackFreeEnv, false>(program, entry, source, start_offset, options, builder)
    }
}

fn hybrid_captures_at<'g, 'i, P: Program, B: Builder, E: Environment, const FRAGMENT: bool>(
    program: &'g P, entry: &str, source: &'i str, start_offset: usize,
    options: MatchOptions, builder: &mut B,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    let mark = builder.checkpoint();
    if options.max_steps == usize::MAX {
        match State::<P, &mut B, true, HybridEnv<E>, FRAGMENT, false, false, true>::run_at(
            program, source, entry, start_offset, options, &mut *builder,
        ) {
            Ok(state) => return Ok(ParseTree { grammar: program, source, nodes: state.nodes, consumed: state.offset }),
            Err(error) => {
                builder.rollback(mark);
                if error.kind != ParseErrorKind::Syntax { return Err(error); }
            }
        }
    }
    let result = State::<P, &mut B, true, HybridEnv<E>, FRAGMENT, true, true, true>::run_at(
        program, source, entry, start_offset, options, &mut *builder,
    ).map(|state| ParseTree { grammar: program, source, nodes: state.nodes, consumed: state.offset });
    if result.is_err() { builder.rollback(mark); }
    result
}

/// Semantic capture entry for generated programs with certified compact scopes.
#[doc(hidden)]
pub fn build_compact_capture_program<'g, 'i, P: Program, B: Builder>(
    program: &'g P, entry: &str, source: &'i str, options: MatchOptions, builder: &mut B,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if program.atomic_fragments().is_some() {
        hybrid_captures::<P, B, CompactEnv, true>(program, entry, source, options, builder)
    } else { hybrid_captures::<P, B, CompactEnv, false>(program, entry, source, options, builder) }
}

/// Semantic capture entry for generated programs whose instruction arena is stack-free.
#[doc(hidden)]
pub fn build_stackfree_compact_capture_program<'g, 'i, P: Program, B: Builder>(
    program: &'g P, entry: &str, source: &'i str, options: MatchOptions, builder: &mut B,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    if program.atomic_fragments().is_some() {
        hybrid_captures::<P, B, StackFreeEnv, true>(program, entry, source, options, builder)
    } else { hybrid_captures::<P, B, StackFreeEnv, false>(program, entry, source, options, builder) }
}

fn hybrid_captures<'g, 'i, P: Program, B: Builder, E: Environment, const FRAGMENT: bool>(
    program: &'g P, entry: &str, source: &'i str, options: MatchOptions, builder: &mut B,
) -> Result<ParseTree<'g, 'i>, ParseError> {
    let mark = builder.checkpoint();
    if options.max_steps == usize::MAX {
        match State::<P, &mut B, true, HybridEnv<E>, FRAGMENT, false, false, true>::run(
            program, source, entry, options, &mut *builder,
        ) {
            Ok(state) => return Ok(ParseTree { grammar: program, source, nodes: state.nodes, consumed: state.offset }),
            Err(error) => {
                builder.rollback(mark);
                if error.kind != ParseErrorKind::Syntax { return Err(error); }
            }
        }
    }
    let result = State::<P, &mut B, true, HybridEnv<E>, FRAGMENT, true, true, true>::run(
        program, source, entry, options, &mut *builder,
    ).map(|state| ParseTree { grammar: program, source, nodes: state.nodes, consumed: state.offset });
    if result.is_err() { builder.rollback(mark); }
    result
}

pub fn recognize_program<P: Program>(
    program: &P,
    entry: &str,
    source: &str,
    options: MatchOptions,
) -> Result<Recognition, ParseError> {
    let state = State::<P, NoBuilder, false>::run(program, source, entry, options, NoBuilder)?;
    Ok(Recognition {
        consumed: state.offset,
        steps: state.steps,
    })
}

/// Match directly into a transactional semantic builder. No capture nodes are
/// allocated. Syntax/resource errors restore the builder's initial state.
pub fn build_program<P: Program, B: Builder>(
    program: &P,
    entry: &str,
    source: &str,
    options: MatchOptions,
    builder: &mut B,
) -> Result<Recognition, ParseError> {
    let mark = builder.checkpoint();
    let result =
        State::<P, &mut B, false>::run(program, source, entry, options, builder).map(|state| {
            Recognition {
                consumed: state.offset,
                steps: state.steps,
            }
        });
    if result.is_err() {
        builder.rollback(mark);
    }
    result
}

/// Construct an AST without capture storage or operation-count instrumentation.
/// The caller owns the semantic result in `builder`; this returns consumed bytes.
/// Explicit work budgets use the ordinary builder engine. Syntax failures
/// restore the builder and rerun with complete expectations; binding errors
/// and structural/resource errors already have precise diagnostics.
pub fn build_ast_program<P: Program, B: Builder>(
    program: &P,
    entry: &str,
    source: &str,
    options: MatchOptions,
    builder: &mut B,
) -> Result<usize, ParseError> {
    if options.max_steps != usize::MAX {
        return build_program(program, entry, source, options, builder).map(|r| r.consumed);
    }
    let mark = builder.checkpoint();
    let result = if program.atomic_fragments().is_some() {
        State::<P, &mut B, false, Env, true, false, false>::run(
            program, source, entry, options, builder,
        ).map(|state| state.offset)
    } else {
        State::<P, &mut B, false, Env, false, false, false>::run(
            program, source, entry, options, builder,
        ).map(|state| state.offset)
    };
    match result {
        Ok(consumed) => Ok(consumed),
        Err(error) => {
            builder.rollback(mark);
            if error.kind == ParseErrorKind::Syntax {
                build_program(program, entry, source, options, builder).map(|r| r.consumed)
            } else {
                Err(error)
            }
        }
    }
}

impl CompiledGrammar {
    /// Build a caller-selected semantic unit while preserving full-source spans.
    /// Declaration visibility and safe publication remain frontend decisions.
    pub fn build_captures_at<'g, 'i, B: Builder>(
        &'g self, entry: &str, source: &'i str, start_offset: usize, builder: &mut B,
    ) -> Result<ParseTree<'g, 'i>, ParseError> {
        build_capture_program_at(self, entry, source, start_offset, MatchOptions::default(), builder)
    }

    /// Construct semantic values and retain contextual captures in one parse.
    /// Only builder-opted-in Pratt islands replace the ordinary PEG backend.
    pub fn build_captures<'g, 'i, B: Builder>(
        &'g self, entry: &str, source: &'i str, builder: &mut B,
    ) -> Result<ParseTree<'g, 'i>, ParseError> {
        build_capture_program(self, entry, source, MatchOptions::default(), builder)
    }

    pub fn build_ast<B: Builder>(
        &self, entry: &str, source: &str, builder: &mut B,
    ) -> Result<usize, ParseError> {
        build_ast_program(self, entry, source, MatchOptions::default(), builder)
    }

    pub fn parse_source<'g, 'i>(
        &'g self,
        entry: &str,
        source: &'i str,
    ) -> Result<ParseTree<'g, 'i>, ParseError> {
        self.parse_source_with_options(entry, source, MatchOptions::default())
    }
    pub fn parse_source_with_options<'g, 'i>(
        &'g self,
        entry: &str,
        source: &'i str,
        options: MatchOptions,
    ) -> Result<ParseTree<'g, 'i>, ParseError> {
        parse_program(self, entry, source, options)
    }
    pub fn recognize(&self, entry: &str, source: &str) -> Result<Recognition, ParseError> {
        self.recognize_with_options(entry, source, MatchOptions::default())
    }
    pub fn recognize_with_options(
        &self,
        entry: &str,
        source: &str,
        options: MatchOptions,
    ) -> Result<Recognition, ParseError> {
        let state = State::<Self, NoBuilder, false>::run(self, source, entry, options, NoBuilder)?;
        Ok(Recognition {
            consumed: state.offset,
            steps: state.steps,
        })
    }
}

#[inline(always)]
fn failure_prefix<'p, G: Program + ?Sized>(
    grammar: &'p G,
    rule: RuleId,
    owned: &'p mut Option<crate::lexical::FailurePrefix>,
) -> Option<&'p crate::lexical::FailurePrefix> {
    if grammar.borrowed_failure_prefixes() {
        grammar.trivia_failure_prefix_ref(rule)
    } else {
        *owned = grammar.trivia_failure_prefix(rule);
        owned.as_ref()
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Atomicity {
    Normal,
    Atomic,
    Compound,
}
#[derive(Clone, Copy)]
struct Env {
    scope: usize,
    flags: u8,
}
#[derive(Clone, Copy)]
struct CompactEnv(usize);
trait Environment: Copy {
    type UndoMark: Copy;
    // Trivia and predicates cannot mutate the builder. Their continuations
    // retain only the grammar stack cursor, never a semantic cursor.
    type PassiveUndoMark: Copy;
    fn passive_undo_mark(len: usize) -> Self::PassiveUndoMark;
    fn passive_undo_value(mark: Self::PassiveUndoMark) -> usize;
    fn undo_mark(len: usize) -> Self::UndoMark;
    fn undo_value(mark: Self::UndoMark) -> usize;
    fn semantic_mark(mark: Self::UndoMark, _: usize) -> Self::UndoMark { mark }
    fn semantic_value(_: Self::UndoMark) -> usize { 0 }
    type CaptureCompletion;
    fn capture_completion(value: Option<usize>) -> Self::CaptureCompletion;
    fn capture_completion_value(value: Self::CaptureCompletion) -> Option<usize>;
    fn root() -> Self;
    fn scope(self) -> usize;
    fn flags(self) -> u8;
    fn with_flags(self, flags: u8) -> Self;
    fn enter<P: Program>(program: &P, rule: RuleId, scope: usize, flags: u8) -> Self;
    fn mode(self) -> Atomicity {
        match self.flags() & 3 {
            0 => Atomicity::Normal,
            1 => Atomicity::Atomic,
            _ => Atomicity::Compound,
        }
    }
    fn predicate(self) -> bool {
        self.flags() & 4 != 0
    }
    fn skipping(self) -> bool {
        self.flags() & 8 != 0
    }
    fn semantic(self) -> bool {
        self.flags() & 32 != 0
    }
    fn captures_suppressed(self) -> bool { self.flags() & 16 != 0 }
    fn with_captures_suppressed(self, value: bool) -> Self {
        self.with_flags((self.flags() & !16) | (u8::from(value) << 4))
    }
    fn with_predicate(self, value: bool) -> Self {
        self.with_flags((self.flags() & !4) | (u8::from(value) << 2))
    }
    fn without_semantic(self) -> Self {
        self.with_flags(self.flags() & !32)
    }
    fn with_skipping(self) -> Self {
        self.with_flags(self.flags() | 8)
    }
    fn profile_flags(self, capture: bool) -> u8 {
        self.flags() & (3 | 4 | 8) | (u8::from(capture) << 4)
    }
    fn allows_skip(self) -> bool {
        self.flags() & (3 | 8) == 0
    }
}
impl Environment for Env {
    type PassiveUndoMark = usize;
    fn passive_undo_mark(len: usize) -> usize { len }
    fn passive_undo_value(mark: usize) -> usize { mark }
    type UndoMark = usize;
    fn undo_mark(len: usize) -> usize {
        len
    }
    fn undo_value(mark: usize) -> usize {
        mark
    }
    type CaptureCompletion = Option<usize>;
    fn capture_completion(value: Option<usize>) -> Self::CaptureCompletion {
        value
    }
    fn capture_completion_value(value: Self::CaptureCompletion) -> Option<usize> {
        value
    }
    fn root() -> Self {
        Self {
            scope: 0,
            flags: 32,
        }
    }
    fn scope(self) -> usize {
        self.scope
    }
    fn flags(self) -> u8 {
        self.flags
    }
    fn with_flags(self, flags: u8) -> Self {
        Self { flags, ..self }
    }
    fn enter<P: Program>(_: &P, _: RuleId, scope: usize, flags: u8) -> Self {
        Self { scope, flags }
    }
}
impl Environment for CompactEnv {
    type PassiveUndoMark = usize;
    fn passive_undo_mark(len: usize) -> usize { len }
    fn passive_undo_value(mark: usize) -> usize { mark }
    type UndoMark = usize;
    fn undo_mark(len: usize) -> usize {
        len
    }
    fn undo_value(mark: usize) -> usize {
        mark
    }
    type CaptureCompletion = Option<std::num::NonZeroUsize>;
    fn capture_completion(value: Option<usize>) -> Self::CaptureCompletion {
        value.map(|v| {
            std::num::NonZeroUsize::new(
                v.checked_add(1)
                    .expect("capture index cannot exceed addressable nodes"),
            )
            .unwrap()
        })
    }
    fn capture_completion_value(value: Self::CaptureCompletion) -> Option<usize> {
        value.map(|v| v.get() - 1)
    }
    fn root() -> Self {
        Self(32)
    }
    fn scope(self) -> usize {
        self.0 >> 6
    }
    #[inline(always)]
    fn flags(self) -> u8 {
        (self.0 & 63) as u8
    }
    fn with_flags(self, flags: u8) -> Self {
        Self((self.0 & !63) | (flags as usize))
    }
    fn enter<P: Program>(program: &P, _: RuleId, scope: usize, flags: u8) -> Self {
        Self(program.compact_capture_scope(scope).word() | flags as usize)
    }
}

#[derive(Clone, Copy)]
struct StackFreeEnv(CompactEnv);
impl Environment for StackFreeEnv {
    type PassiveUndoMark = ();
    fn passive_undo_mark(_: usize) {}
    fn passive_undo_value(_: ()) -> usize { 0 }
    type UndoMark = ();
    fn undo_mark(_: usize) {}
    fn undo_value(_: ()) -> usize {
        0
    }
    type CaptureCompletion = <CompactEnv as Environment>::CaptureCompletion;
    fn capture_completion(value: Option<usize>) -> Self::CaptureCompletion {
        CompactEnv::capture_completion(value)
    }
    fn capture_completion_value(value: Self::CaptureCompletion) -> Option<usize> {
        CompactEnv::capture_completion_value(value)
    }
    fn root() -> Self {
        Self(CompactEnv::root())
    }
    fn scope(self) -> usize {
        self.0.scope()
    }
    #[inline(always)]
    fn flags(self) -> u8 {
        self.0.flags()
    }
    fn with_flags(self, flags: u8) -> Self {
        Self(self.0.with_flags(flags))
    }
    fn enter<P: Program>(program: &P, rule: RuleId, scope: usize, flags: u8) -> Self {
        Self(CompactEnv::enter(program, rule, scope, flags))
    }
}

#[derive(Clone, Copy)]
struct HybridEnv<E: Environment = Env>(E);
impl<E: Environment> Environment for HybridEnv<E> {
    type PassiveUndoMark = E::PassiveUndoMark;
    fn passive_undo_mark(len: usize) -> Self::PassiveUndoMark { E::passive_undo_mark(len) }
    fn passive_undo_value(mark: Self::PassiveUndoMark) -> usize { E::passive_undo_value(mark) }
    type UndoMark = (E::UndoMark, usize);
    fn undo_mark(len: usize) -> Self::UndoMark { (E::undo_mark(len), 0) }
    fn undo_value(mark: Self::UndoMark) -> usize { E::undo_value(mark.0) }
    fn semantic_mark(mark: Self::UndoMark, semantic: usize) -> Self::UndoMark { (mark.0, semantic) }
    fn semantic_value(mark: Self::UndoMark) -> usize { mark.1 }
    type CaptureCompletion = E::CaptureCompletion;
    fn capture_completion(value: Option<usize>) -> Self::CaptureCompletion { E::capture_completion(value) }
    fn capture_completion_value(value: Self::CaptureCompletion) -> Option<usize> { E::capture_completion_value(value) }
    fn root() -> Self { Self(E::root()) }
    fn scope(self) -> usize { self.0.scope() }
    #[inline(always)]
    fn flags(self) -> u8 { self.0.flags() }
    fn with_flags(self, flags: u8) -> Self { Self(self.0.with_flags(flags)) }
    fn enter<P: Program>(program: &P, rule: RuleId, scope: usize, flags: u8) -> Self {
        Self(E::enter(program, rule, scope, flags))
    }
}

#[derive(Clone, Copy)]
enum StackValue {
    Source(Span),
    Literal(ExprId),
}
enum Undo {
    RemovePush,
    RestorePop(StackValue),
}
#[derive(Clone, Copy)]
struct Checkpoint<E: Environment = Env> {
    offset: usize,
    undo: E::UndoMark,
    // Capture length or semantic cursor. Hybrid stores its second cursor in
    // the environment-specific undo mark, leaving ordinary task sizes unchanged.
    mark: usize,
}
impl<E: Environment> Checkpoint<E> {
    // Rust slices (including str) cannot exceed isize::MAX bytes. Only the
    // FinishRule checkpoint uses the spare high bit; all source positions are
    // read through offset(). Capture and builder cursors keep their full width.
    const CAPTURE_FLAG: usize = 1usize << (usize::BITS - 1);
    fn offset(self) -> usize { self.offset & !Self::CAPTURE_FLAG }
    fn captures_rule(self) -> bool { self.offset & Self::CAPTURE_FLAG != 0 }
    fn with_rule_capture(mut self, capture: bool) -> Self {
        debug_assert_eq!(self.offset & Self::CAPTURE_FLAG, 0);
        if capture { self.offset |= Self::CAPTURE_FLAG; }
        self
    }
}
#[derive(Clone, Copy)]
struct PassiveCheckpoint<E: Environment> {
    offset: usize,
    undo: E::PassiveUndoMark,
    mark: usize,
}
impl<E: Environment> PassiveCheckpoint<E> {
    fn offset(self) -> usize { self.offset }
}
#[derive(Clone, Copy)]
struct Repetition<E: Environment> {
    child: ExprId,
    env: E,
    min: u32,
    max: Option<u32>,
    count: u32,
}

struct RepeatFrame<E: Environment> {
    repetition: Repetition<E>,
    checkpoint: Checkpoint<E>,
}

#[derive(Clone, Copy)]
struct PendingOperator {
    operator: crate::program::BoundOperator,
    span: Span,
}
enum OperatorUndo {
    Remove,
    Restore(PendingOperator),
}
#[derive(Clone, Copy)]
enum PrattStep { Operand, AfterPrefix, AfterAtom, Following, AfterFollowing, AfterOperatorTrivia }
struct PrattFrame<E: Environment> {
    step: PrattStep,
    candidate: usize,
    target: RuleId,
    env: E,
    // The selected island may suppress its descendants. Fallback must restore
    // the enclosing capture policy, including nested already-suppressed islands.
    enclosing_captures_suppressed: bool,
    operators: Vec<PendingOperator>,
    undo: Vec<OperatorUndo>,
    attempt: Option<(Checkpoint<E>, usize)>,
    // Only an operator match needs these fields. Keep them in the existing
    // boxed frame instead of widening every PEG task in a hybrid parse.
    following: Option<(usize, Checkpoint<E>, usize)>,
}

// Ordinary capture mode has no Pratt continuations. Hybrid semantic capture
// opts into boxed frames without expanding the ordinary capture task layout.
struct TaskMode<const CAPTURE: bool, const HYBRID: bool = false>;
trait PrattTaskMode<E: Environment> {
    type Completion;
    fn completion(value: Option<usize>, semantic: Option<usize>) -> Self::Completion;
    fn observed_completion(_: &Self::Completion) -> bool { false }
    fn completion_value(value: Self::Completion, checkpoint: Checkpoint<E>) -> Option<usize>;
    type Frame;
    fn wrap(frame: Box<PrattFrame<E>>) -> Self::Frame;
    fn unwrap(frame: Self::Frame) -> Box<PrattFrame<E>>;
}
impl<E: Environment> PrattTaskMode<E> for TaskMode<true> {
    type Completion = E::CaptureCompletion;
    fn completion(value: Option<usize>, _: Option<usize>) -> Self::Completion {
        E::capture_completion(value)
    }
    fn completion_value(value: Self::Completion, _: Checkpoint<E>) -> Option<usize> {
        E::capture_completion_value(value)
    }

    type Frame = std::convert::Infallible;
    fn wrap(_frame: Box<PrattFrame<E>>) -> Self::Frame {
        unreachable!("capture mode cannot create Pratt continuations")
    }
    fn unwrap(frame: Self::Frame) -> Box<PrattFrame<E>> {
        match frame {}
    }
}
impl<E: Environment> PrattTaskMode<E> for TaskMode<false> {
    type Completion = Option<usize>;
    fn completion(value: Option<usize>, _: Option<usize>) -> Self::Completion {
        value
    }
    fn completion_value(value: Self::Completion, _: Checkpoint<E>) -> Option<usize> {
        value
    }

    type Frame = Box<PrattFrame<E>>;
    fn wrap(frame: Box<PrattFrame<E>>) -> Self::Frame {
        frame
    }
    fn unwrap(frame: Self::Frame) -> Box<PrattFrame<E>> {
        frame
    }
}

impl<E: Environment> PrattTaskMode<E> for TaskMode<true, true> {
    // Full-width observed IDs live in State's nested rule stack. A task only
    // records whether its successful begin contributed a pending hook.
    type Completion = bool;
    type Frame = Box<PrattFrame<E>>;
    fn completion(_: Option<usize>, semantic: Option<usize>) -> Self::Completion {
        semantic.is_some()
    }
    fn completion_value(_: Self::Completion, checkpoint: Checkpoint<E>) -> Option<usize> {
        checkpoint.captures_rule().then_some(checkpoint.mark)
    }
    fn observed_completion(value: &Self::Completion) -> bool { *value }
    fn wrap(frame: Self::Frame) -> Self::Frame { frame }
    fn unwrap(frame: Self::Frame) -> Self::Frame { frame }
}

struct SharedPrefixFrame<E: Environment> {
    spec: crate::program::SharedPrefixChoice,
    env: E,
    step: u8,
    start: Checkpoint<E>,
    prefix: Option<Checkpoint<E>>,
    capture: bool,
    generation: usize,
}
struct SharedInnerFrame<E:Environment>{spec:crate::program::SharedInnerPrefix,env:E,start:Checkpoint<E>,prefix:Option<Checkpoint<E>>,step:u8}
struct SharedGroupFrame<E:Environment>{spec:crate::program::SharedGroupPrefix,env:E,start:Checkpoint<E>,wrapper:Option<Checkpoint<E>>,prefix:Option<Checkpoint<E>>,step:u8,capture:bool}
struct SharedWrapperFrame<E:Environment>{spec:crate::program::SharedWrapperPrefix,env:E,start:Checkpoint<E>,prefix:Option<Checkpoint<E>>,step:u8,capture:bool}
struct ChoiceHeadsFrame<E:Environment>{branches:&'static[crate::program::ChoiceHeadBranch],remaining:u64,env:E}
struct OwnedCaptureFrame<E:Environment,C>{checkpoint:Checkpoint<E>,body_checkpoint:Checkpoint<E>,completion:C,expression:ExprId,env:E,end:usize}
enum Task<F, E: Environment, C> {SharedInner(Box<SharedInnerFrame<E>>),OwnedCapture(Box<OwnedCaptureFrame<E,C>>),SharedGroup(Box<SharedGroupFrame<E>>),SharedWrapper(usize),
    ChoiceHeads(Box<ChoiceHeadsFrame<E>>),
    SharedPrefix(Box<SharedPrefixFrame<E>>),
    Pratt(F),
    RunBuiltin(Builtin, E),
    Eval(ExprId, E),
    EvalCoveredSequence(ExprId, E),
    FinishExpr(Checkpoint<E>),
    EnterRule(RuleId, E),
    FinishRule {
        checkpoint: Checkpoint<E>,
        completion: C,
    },
    Sequence {
        right: ExprId,
        env: E,
    },
    AfterSkip {
        right: ExprId,
        env: E,
    },
    Choice {
        right: ExprId,
        env: E,
    },
    Repeat(usize),
    RepeatAfterSkip(usize),
    RepeatAfterChild(usize),
    Predicate {
        checkpoint: PassiveCheckpoint<E>,
        positive: bool,
    },
    Push {
        start: usize,
    },
    Skip(E),
    AfterWhitespace {
        checkpoint: PassiveCheckpoint<E>,
        env: E,
    },
    AfterComment {
        checkpoint: PassiveCheckpoint<E>,
        env: E,
    },
}

type ExecutionTask<E, const CAPTURE: bool, const HYBRID: bool = false> = Task<
    <TaskMode<CAPTURE, HYBRID> as PrattTaskMode<E>>::Frame,
    E,
    <TaskMode<CAPTURE, HYBRID> as PrattTaskMode<E>>::Completion,
>;

// Token rules need only cursor checkpoints and predicate state. Reuse this
// compact scratch stack across rules instead of full capture/builder frames.
#[derive(Clone, Copy)]
struct AtomicRepeat {
    child: ExprId,
    predicate: bool,
    min: u32,
    max: u32,
    bounded: bool,
    count: u32,
}
enum FragmentOutcome<'g> {
    Matched(usize, usize, usize, usize),
    Rejected(usize, &'g [crate::program::FragmentExpected<'g>]),
}
enum AtomicTask {
    Eval(ExprId, bool),
    Restore(usize),
    Sequence(ExprId, bool),
    Choice(ExprId, bool),
    Repeat(AtomicRepeat),
    AfterSkip(AtomicRepeat, usize),
    AfterChild(AtomicRepeat, usize),
    Predicate(usize, bool),
    EnterRule(RuleId, bool),
    FinishRule(usize),
}

// Failure descriptions remain borrowed/compact during speculation. Render only
// if the overall parse fails, rather than allocating strings for successful
// parses' discarded alternatives and terminating repetition attempts.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Expected<'g> {
    Literal(&'g str),
    Range(char, char),
    Builtin(Builtin),
    Named(&'static str),
}
impl Expected<'_> {
    fn render(self) -> String {
        match self {
            Self::Literal(text) => format!("{text:?}"),
            Self::Range(start, end) => format!("{start:?}..{end:?}"),
            Self::Builtin(builtin) => format!("{builtin:?}"),
            Self::Named(text) => text.into(),
        }
    }
}

struct State<
    'g,
    'i,
    P: Program,
    B: Builder,
    const CAPTURE: bool,
    E: Environment = Env,
    const FRAGMENT: bool = false,
    const DIAGNOSTICS: bool = true,
    const ACCOUNT: bool = true,
    const HYBRID: bool = false,
> where
    TaskMode<CAPTURE, HYBRID>: PrattTaskMode<E>,
{
    grammar: &'g P,
    builder: B,
    builder_mark: usize,
    capture_generation: usize,
    shared_wrapper_frames: Vec<SharedWrapperFrame<E>>,
    semantic_rules: Vec<RuleId>,
    source: &'i str,
    offset: usize,
    nodes: Nodes,
    stack: Vec<StackValue>,
    undo: Vec<Undo>,
    tasks: Vec<ExecutionTask<E, CAPTURE, HYBRID>>,
    repetitions: Vec<RepeatFrame<E>>,
    atomic_tasks: Vec<AtomicTask>,
    matched: bool,
    rule_depth: usize,
    steps: usize,
    options: MatchOptions,
    farthest: usize,
    expected: Vec<Expected<'g>>,
    expected_unique_len: usize,
}

impl<'g, 'i, P: Program, B: Builder, const CAPTURE: bool, E: Environment, const FRAGMENT: bool, const DIAGNOSTICS: bool, const ACCOUNT: bool, const HYBRID: bool>
    State<'g, 'i, P, B, CAPTURE, E, FRAGMENT, DIAGNOSTICS, ACCOUNT, HYBRID>
where
    TaskMode<CAPTURE, HYBRID>: PrattTaskMode<E>,
{
    fn run(
        grammar: &'g P,
        source: &'i str,
        entry: &str,
        options: MatchOptions,
        builder: B,
    ) -> Result<Self, ParseError> {
        Self::run_at(grammar, source, entry, 0, options, builder)
    }

    fn run_at(
        grammar: &'g P,
        source: &'i str,
        entry: &str,
        start_offset: usize,
        options: MatchOptions,
        builder: B,
    ) -> Result<Self, ParseError> {
        let entry = grammar
            .rule_id(entry)
            .map(Reference::Rule)
            .or_else(|| Builtin::from_name(entry).map(Reference::Builtin))
            .ok_or_else(|| ParseError {
                kind: ParseErrorKind::UnknownEntry,
                offset: start_offset,
                expected: Vec::new(),
                message: format!("unknown entry rule '{entry}'"),
            })?;
        let env = E::root();
        let mut tasks = Vec::with_capacity(32);
        tasks.push(match entry {
            Reference::Rule(rule) => Task::EnterRule(rule, env),
            Reference::Builtin(builtin) => Task::RunBuiltin(builtin, env),
        });
        let mut nodes = Nodes::new(source.len());
        if CAPTURE {
            let hint = ((source.len() - start_offset) / 4).clamp(4, 32).min(options.max_steps);
            if hint != 0 {
                nodes.reserve(1usize << (usize::BITS - 1 - hint.leading_zeros()));
            }
        }
        let builder_mark = if !CAPTURE || HYBRID { builder.checkpoint() } else { 0 };
        let mut state = Self {
            grammar,
            builder,
            builder_mark,
            capture_generation: 0,
            shared_wrapper_frames: Vec::new(),
            semantic_rules: Vec::new(),
            source,
            offset: start_offset,
            nodes,
            stack: Vec::new(),
            undo: Vec::new(),
            tasks,
            repetitions: Vec::new(),
            atomic_tasks: Vec::new(),
            matched: true,
            rule_depth: 0,
            steps: 0,
            options,
            farthest: start_offset,
            expected: Vec::new(),
            expected_unique_len: 0,
        };
        while let Some(task) = state.tasks.pop() {
            if ACCOUNT {
                if state.steps >= options.max_steps {
                    return Err(state.error(
                        ParseErrorKind::WorkLimit,
                        "source parse work limit exceeded",
                    ));
                }
                state.steps += 1;
            }
            state.execute(task)?;
        }
        if !state.matched {
            state.dedup_expected();
            return Err(ParseError {
                kind: ParseErrorKind::Syntax,
                offset: state.farthest,
                expected: state.expected.into_iter().map(Expected::render).collect(),
                message: "source does not match the grammar".into(),
            });
        }
        debug_assert!(state.semantic_rules.is_empty(), "every observed rule completes or falls back");
        state.builder.try_committed(source, state.offset)?;
        Ok(state)
    }
    fn checkpoint(&self) -> Checkpoint<E> {
        Checkpoint {
            offset: self.offset,
            undo: E::semantic_mark(E::undo_mark(self.undo.len()), if HYBRID { self.builder_mark } else { 0 }),
            mark: if CAPTURE {
                self.nodes.len()
            } else {
                self.builder_mark
            },
        }
    }
    fn passive_checkpoint(&self) -> PassiveCheckpoint<E> {
        PassiveCheckpoint {
            offset: self.offset,
            undo: E::passive_undo_mark(self.undo.len()),
            mark: self.nodes.len(),
        }
    }
    fn restore_passive(&mut self, checkpoint: PassiveCheckpoint<E>) {
        if CAPTURE { self.nodes.truncate(checkpoint.mark); }
        self.offset = checkpoint.offset;
        while self.undo.len() > E::passive_undo_value(checkpoint.undo) {
            match self.undo.pop().unwrap() {
                Undo::RemovePush => { self.stack.pop(); }
                Undo::RestorePop(value) => self.stack.push(value),
            }
        }
    }
    fn try_fast_trivia_failure(
        &mut self,
        env: E,
        mut whitespace_done: bool,
        continuation_steps: usize,
    ) -> bool {
        // Certified trivia failures perform no capture, stack or semantic
        // mutations. This also holds for hybrid captures: implicit skipping
        // suppresses builder hooks. Non-capture modes retain their old path.
        if !CAPTURE || !self.grammar.fast_trivia_failures() {
            return false;
        }
        let scope = self.grammar.scope(env.scope());
        if continuation_steps > 0
            && let Some(class) = scope.whitespace_prefix_class
        {
            // Skip checks depth even when its certified scanner consumes
            // zero bytes. Only bypass a zero-length scan, at the same depth.
            if self.rule_depth >= self.options.max_rule_depth
                || self
                    .source
                    .as_bytes()
                    .get(self.offset)
                    .is_some_and(|byte| class.contains(*byte))
            {
                return false;
            }
            whitespace_done = scope.whitespace_class.is_some();
        }
        let (cost, depth) = match self.grammar.trivia_failure_cost(
            env.scope(),
            whitespace_done,
            self.source.as_bytes().get(self.offset).copied(),
        ) {
            crate::program::TriviaFailure::Possible => return false,
            crate::program::TriviaFailure::Miss { steps, depth } => {
                let Some(cost) = steps.checked_add(continuation_steps) else {
                    return false;
                };
                (cost, depth)
            }
            crate::program::TriviaFailure::Unknown => {
                let mut cost: usize = continuation_steps + if whitespace_done { 1 } else { 2 };
                let mut depth = 0;
                for rule in scope
                    .whitespace
                    .filter(|_| !whitespace_done)
                    .into_iter()
                    .chain(scope.comment)
                {
                    let mut owned_prefix = None;
                    let Some(prefix) = failure_prefix(self.grammar, rule, &mut owned_prefix) else {
                        return false;
                    };
                    if self.source.as_bytes().get(self.offset).is_some_and(|byte| {
                        prefix.bytes[usize::from(*byte / 64)] & (1u64 << (*byte % 64)) != 0
                    }) {
                        return false;
                    }
                    let Some(next_cost) = cost
                        .checked_add(prefix.steps)
                        .and_then(|cost| cost.checked_add(2))
                    else {
                        return false;
                    };
                    let Some(rule_depth) = prefix.depth.checked_add(1) else {
                        return false;
                    };
                    cost = next_cost; // expression, EnterRule, FinishRule
                    depth = depth.max(rule_depth);
                }
                (cost, depth)
            }
        };
        if !self
            .rule_depth
            .checked_add(depth)
            .is_some_and(|depth| depth <= self.options.max_rule_depth)
            || !self
                .steps
                .checked_add(cost)
                .is_some_and(|steps| steps <= self.options.max_steps)
        {
            // Tight guards execute the original tasks to preserve which
            // limit fails first and its exact input offset.
            return false;
        }
        if ACCOUNT { self.steps += cost; }
        self.matched = true;
        true
    }
    fn try_choice_heads(&mut self,id:ExprId,env:E)->bool {
        if ACCOUNT || DIAGNOSTICS || !CAPTURE || !HYBRID || !env.semantic() || env.predicate() || env.skipping()
            || !self.builder.supports_choice_heads() || !self.grammar.match_observers_are_noop()
            || !self.grammar.choice_heads_metadata_is_current() || !self.grammar.choice_heads_candidate(id) {return false;}
        let Some(spec)=self.grammar.choice_heads(id) else{return false};
        if spec.expression!=id || spec.scope!=env.scope()
            || self.rule_depth.checked_add(spec.depth).is_none_or(|d|d>self.options.max_rule_depth)
            {return false;}
        let per_branch=self.grammar.choice_head_branch_metadata_is_current();
        if !per_branch && spec.rules.iter().any(|&r|self.builder.observes_rule(r)) {return false;}
        let Some(&byte)=self.source.as_bytes().get(self.offset) else{return false};
        if spec.branches.len()>64 {return false;}
        let mut mask=spec.first[usize::from(byte)];let mut refine=mask;
        while refine!=0{
            let i=refine.trailing_zeros() as usize;refine&=refine-1;
            let branch=&spec.branches[i];
            if branch.head.is_some_and(|head|head.class[usize::from(byte/64)]&(1u64<<(byte%64))==0
                && !head.words.iter().any(|(text,fold)|crate::lexical::literal_matches(&self.source[self.offset..],text,*fold))) {mask&=!(1u64<<i);}
        }
        // Do not front-load classification when the first PEG branch is still
        // possible. Common early-success forms keep their original short path.
        if mask&1!=0 {return false;}
        // A hook-bearing alternative stays in the PEG candidate set even if
        // its lexical head misses. Other inert alternatives can still be omitted.
        let mut omitted = if spec.branches.len()==64 { !mask } else { ((1u64<<spec.branches.len())-1)&!mask };
        while per_branch && omitted!=0 {
            let i=omitted.trailing_zeros() as usize;omitted&=omitted-1;
            let rules=self.grammar.choice_head_branch_rules(id,i).unwrap_or(spec.rules);
            if rules.iter().any(|&r|self.builder.observes_rule(r)) {mask|=1u64<<i;}
        }
        if mask.count_ones() as usize==spec.branches.len(){return false;}
        if mask==0{self.matched=false;return true;}
        let first=mask.trailing_zeros() as usize;mask&=mask-1;
        if mask!=0{self.tasks.push(Task::ChoiceHeads(Box::new(ChoiceHeadsFrame{branches:spec.branches,remaining:mask,env})));}
        self.tasks.push(Task::Eval(spec.branches[first].expression,env));true
    }
    fn try_native_choice_prefix(
        &mut self,
        prefix: &crate::program::NativeChoicePrefix<'_>,
        env: E,
    ) -> bool {
        // A failed lexical prefix still enters semantic rules in ordinary
        // matching. Only explicitly inert failure hooks may be omitted.
        if HYBRID && env.semantic() && !env.predicate() && !env.skipping()
            && prefix.events.iter().any(|event| match *event {
                crate::program::NativeChoiceEvent::Call { rule, .. } =>
                    self.builder.observes_rule(rule)
                        && !self.builder.failure_prefix_hooks_are_inert(rule),
                _ => false,
            }) {
            return false;
        }
        let diagnostics =
            DIAGNOSTICS && !env.skipping() && !env.predicate() && self.offset >= self.farthest;
        if diagnostics && !prefix.diagnostic_complete {
            return false;
        }
        // Reserve the skipped left subtree, its Choice continuation, and the
        // right Eval. Tight budgets retain real tasks and original error order.
        let Some(cost) = prefix.steps.checked_add(2) else {
            return false;
        };
        if !self
            .steps
            .checked_add(cost)
            .is_some_and(|steps| steps <= self.options.max_steps)
            || !self
                .rule_depth
                .checked_add(prefix.depth)
                .is_some_and(|depth| depth <= self.options.max_rule_depth)
        {
            return false;
        }
        if ACCOUNT {
            self.steps += cost - 1;
        }
        for event in prefix.events {
            match *event {
                crate::program::NativeChoiceEvent::Expression(id) => {
                    self.grammar.observe_expression(id, self.offset)
                }
                crate::program::NativeChoiceEvent::Call { expression, rule } => {
                    self.grammar.observe_expression(expression, self.offset);
                    self.grammar.observe_rule_attempt(
                        rule,
                        self.offset,
                        env.profile_flags(CAPTURE),
                    );
                    if diagnostics {
                        for &id in self.grammar.failure_terminals(rule) {
                            let expected = match self.grammar.instruction(id) {
                                Instruction::Literal { text, .. } => Expected::Literal(text),
                                Instruction::Range { start, end } => Expected::Range(start, end),
                                Instruction::Builtin(Builtin::Newline) => {
                                    Expected::Named("NEWLINE")
                                }
                                Instruction::Builtin(builtin) => Expected::Builtin(builtin),
                                _ => unreachable!("certified rule failure terminal"),
                            };
                            self.failure(env, expected);
                        }
                    }
                }
            }
        }
        self.matched = false;
        true
    }
    fn shared_prefix_valid(&self, spec: crate::program::SharedPrefixChoice, id: ExprId, left: ExprId, right: ExprId, env: E) -> bool {
        fn ungroup<P: Program>(p: &P, mut id: ExprId) -> Option<ExprId> {
            for _ in 0..128 { match p.instruction(id) { Instruction::Group(child) => id=child, _ => return Some(id) } } None
        }
        let target=self.grammar.rule(spec.target); let wrapper=self.grammar.rule(spec.wrapper);
        if target.expression != spec.target_expression || wrapper.expression != spec.branch_expression
            || target.mode != RuleMode::Normal || wrapper.mode != RuleMode::Normal
            || target.scope != spec.scope || wrapper.scope != spec.scope || env.scope() != spec.scope
            || right != spec.right || ungroup(self.grammar,target.expression) != Some(id)
            || ungroup(self.grammar,left) != Some(spec.left_choice)
            || ungroup(self.grammar,wrapper.expression) != Some(spec.branch_sequence)
            || self.grammar.source_pratt(spec.target).is_some()
            || self.grammar.source_pratt(spec.wrapper).is_some()
            || self.grammar.source_pratt(spec.prefix_rule).is_some()
            || spec.suffix_rule.is_some_and(|rule|self.grammar.source_pratt(rule).is_some()) { return false; }
        let Some(wrapper_call)=ungroup(self.grammar,spec.wrapper_call) else{return false};
        let Some(prefix)=ungroup(self.grammar,spec.prefix) else{return false};
        let Some(fallback)=ungroup(self.grammar,right) else{return false};
        matches!(self.grammar.instruction(spec.left_choice),Instruction::Choice{left,right} if left==spec.prior && ungroup(self.grammar,right)==Some(wrapper_call))
            && matches!(self.grammar.instruction(wrapper_call),Instruction::Call(rule) if rule==spec.wrapper)
            && matches!(self.grammar.instruction(spec.branch_sequence),Instruction::Sequence{left,right} if left==spec.prefix && right==spec.suffix)
            && matches!(self.grammar.instruction(prefix),Instruction::Call(rule) if rule==spec.prefix_rule)
            && matches!(self.grammar.instruction(fallback),Instruction::Call(rule) if rule==spec.prefix_rule)
    }
    fn wrapper_prefix_valid(&self,s:crate::program::SharedWrapperPrefix,id:ExprId,left:ExprId,right:ExprId,env:E)->bool {
        fn ug<P:Program>(p:&P,mut e:usize)->usize {for _ in 0..128 {if let Instruction::Group(n)=p.instruction(e){e=n}else{return e}}e}
        fn same<P:Program>(p:&P,a:usize,b:usize,depth:usize)->bool {if depth>16{return false;}let a=ug(p,a);let b=ug(p,b);match(p.instruction(a),p.instruction(b)){
            (Instruction::Literal{text:a,insensitive:ai},Instruction::Literal{text:b,insensitive:bi})=>a==b&&ai==bi,
            (Instruction::Call(a),Instruction::Call(b))=>a==b&&p.source_pratt(a).is_none(),
            (Instruction::Sequence{left:a,right:b},Instruction::Sequence{left:c,right:d})|(Instruction::Choice{left:a,right:b},Instruction::Choice{left:c,right:d})=>same(p,a,c,depth+1)&&same(p,b,d,depth+1),_=>false}}
        let m=self.grammar.rule(s.branch);let p=self.grammar.rule(s.fallback_rule);
        // Structural certification precedes Pratt annotation registration. Even
        // immutable geometry cannot prove that these semantic overrides are absent.
        if self.grammar.source_pratt(s.branch).is_some()
            || self.grammar.source_pratt(s.fallback_rule).is_some() {return false;}
        if self.grammar.shared_wrapper_prefix_metadata_is_current() {
            fn pratt_free<P:Program>(p:&P,e:usize,depth:usize)->bool {
                if depth>16{return false;}
                match p.instruction(ug(p,e)) {
                    Instruction::Literal{..}=>true,
                    Instruction::Call(rule)=>p.source_pratt(rule).is_none(),
                    Instruction::Sequence{left,right}|Instruction::Choice{left,right}=>
                        pratt_free(p,left,depth+1)&&pratt_free(p,right,depth+1),
                    _=>false,
                }
            }
            return id==s.choice && right==s.fallback && m.scope==env.scope()
                && pratt_free(self.grammar,s.prefix,0);
        }
        let scope=self.grammar.scope(m.scope);
        if [scope.whitespace,scope.comment].into_iter().flatten().any(|r|r==s.branch||r==s.fallback_rule){return false;}
        id==s.choice && right==s.fallback && m.mode==RuleMode::Normal&&p.mode==RuleMode::Normal&&m.scope==p.scope&&m.scope==env.scope()
            && self.grammar.source_pratt(s.branch).is_none()&&self.grammar.source_pratt(s.fallback_rule).is_none()
            && matches!(self.grammar.instruction(ug(self.grammar,left)),Instruction::Call(r) if r==s.branch)
            && matches!(self.grammar.instruction(ug(self.grammar,right)),Instruction::Call(r) if r==s.fallback_rule)
            && matches!(self.grammar.instruction(ug(self.grammar,m.expression)),Instruction::Sequence{left,right} if left==s.prefix&&right==s.suffix)
            && same(self.grammar,s.prefix,p.expression,0)
    }
    fn restore_builder(&mut self, mark: usize) {
        if mark != self.builder_mark {
            self.builder.rollback(mark);
            self.builder_mark = self.builder.checkpoint();
        }
    }
    fn restore(&mut self, checkpoint: Checkpoint<E>) {
        if CAPTURE {
            self.nodes.truncate(checkpoint.mark);
            if HYBRID { self.restore_builder(E::semantic_value(checkpoint.undo)); }
        } else {
            self.restore_builder(checkpoint.mark);
        }
        self.restore_matcher(checkpoint);
    }
    fn restore_rejected_pratt(&mut self, checkpoint: Checkpoint<E>, rule: RuleId, span: Span) {
        debug_assert!(CAPTURE && HYBRID);
        let mark = E::semantic_value(checkpoint.undo);
        if mark != self.builder_mark {
            self.builder.rollback_pratt(mark, rule, span);
            self.builder_mark = self.builder.checkpoint();
        }
        // Do not perform a second builder rollback: a granted receipt owner
        // may have journaled replacement observations at the restored mark.
        self.nodes.truncate(checkpoint.mark);
        self.restore_matcher(checkpoint);
    }
    fn restore_matcher(&mut self, checkpoint: Checkpoint<E>) {
        self.offset = checkpoint.offset();
        while self.undo.len() > E::undo_value(checkpoint.undo) {
            match self.undo.pop().unwrap() {
                Undo::RemovePush => {
                    self.stack.pop();
                }
                Undo::RestorePop(value) => self.stack.push(value),
            }
        }
    }
    fn error(&self, kind: ParseErrorKind, message: impl Into<String>) -> ParseError {
        ParseError {
            kind,
            offset: self.offset,
            expected: Vec::new(),
            message: message.into(),
        }
    }
    fn failure(&mut self, env: E, expected: Expected<'g>) {
        self.matched = false;
        if !DIAGNOSTICS || env.predicate() || env.skipping() || self.offset < self.farthest {
            return;
        }
        if self.offset > self.farthest {
            self.farthest = self.offset;
            self.expected.clear();
            self.expected_unique_len = 0;
        }
        self.expected.push(expected);
        // Defer uniqueness checks on the common successful path, but retain
        // at most 512 redundant records above the last compacted unique prefix.
        if self.expected.len() - self.expected_unique_len >= 512 {
            self.dedup_expected();
        }
    }
    fn dedup_expected(&mut self) {
        let mut unique = self.expected_unique_len;
        let first_pending = unique;
        for read in first_pending..self.expected.len() {
            let expected = self.expected[read];
            if !self.expected[..unique].contains(&expected) {
                self.expected[unique] = expected;
                unique += 1;
            }
        }
        self.expected.truncate(unique);
        self.expected_unique_len = unique;
    }
    fn push(&mut self, value: StackValue) {
        self.stack.push(value);
        self.undo.push(Undo::RemovePush);
    }
    fn pop(&mut self) -> Option<StackValue> {
        let value = self.stack.pop()?;
        self.undo.push(Undo::RestorePop(value));
        Some(value)
    }
    fn stack_text(&self, value: StackValue) -> &str {
        match value {
            StackValue::Source(span) => &self.source[span.start..span.end],
            StackValue::Literal(id) => match self.grammar.instruction(id) {
                Instruction::PushLiteral(text) => text,
                _ => unreachable!(),
            },
        }
    }
    fn pratt_operator_impossible(&self, rule: RuleId, env: E) -> bool {
        // On the uninstrumented success path, a certified first-byte miss
        // needs no operator tasks or capture/builder checkpoints. Limited and
        // diagnostic parses retain the original ordering and exact failures.
        if ACCOUNT || DIAGNOSTICS { return false; }
        let mut owned = None;
        let Some(prefix) = failure_prefix(self.grammar, rule, &mut owned) else { return false; };
        let misses = self.source.as_bytes().get(self.offset)
            .is_none_or(|byte| !prefix.contains(*byte));
        if misses && self.rule_depth.checked_add(prefix.depth + 1)
            .is_some_and(|depth| depth <= self.options.max_rule_depth) {
            self.grammar.observe_rule_attempt(rule, self.offset, env.without_semantic().profile_flags(CAPTURE));
            true
        } else { false }
    }

    fn pratt_push(
        &self,
        frame: &mut PrattFrame<E>,
        operator: PendingOperator,
    ) -> Result<(), ParseError> {
        if frame.operators.len() >= self.options.max_rule_depth {
            return Err(self.error(
                ParseErrorKind::DepthLimit,
                "source expression operator nesting limit exceeded",
            ));
        }
        frame.operators.push(operator);
        frame.undo.push(OperatorUndo::Remove);
        Ok(())
    }
    fn pratt_reduce(&mut self, env: E, operator: PendingOperator) -> Result<(), ParseError> {
        if env.semantic() && !env.predicate() && !env.skipping() {
            self.builder.try_reduce(
                operator.operator.rule,
                operator.operator.fixity,
                operator.span,
            )?;
            self.builder_mark = self.builder.checkpoint();
        }
        Ok(())
    }
    fn pratt_finish(&mut self, mut frame: PrattFrame<E>) -> Result<(), ParseError> {
        while let Some(operator) = frame.operators.pop() {
            self.pratt_reduce(frame.env, operator)?;
        }
        if HYBRID && frame.env.semantic() {
            // The island is the body of the pending FinishRule. Reuse that
            // continuation instead of growing every parser task/checkpoint.
            let Some(Task::FinishRule { checkpoint, .. }) = self.tasks.last() else {
                unreachable!("Pratt island must retain its rule completion")
            };
            let span = Span { start: checkpoint.offset(), end: self.offset };
            if !self.builder.accepts_pratt(frame.target, span) {
                let Some(Task::FinishRule { checkpoint, completion }) = self.tasks.pop() else {
                    unreachable!()
                };
                if TaskMode::<CAPTURE, HYBRID>::observed_completion(&completion) {
                    let observed = self.semantic_rules.pop().expect("pending observed Pratt owner");
                    debug_assert_eq!(observed, frame.target);
                }
                let capture = TaskMode::<CAPTURE, HYBRID>::completion_value(completion, checkpoint);
                self.restore_rejected_pratt(checkpoint, frame.target, span);
                if let Some(index) = capture {
                    debug_assert_eq!(index, self.nodes.len());
                    self.nodes.push(Node {
                        rule: CaptureRule::User(frame.target).encode(),
                        span: Span { start: self.offset, end: self.offset },
                        subtree_end: index + 1,
                    });
                }
                self.tasks.push(Task::FinishRule {
                    checkpoint,
                    completion: TaskMode::<CAPTURE, HYBRID>::completion(capture, None),
                });
                self.tasks.push(Task::Eval(
                    self.grammar.rule(frame.target).expression,
                    frame.env.without_semantic()
                        .with_captures_suppressed(frame.enclosing_captures_suppressed),
                ));
            }
        }
        self.matched = true;
        Ok(())
    }
    fn pratt_following(&mut self, frame: Box<PrattFrame<E>>) {
        let checkpoint = self.checkpoint();
        let env = frame.env;
        self.tasks.push(Task::Pratt(Self::pratt_pending(PrattStep::Following, (
            frame, 0, checkpoint, self.offset,
        ))));
        if env.allows_skip() {
            self.tasks.push(Task::Skip(env));
        }
    }
    fn try_terminal_choice(&mut self, id: ExprId, env: E) -> bool {
        let Some(choice) = self.grammar.terminal_choice(id) else {
            return false;
        };
        // All intermediate budget errors retain ordinary dispatch and offsets.
        // Pure terminal choices cannot change stack/captures or rule depth.
        if !self
            .steps
            .checked_add(choice.failure_steps - 1)
            .is_some_and(|steps| steps <= self.options.max_steps)
        {
            return false;
        }
        let remaining = &self.source[self.offset..];
        let candidates = if !DIAGNOSTICS || env.predicate() || env.skipping() || self.offset < self.farthest {
            self.grammar
                .terminal_candidates(id, remaining.as_bytes().first().copied())
        } else {
            None
        };
        let count = candidates.map_or(choice.alternatives.len(), <[usize]>::len);
        for index in 0..count {
            let alternative = &choice.alternatives[candidates.map_or(index, |items| items[index])];
            let (length, expected) = match self.grammar.instruction(alternative.expression) {
                Instruction::Literal { text, insensitive } => (
                    crate::lexical::literal_matches(remaining, text, insensitive)
                        .then_some(text.len()),
                    Expected::Literal(text),
                ),
                Instruction::Range { start, end } => (
                    remaining
                        .chars()
                        .next()
                        .filter(|ch| *ch >= start && *ch <= end)
                        .map(char::len_utf8),
                    Expected::Range(start, end),
                ),
                Instruction::Builtin(builtin) => (
                    remaining
                        .as_bytes()
                        .first()
                        .copied()
                        .filter(|byte| crate::lexical::ascii_predicate(builtin).unwrap()(*byte))
                        .map(|_| 1),
                    Expected::Builtin(builtin),
                ),
                _ => unreachable!("invalid terminal choice certificate"),
            };
            if let Some(length) = length {
                if ACCOUNT { self.steps += alternative.success_steps - 1; }
                self.offset += length;
                self.matched = true;
                return true;
            }
            self.failure(env, expected);
        }
        if ACCOUNT { self.steps += choice.failure_steps - 1; }
        self.matched = false;
        true
    }
    fn literal(&mut self, text: &'g str, insensitive: bool, env: E) {
        let remaining = &self.source[self.offset..];
        let matched = if insensitive {
            remaining
                .as_bytes()
                .get(..text.len())
                .is_some_and(|bytes| bytes.eq_ignore_ascii_case(text.as_bytes()))
        } else {
            remaining.starts_with(text)
        };
        if matched {
            self.offset += text.len();
            self.matched = true;
        } else {
            self.failure(env, Expected::Literal(text));
        }
    }
    #[inline(always)]
    fn charge_step(&mut self) -> Result<(), ParseError> {
        if !ACCOUNT { return Ok(()); }
        if self.steps >= self.options.max_steps {
            return Err(self.error(
                ParseErrorKind::WorkLimit,
                "source parse work limit exceeded",
            ));
        }
        if ACCOUNT { self.steps += 1; }
        Ok(())
    }
    fn finish_repeat(&mut self, frame: usize) {
        debug_assert_eq!(frame + 1, self.repetitions.len());
        self.repetitions.pop();
    }
    fn finish_rule(
        &mut self,
        checkpoint: Checkpoint<E>,
        completion: Option<usize>,
        semantic_completion: Option<usize>,
    ) -> Result<(), ParseError> {
        self.rule_depth -= 1;
        if !self.matched {
            self.restore(checkpoint);
        } else if CAPTURE && let Some(index) = completion {
            self.nodes.finish(index, self.offset);
        }
        if self.matched
            && let Some(rule) = if HYBRID { semantic_completion } else if !CAPTURE { completion } else { None }
        {
            let capture = if HYBRID && CAPTURE && self.builder.observes_capture(rule) {
                completion.map(|index| {
                    let node = self.nodes.get(index);
                    crate::tree::CaptureHandle { index, rule: node.rule, span: node.span,
                        source_ptr: self.source.as_ptr() as usize, source_len: self.source.len() }
                })
            } else { None };
            if capture.is_some() { self.capture_generation = self.capture_generation.saturating_add(1); }
            self.builder.try_finish_capture(
                rule,
                Span { start: checkpoint.offset(), end: self.offset },
                self.source,
                capture,
            )?;
            if HYBRID && CAPTURE && self.builder.wants_completed_capture(rule)
                && let Some(index) = completion {
                let tree = crate::tree::ParseTree::completed_subtree(self.grammar, self.source, &self.nodes, index);
                self.builder.try_completed_capture(rule,
                    Span { start: checkpoint.offset(), end: self.offset }, tree)?;
            }
            self.builder_mark = self.builder.checkpoint();
            if HYBRID && let Some(index) = completion
                && self.builder.prune_capture_children(rule, Span { start: checkpoint.offset(), end: self.offset }) {
                self.nodes.truncate(index + 1);
                self.nodes.finish(index, self.offset);
            }
        }
        Ok(())
    }
    fn ascii_repeat(
        &mut self,
        builtin: Builtin,
        accepts: fn(u8) -> bool,
        env: E,
        min: u32,
        max: Option<u32>,
    ) -> Result<(), ParseError> {
        // Each successful ASCII item performs four virtual tasks. Scan only
        // complete items covered by the remaining budget, then let the ordinary
        // loop handle the final miss or exact resource boundary in its old order.
        let limit = (max.unwrap_or(u32::MAX) as usize)
            .min(self.options.max_steps.saturating_sub(self.steps) / 4);
        let scanned = self.source.as_bytes()[self.offset..]
            .iter()
            .take(limit)
            .take_while(|&&byte| accepts(byte))
            .count();
        self.offset += scanned;
        if ACCOUNT { self.steps += scanned * 4; }
        let mut count = scanned as u32;
        loop {
            // Virtual Repeat task. Keep precisely the same budget boundaries
            // and offsets as the continuation engine, even on resource errors.
            self.charge_step()?;
            if max.is_some_and(|max| count >= max) {
                self.matched = true;
                break;
            }
            self.charge_step()?; // RepeatAfterSkip (no skip in atomic mode)
            self.charge_step()?; // Eval builtin
            if self
                .source
                .as_bytes()
                .get(self.offset)
                .copied()
                .is_some_and(accepts)
            {
                self.offset += 1;
                self.matched = true;
            } else {
                self.failure(env, Expected::Builtin(builtin));
            }
            self.charge_step()?; // RepeatAfterChild
            if !self.matched && (max.is_none() || count < min) {
                self.matched = count >= min;
                break;
            }
            // Failed bounded optional items still charge their later attempts,
            // exactly like the compact continuation path (without input movement).
            count += 1;
        }
        Ok(())
    }
    fn try_empty_success(&self,rule:RuleId,incoming:E)->bool{
        if ACCOUNT||DIAGNOSTICS||!HYBRID||!CAPTURE||incoming.predicate()||incoming.skipping()
            ||!self.grammar.empty_success_metadata_is_current()||!self.grammar.match_observers_are_noop(){return false;}
        let Some(head)=self.grammar.empty_success_head(rule)else{return false};
        if self.source.as_bytes().get(self.offset).is_some_and(|&b|head.first[usize::from(b/64)]&(1u64<<(b%64))!=0){return false;}
        let d=self.grammar.rule(rule);
        d.mode==RuleMode::Silent&&d.expression==head.expression&&d.scope==head.scope&&incoming.scope()==head.scope
            &&self.grammar.source_pratt(rule).is_none()&&self.rule_depth<self.options.max_rule_depth
            &&self.builder.empty_success_is_inert(rule)&&!self.builder.observes_rule(rule)&&!self.builder.observes_capture(rule)
    }
    fn try_empty_suffix(&self,id:usize,child:usize,env:E)->bool{
        if ACCOUNT||DIAGNOSTICS||!HYBRID||!CAPTURE||env.predicate()||env.skipping()
            || !self.grammar.empty_suffix_metadata_is_current()||!self.grammar.match_observers_are_noop(){return false;}
        let Some(h)=self.grammar.empty_suffix_head(id)else{return false;};
        if h.repetition!=id||h.child!=child||h.scope!=env.scope()||h.calls_len as usize>h.calls.len()
            ||self.rule_depth.checked_add(h.depth).is_none_or(|d|d>self.options.max_rule_depth)
            ||self.source.as_bytes().get(self.offset).is_some_and(|&b|h.first[usize::from(b/64)]&(1u64<<(b%64))!=0){return false;}
        let owner=self.grammar.rule(h.owner);
        if owner.expression!=h.owner_body||owner.scope!=h.scope||self.grammar.source_pratt(h.owner).is_some()
            ||!self.builder.empty_suffix_is_inert(h.owner){return false;}
        for &rule in &h.calls[..h.calls_len as usize]{if self.grammar.source_pratt(rule).is_some()||self.builder.observes_rule(rule)||self.builder.observes_capture(rule){return false;}}
        true
    }
    fn scan_fragment(
        &self,
        fragment: &crate::program::AtomicFragment<'g>,
        partial: bool,
    ) -> Option<FragmentOutcome<'g>> {
        let rejected = |cost: usize, expected: &'g [crate::program::FragmentExpected<'g>]| {
            self.steps
                .checked_add(cost)?
                .checked_add(1)
                .filter(|steps| *steps <= self.options.max_steps)
                .map(|_| FragmentOutcome::Rejected(cost, expected))
        };
        let bytes = self.source.as_bytes();
        if let Some(literal) = fragment.literal {
            if !crate::lexical::literal_matches(
                &self.source[self.offset..],
                literal.text,
                literal.insensitive,
            ) {
                return rejected(literal.literal_failure_steps, fragment.head_expected);
            }
            let end = self.offset.checked_add(literal.text.len())?;
            let boundary = bytes
                .get(end)
                .map_or(0, |byte| usize::from(literal.boundary[usize::from(*byte)]));
            let cost = literal.boundary_fixed_steps.checked_add(if boundary == 0 {
                literal.boundary_failure_steps
            } else {
                boundary
            })?;
            if boundary != 0 {
                return rejected(cost, &[]);
            }
            self.steps
                .checked_add(cost)?
                .checked_add(1)
                .filter(|steps| *steps <= self.options.max_steps)?;
            return Some(FragmentOutcome::Matched(end, cost, end, 0));
        }
        let head = bytes
            .get(self.offset)
            .map_or(0, |byte| usize::from(fragment.first[usize::from(*byte)]));
        if partial && head == usize::from(u16::MAX) { return None; }
        if head == 0 {
            if !fragment.head_failure_safe {
                return None;
            }
            let guard_cost = if let Some(guard) = fragment.guard {
                if self.rule_depth >= self.options.max_rule_depth {
                    return None;
                }
                guard
                    .choice_failure_fixed
                    .checked_add(guard.failure_steps)?
            } else {
                0
            };
            let cost = fragment
                .failure_head_fixed
                .checked_add(fragment.head_failure_steps)?
                .checked_add(guard_cost)?;
            return rejected(cost, fragment.head_expected);
        }
        let guard_cost = if let Some(guard) = fragment.guard {
            if self.rule_depth >= self.options.max_rule_depth {
                return None;
            }
            let source = &self.source[self.offset..];
            if let Some(literal) = guard.candidates[usize::from(bytes[self.offset])]
                .iter()
                .map(|&index| &guard.alternatives[usize::from(index)])
                .find(|literal| {
                    crate::lexical::literal_matches(source, literal.text, literal.insensitive)
                })
            {
                let boundary_end = self.offset + literal.text.len();
                let boundary = bytes
                    .get(boundary_end)
                    .map_or(0, |byte| usize::from(guard.boundary[usize::from(*byte)]));
                if boundary == 0 {
                    let cost = fragment
                        .failure_guard_fixed
                        .checked_add(guard.boundary_fixed)?
                        .checked_add(literal.success_steps)?
                        .checked_add(guard.boundary_failure_steps)?;
                    return rejected(cost, &[]);
                }
                guard
                    .boundary_fixed
                    .checked_add(literal.success_steps)?
                    .checked_add(boundary)?
            } else {
                guard
                    .choice_failure_fixed
                    .checked_add(guard.failure_steps)?
            }
        } else {
            0
        };
        let available = self.options.max_steps - self.steps;
        let mut cost = fragment
            .fixed_steps
            .checked_add(head)?
            .checked_add(guard_cost)?;
        if cost > available {
            return None;
        }
        let mut end = self.offset + 1;
        while let Some(&byte) = bytes.get(end) {
            let next = usize::from(fragment.tail[usize::from(byte)]);
            if partial && next == usize::from(u16::MAX) { return None; }
            if next == 0 {
                break;
            }
            cost = cost.checked_add(next + 3)?;
            if cost > available {
                return None;
            }
            end += 1;
        }
        let diagnostic = end;
        let mut suffix_prior = 0;
        if let Some(suffix) = fragment.suffix {
            if let Some((index, literal)) =
                suffix.alternatives.iter().enumerate().find(|(_, literal)| {
                    crate::lexical::literal_matches(
                        &self.source[end..],
                        literal.text,
                        literal.insensitive,
                    )
                })
            {
                cost = cost.checked_add(literal.success_steps + 6)?;
                suffix_prior = index;
                end += 1;
            } else {
                cost = cost.checked_add(suffix.failure_steps + 6)?;
                suffix_prior = suffix.alternatives.len();
            }
            if cost > available {
                return None;
            }
        }
        self.steps
            .checked_add(cost)?
            .checked_add(1)
            .filter(|steps| *steps <= self.options.max_steps)?;
        Some(FragmentOutcome::Matched(
            end,
            cost,
            diagnostic,
            suffix_prior,
        ))
    }
    fn scan_lexical_rule(&mut self, rule: RuleId) -> bool {
        // The scanner skips the expression/rule attempts inside the token.
        // Forwarded certificates alone do not grant permission to skip a
        // custom Program's tracing callbacks.
        if ACCOUNT || DIAGNOSTICS || !self.grammar.match_observers_are_noop() { return false; }
        let native_end = if self.grammar.native_token_metadata_is_current() {
            match self.grammar.native_token_result(rule, self.source, self.offset, self.options.max_rule_depth - self.rule_depth) {
                Some(Some(end)) => Some(end),
                Some(None) => { self.matched = false; return true; },
                None => None,
            }
        } else { None };
        let end = native_end.or_else(|| self.grammar.delimited_fragment(rule)
            .and_then(|fragment| crate::quoted::scan_delimited(fragment, self.source, self.offset)))
            .or_else(|| self.grammar.quoted_fragment(rule)
                .and_then(|fragment| crate::quoted::scan(fragment, self.source, self.offset)));
        if let Some(end) = end {
            self.offset = end;
            self.matched = true;
            true
        } else { false }
    }
    fn atomic_run(&mut self, expression: ExprId, env: E) -> Result<(), ParseError> {
        debug_assert!(self.atomic_tasks.is_empty());
        self.atomic_tasks
            .push(AtomicTask::Eval(expression, env.predicate()));
        while let Some(mut task) = self.atomic_tasks.pop() {
            loop {
                macro_rules! atomic_dispatch {
                    ($next:expr) => {{
                        if self.grammar.atomic_tail_dispatch() {
                            task = $next;
                            continue;
                        }
                        self.atomic_tasks.push($next);
                    }};
                }
                self.charge_step()?;
                match task {
                    AtomicTask::Eval(id, predicate) => {
                        self.grammar.observe_expression(id, self.offset);
                        let env = env.with_predicate(predicate);
                        let instruction = self.grammar.instruction(id);
                        if matches!(
                            instruction,
                            Instruction::Choice { .. } | Instruction::Group(_)
                        ) && self.grammar.has_terminal_choice(id)
                            && self.try_terminal_choice(id, env)
                        {
                            break;
                        }
                        match instruction {
                            Instruction::Literal { text, insensitive } => {
                                self.literal(text, insensitive, env)
                            }
                            Instruction::Range { start, end } => {
                                if let Some(ch) = self.source[self.offset..]
                                    .chars()
                                    .next()
                                    .filter(|ch| *ch >= start && *ch <= end)
                                {
                                    self.offset += ch.len_utf8();
                                    self.matched = true;
                                } else {
                                    self.failure(env, Expected::Range(start, end));
                                }
                            }
                            Instruction::Builtin(builtin) => self.builtin(builtin, env)?,
                            Instruction::Call(rule) => {
                                atomic_dispatch!(AtomicTask::EnterRule(rule, predicate))
                            }
                            Instruction::Group(child) => {
                                atomic_dispatch!(AtomicTask::Eval(child, predicate))
                            }
                            Instruction::Sequence { left, right } => {
                                self.atomic_tasks.push(AtomicTask::Restore(self.offset));
                                self.atomic_tasks
                                    .push(AtomicTask::Sequence(right, predicate));
                                atomic_dispatch!(AtomicTask::Eval(left, predicate));
                            }
                            Instruction::Choice { left, right } => {
                                self.atomic_tasks.push(AtomicTask::Choice(right, predicate));
                                atomic_dispatch!(AtomicTask::Eval(left, predicate));
                            }
                            Instruction::Predicate {
                                expression,
                                positive,
                            } => {
                                self.atomic_tasks
                                    .push(AtomicTask::Predicate(self.offset, positive));
                                atomic_dispatch!(AtomicTask::Eval(expression, true));
                            }
                            Instruction::Repeat {
                                expression,
                                min,
                                max,
                            } => {
                                self.atomic_tasks.push(AtomicTask::Restore(self.offset));
                                atomic_dispatch!(AtomicTask::Repeat(AtomicRepeat {
                                    child: expression,
                                    predicate,
                                    min,
                                    max: max.unwrap_or(u32::MAX),
                                    bounded: max.is_some(),
                                    count: 0,
                                }));
                            }
                            _ => unreachable!("invalid atomic expression certificate"),
                        }
                    }
                    AtomicTask::Restore(start) => {
                        if !self.matched {
                            self.offset = start;
                        }
                    }
                    AtomicTask::Sequence(right, predicate) => {
                        if self.matched {
                            atomic_dispatch!(AtomicTask::Eval(right, predicate));
                        }
                    }
                    AtomicTask::Choice(right, predicate) => {
                        if !self.matched {
                            atomic_dispatch!(AtomicTask::Eval(right, predicate));
                        }
                    }
                    AtomicTask::Predicate(start, positive) => {
                        self.offset = start;
                        self.matched = self.matched == positive;
                    }
                    AtomicTask::EnterRule(rule, predicate) => {
                        self.grammar.observe_rule_attempt(
                            rule,
                            self.offset,
                            env.with_predicate(predicate).profile_flags(CAPTURE),
                        );
                        let mut owned_prefix = None;
                        if (predicate || env.skipping())
                            && let Some(prefix) =
                                failure_prefix(self.grammar, rule, &mut owned_prefix)
                            && self.source.as_bytes().get(self.offset).is_none_or(|byte| {
                                prefix.bytes[usize::from(*byte / 64)] & (1u64 << (*byte % 64)) == 0
                            })
                            && self
                                .rule_depth
                                .checked_add(prefix.depth + 1)
                                .is_some_and(|depth| depth <= self.options.max_rule_depth)
                            && self
                                .steps
                                .checked_add(prefix.steps + 1)
                                .is_some_and(|steps| steps <= self.options.max_steps)
                        {
                            if ACCOUNT { self.steps += prefix.steps + 1; }
                            self.matched = false;
                            break;
                        }
                        if self.rule_depth >= self.options.max_rule_depth {
                            return Err(self.error(
                                ParseErrorKind::DepthLimit,
                                "source rule nesting limit exceeded",
                            ));
                        }
                        if self.scan_lexical_rule(rule) { break; }
                        self.rule_depth += 1;
                        self.atomic_tasks.push(AtomicTask::FinishRule(self.offset));
                        atomic_dispatch!(AtomicTask::Eval(
                            self.grammar.rule(rule).expression,
                            predicate,
                        ));
                    }
                    AtomicTask::FinishRule(start) => {
                        self.rule_depth -= 1;
                        if !self.matched {
                            self.offset = start;
                        }
                    }
                    AtomicTask::Repeat(mut repetition) => {
                        let prefix = self.grammar.leading_bytes(repetition.child).or_else(|| {
                            // A silent rule with a certified byte-only body has no
                            // captures or effects in this atomic matcher. Its call
                            // boundary adds Eval(Call), EnterRule and FinishRule.
                            // Keep the original tasks when that boundary could hit
                            // the depth limit; final misses still use ordinary PEG.
                            let Instruction::Call(rule) = self.grammar.instruction(repetition.child) else { return None; };
                            let definition = self.grammar.rule(rule);
                            if definition.mode != RuleMode::Silent
                                || self.rule_depth >= self.options.max_rule_depth { return None; }
                            self.grammar.leading_bytes(definition.expression)
                                .map(|prefix| crate::lexical::LeadingBytes {
                                    steps: prefix.steps + 3, ..prefix
                                })
                        });
                        if let Some(prefix) = prefix {
                            let cost = prefix.steps + 3; // child, AfterSkip, AfterChild, next Repeat
                            let remaining = self.options.max_steps - self.steps;
                            let count_limit = repetition.max - repetition.count;
                            let limit = (remaining / cost)
                                .min(count_limit as usize)
                                .min(self.source.len() - self.offset);
                            let bytes = &self.source.as_bytes()[self.offset..self.offset + limit];
                            let count = bytes
                                .iter()
                                .position(|byte| !prefix.class.contains(*byte))
                                .unwrap_or(bytes.len());
                            if ACCOUNT { self.steps += count * cost; }
                            self.offset += count;
                            repetition.count += count as u32;
                            if count > 0 {
                                self.matched = true;
                            }
                            // The next Repeat was already charged. Any miss/tight
                            // budget executes real child tasks at this exact offset.
                        }
                        if repetition.bounded && repetition.count >= repetition.max {
                            self.matched = true;
                        } else {
                            self.matched = true;
                            atomic_dispatch!(AtomicTask::AfterSkip(repetition, self.offset));
                        }
                    }
                    AtomicTask::AfterSkip(repetition, start) => {
                        self.atomic_tasks
                            .push(AtomicTask::AfterChild(repetition, start));
                        atomic_dispatch!(AtomicTask::Eval(repetition.child, repetition.predicate));
                    }
                    AtomicTask::AfterChild(mut repetition, start) => {
                        if !self.matched {
                            self.offset = start;
                            if repetition.bounded && repetition.count >= repetition.min {
                                repetition.count += 1;
                                atomic_dispatch!(AtomicTask::Repeat(repetition));
                            } else {
                                self.matched = repetition.count >= repetition.min;
                            }
                        } else {
                            if self.offset == start && !repetition.bounded {
                                return Err(self.error(
                                    ParseErrorKind::NonProgress,
                                    "unbounded repetition succeeded without consuming input",
                                ));
                            }
                            repetition.count += 1;
                            atomic_dispatch!(AtomicTask::Repeat(repetition));
                        }
                    }
                }
                break;
            }
        }
        Ok(())
    }
    fn pratt_frame(mut frame: Box<PrattFrame<E>>, step: PrattStep, candidate: usize)
        -> <TaskMode<CAPTURE, HYBRID> as PrattTaskMode<E>>::Frame {
        frame.step = step;
        frame.candidate = candidate;
        TaskMode::<CAPTURE, HYBRID>::wrap(frame)
    }
    fn pratt_pending(step: PrattStep, value: (Box<PrattFrame<E>>, usize, Checkpoint<E>, usize))
        -> <TaskMode<CAPTURE, HYBRID> as PrattTaskMode<E>>::Frame {
        let (mut frame, candidate, checkpoint, start) = value;
        frame.following = Some((candidate, checkpoint, start));
        Self::pratt_frame(frame, step, candidate)
    }
    fn pratt_pending_value(mut frame: Box<PrattFrame<E>>)
        -> (Box<PrattFrame<E>>, usize, Checkpoint<E>, usize) {
        let (candidate, checkpoint, start) = frame.following.take().expect("pending Pratt continuation");
        (frame, candidate, checkpoint, start)
    }
    // Ordinary PEG tasks share one small dispatch arm. The expression state
    // machine stays out of the large hot matcher without another allocation.
    #[inline(never)]
    fn execute_pratt(&mut self, frame: Box<PrattFrame<E>>) -> Result<(), ParseError> {
        match frame.step {
                PrattStep::Operand => {
                    let candidate = frame.candidate;
                    let spec = self.grammar.source_pratt(frame.target).unwrap();
                    if let Some((candidate, operator)) = spec
                        .operators
                        .iter()
                        .enumerate()
                        .skip(candidate)
                        .find(|(_, op)| op.fixity == crate::pratt::Fixity::Prefix
                            && !self.pratt_operator_impossible(op.rule, frame.env))
                    {
                        let operator = *operator;
                        let env = frame.env.without_semantic();
                        self.tasks.push(Task::Pratt(Self::pratt_pending(PrattStep::AfterPrefix, (
                            frame, candidate, self.checkpoint(), self.offset,
                        ))));
                        self.tasks.push(Task::EnterRule(operator.rule, env));
                    } else {
                        let atom = if let Some((_,candidate))=frame.attempt.as_ref() {
                            let operator=spec.operators[*candidate].rule;
                            self.grammar.source_pratt_operator_rhs_atom(frame.target,operator).unwrap_or_else(|| {
                                if self.grammar.source_pratt_rhs_operator(frame.target).is_none_or(|rule|rule==operator) {
                                    self.grammar.source_pratt_rhs_atom(frame.target).unwrap_or(spec.atom)
                                } else {spec.atom}
                            })
                        } else { spec.atom };
                        let env = frame.env;
                        self.tasks
                            .push(Task::Pratt(Self::pratt_frame(frame, PrattStep::AfterAtom, 0)));
                        self.tasks.push(Task::EnterRule(atom, env));
                    }
                }
                PrattStep::AfterPrefix => {
                    let (mut frame, candidate, checkpoint, _) =
                        Self::pratt_pending_value(frame);
                    if self.matched {
                        if self.offset == checkpoint.offset() {
                            return Err(self.error(
                                ParseErrorKind::NonProgress,
                                "Pratt prefix succeeded without consuming input",
                            ));
                        }
                        let operator =
                            self.grammar.source_pratt(frame.target).unwrap().operators[candidate];
                        self.pratt_push(
                            &mut frame,
                            PendingOperator {
                                operator,
                                span: Span {
                                    start: checkpoint.offset(),
                                    end: self.offset,
                                },
                            },
                        )?;
                        let env = frame.env;
                        self.tasks.push(Task::Pratt(Self::pratt_frame(frame, PrattStep::Operand, 0)));
                        if env.allows_skip() {
                            self.tasks.push(Task::Skip(env));
                        }
                    } else {
                        self.tasks.push(Task::Pratt(Self::pratt_frame(frame, PrattStep::Operand, candidate + 1)));
                    }
                }
                PrattStep::AfterAtom => {
                    let mut frame = frame;
                    if self.matched {
                        frame.attempt = None;
                        frame.undo.clear();
                        self.pratt_following(frame);
                    } else if let Some((checkpoint, candidate)) = frame.attempt.take() {
                        // A grammar repetition rolls back a consumed infix whose
                        // RHS failed. Restore popped operator metadata as well as
                        // builder values/input; eager reductions may have happened.
                        self.restore(checkpoint);
                        while let Some(undo) = frame.undo.pop() {
                            match undo {
                                OperatorUndo::Remove => {
                                    frame.operators.pop().unwrap();
                                }
                                OperatorUndo::Restore(op) => frame.operators.push(op),
                            }
                        }
                        // A shorter operator can have consumed the prefix of a
                        // later alternative (Lua '~' versus '~='). Its failed RHS
                        // must allow the remaining ordered candidates to run.
                        let env = frame.env;
                        self.tasks.push(Task::Pratt(Self::pratt_pending(PrattStep::Following, (
                            frame, candidate + 1, checkpoint, self.offset,
                        ))));
                        if env.allows_skip() {
                            self.tasks.push(Task::Skip(env));
                        }
                    }
                    // Initial operand failure leaves matched=false. FinishRule
                    // restores the expression's complete transactional checkpoint.
                }
                PrattStep::Following => {
                    let (frame, candidate, mut checkpoint, _) =
                        Self::pratt_pending_value(frame);
                    let spec = self.grammar.source_pratt(frame.target).unwrap();
                    if candidate == 0
                        && frame.env.semantic()
                        && !frame.env.predicate()
                        && !frame.env.skipping()
                    {
                        self.builder.try_pratt_boundary(frame.target, self.offset)?;
                        self.builder_mark = self.builder.checkpoint();
                    }
                    if candidate == 0
                        && spec.trailing_trivia == crate::islands::TrailingTrivia::Consume
                    {
                        checkpoint = self.checkpoint();
                    }
                    if let Some((candidate, operator)) = spec
                        .operators
                        .iter()
                        .enumerate()
                        .skip(candidate)
                        .find(|(_, op)| op.fixity != crate::pratt::Fixity::Prefix
                            && !self.pratt_operator_impossible(op.rule, frame.env))
                    {
                        let operator = *operator;
                        let env = frame.env.without_semantic();
                        self.tasks
                            .push(Task::Pratt(Self::pratt_pending(PrattStep::AfterFollowing, (
                                frame,
                                candidate,
                                checkpoint,
                                self.offset,
                            ))));
                        self.tasks.push(Task::EnterRule(operator.rule, env));
                    } else {
                        self.restore(checkpoint); // keep trailing trivia for caller
                        self.pratt_finish(*frame)?;
                    }
                }
                PrattStep::AfterOperatorTrivia => {
                    let env=frame.env;
                    if self.matched {
                        self.tasks.push(Task::Pratt(Self::pratt_frame(frame,PrattStep::Operand,0)));
                        if env.allows_skip() {self.tasks.push(Task::Skip(env));}
                    } else {
                        // Treat a failed separator as a failed RHS repetition;
                        // the original infix checkpoint restores input, captures,
                        // stack effects, reductions and builder metadata together.
                        self.tasks.push(Task::Pratt(Self::pratt_frame(frame,PrattStep::AfterAtom,0)));
                    }
                }
                PrattStep::AfterFollowing => {
                    let (mut frame, candidate, checkpoint, start) =
                        Self::pratt_pending_value(frame);
                    if !self.matched {
                        self.tasks.push(Task::Pratt(Self::pratt_pending(PrattStep::Following, (
                            frame, candidate + 1, checkpoint, self.offset,
                        ))));
                    } else {
                        if self.offset == start {
                            return Err(self.error(
                                ParseErrorKind::NonProgress,
                                "Pratt operator succeeded without consuming input",
                            ));
                        }
                        let operator =
                            self.grammar.source_pratt(frame.target).unwrap().operators[candidate];
                        while frame.operators.last().is_some_and(|top| {
                            top.operator.precedence > operator.precedence
                                || (top.operator.precedence == operator.precedence
                                    && operator.fixity != crate::pratt::Fixity::InfixRight)
                        }) {
                            let top = frame.operators.pop().unwrap();
                            frame.undo.push(OperatorUndo::Restore(top));
                            self.pratt_reduce(frame.env, top)?;
                        }
                        let pending = PendingOperator {
                            operator,
                            span: Span {
                                start,
                                end: self.offset,
                            },
                        };
                        if operator.fixity == crate::pratt::Fixity::Postfix {
                            self.pratt_reduce(frame.env, pending)?;
                            frame.undo.clear();
                            self.pratt_following(frame);
                        } else {
                            self.pratt_push(&mut frame, pending)?;
                            frame.attempt = Some((checkpoint, candidate));
                            let env = frame.env;
                            if let Some(trivia)=self.grammar.source_pratt_operator_trivia(frame.target) {
                                self.tasks.push(Task::Pratt(Self::pratt_frame(frame,PrattStep::AfterOperatorTrivia,0)));
                                self.tasks.push(Task::EnterRule(trivia,env));
                            } else {
                                self.tasks.push(Task::Pratt(Self::pratt_frame(frame, PrattStep::Operand, 0)));
                            }
                            if env.allows_skip() {self.tasks.push(Task::Skip(env));}
                        }
                    }
                }
        }
        Ok(())
    }
    fn execute(&mut self, mut task: ExecutionTask<E, CAPTURE, HYBRID>) -> Result<(), ParseError> {
        loop {
            macro_rules! dispatch {
                ($next:expr) => {{
                    let next = $next;
                    if self.grammar.tail_dispatch() && self.grammar.tail_continuations() {
                        self.charge_step()?;
                        task = next;
                        continue;
                    }
                    self.tasks.push(next);
                }};
            }
            let covered_sequence = matches!(&task, Task::EvalCoveredSequence(..));
            match task {
                Task::Pratt(frame) => self.execute_pratt(TaskMode::<CAPTURE, HYBRID>::unwrap(frame))?,
                Task::RunBuiltin(builtin, env) => self.builtin(builtin, env)?,
                Task::Eval(id, env) | Task::EvalCoveredSequence(id, env) => {
                    self.grammar.observe_expression(id, self.offset);
                    let instruction = self.grammar.instruction(id);
                    if matches!(
                        instruction,
                        Instruction::Choice { .. } | Instruction::Group(_)
                    ) && self.grammar.has_terminal_choice(id)
                        && self.try_terminal_choice(id, env)
                    {
                        return Ok(());
                    }
                    if (env.predicate() || env.skipping())
                        && matches!(
                            instruction,
                            Instruction::Choice { .. } | Instruction::Group(_)
                        )
                        && let Some(cost) = self
                            .grammar
                            .literal_choice_failure(id, &self.source[self.offset..])
                        && self
                            .steps
                            .checked_add(cost - 1)
                            .is_some_and(|steps| steps <= self.options.max_steps)
                    {
                        // A pure literal-choice failure leaves input/state intact
                        // and emits no diagnostics in these environments. Tight
                        // budgets run the original tasks to preserve error order.
                        if ACCOUNT { self.steps += cost - 1; } // root Eval already charged
                        self.matched = false;
                        return Ok(());
                    }
                    // Terminals cannot leave partial input/captures/stack on
                    // failure. Calls and predicates already own their rollback;
                    // choice/group/push delegate failure rollback to their child.
                    // Only these operations can fail after earlier mutations.
                    if matches!(
                        instruction,
                        Instruction::Sequence { .. }
                            | Instruction::Repeat { .. }
                            | Instruction::PeekSlice { .. }
                            | Instruction::Builtin(Builtin::PeekAll | Builtin::PopAll)
                    ) && !(covered_sequence && matches!(instruction, Instruction::Sequence { .. })) {
                        self.tasks.push(Task::FinishExpr(self.checkpoint()));
                    }
                    match instruction {
                        Instruction::Literal { text, insensitive } => {
                            self.literal(text, insensitive, env)
                        }
                        Instruction::Range { start, end } => {
                            if let Some(ch) = self.source[self.offset..]
                                .chars()
                                .next()
                                .filter(|ch| *ch >= start && *ch <= end)
                            {
                                self.offset += ch.len_utf8();
                                self.matched = true;
                            } else {
                                self.failure(env, Expected::Range(start, end));
                            }
                        }
                        Instruction::Call(rule) => {
                            if self.grammar.tail_dispatch() {
                                self.charge_step()?;
                                task = Task::EnterRule(rule, env);
                                continue;
                            }
                            self.tasks.push(Task::EnterRule(rule, env));
                        }
                        Instruction::Builtin(builtin) => self.builtin(builtin, env)?,
                        Instruction::Sequence { left, right } => {
                            self.tasks.push(Task::Sequence { right, env });
                            // If the left sequence fails, this sequence cannot
                            // continue. Its own FinishExpr restores all three
                            // journals before an enclosing choice can retry.
                            let left = if !ACCOUNT && !DIAGNOSTICS {
                                Task::EvalCoveredSequence(left, env)
                            } else { Task::Eval(left, env) };
                            if self.grammar.tail_dispatch() {
                                self.charge_step()?;
                                task = left;
                                continue;
                            }
                            self.tasks.push(left);
                        }
                        Instruction::Choice { left, right } => {
                            if !ACCOUNT && !DIAGNOSTICS && CAPTURE && HYBRID && self.grammar.choice_heads_candidate(id) && self.try_choice_heads(id,env){return Ok(());}
                            if CAPTURE && HYBRID && self.options.max_steps==usize::MAX
                                && env.semantic() && !env.predicate() && !env.skipping()
                                && self.grammar.match_observers_are_noop()
                                && let Some(spec)=self.grammar.shared_inner_prefix(id)
                                && self.builder.supports_shared_inner_prefix(spec.target)
                                && crate::inner_prefix::valid(self.grammar,spec,id,left,right,env.scope()) {
                                let start=self.checkpoint();
                                self.tasks.push(Task::SharedInner(Box::new(SharedInnerFrame{spec,env,start,prefix:None,step:1})));
                                dispatch!(Task::Eval(spec.prefix,env));return Ok(());
                            }
                            if CAPTURE && HYBRID && self.options.max_steps==usize::MAX
                                && self.source.as_bytes().get(self.offset)==Some(&b'(')
                                && env.semantic() && !env.predicate() && !env.skipping()
                                && self.grammar.match_observers_are_noop()
                                && let Some(spec)=self.grammar.shared_group_prefix(id)
                                && self.builder.supports_shared_group_prefix(spec.branch,spec.fallback_rule)
                                && !self.builder.observes_rule(spec.branch) && !self.builder.observes_rule(spec.fallback_rule)
                                && !self.builder.observes_capture(spec.branch) && !self.builder.observes_capture(spec.fallback_rule)
                                && crate::group_prefix::valid(self.grammar,spec,id,left,right,env.scope()) {
                                let start=self.checkpoint();
                                self.tasks.push(Task::SharedGroup(Box::new(SharedGroupFrame{spec,env,start,wrapper:None,prefix:None,step:1,capture:false})));
                                dispatch!(Task::Eval(spec.open,env));return Ok(());
                            }
                            if CAPTURE && HYBRID
                                && let Some(spec)=self.grammar.shared_wrapper_prefix(id)
                                && self.builder.supports_shared_wrapper_prefix_at(spec.branch,spec.fallback_rule,self.offset,self.source)
                                && self.options.max_steps==usize::MAX
                                && env.semantic() && !env.predicate() && !env.skipping()
                                && self.grammar.match_observers_are_noop()
                                && !self.builder.observes_rule(spec.branch) && !self.builder.observes_rule(spec.fallback_rule)
                                && self.wrapper_prefix_valid(spec,id,left,right,env) {
                                let start=self.checkpoint();
                                let frame=self.shared_wrapper_frames.len();
                                if self.shared_wrapper_frames.capacity()==0 {self.shared_wrapper_frames.reserve_exact(1);}
                                self.shared_wrapper_frames.push(SharedWrapperFrame{spec,env,start,prefix:None,step:0,capture:false});
                                self.tasks.push(Task::SharedWrapper(frame));
                                return Ok(());
                            }

                            if CAPTURE && HYBRID && self.options.max_steps == usize::MAX
                                && env.semantic() && !env.predicate() && !env.skipping()
                                && self.grammar.match_observers_are_noop()
                                && let Some(spec) = self.grammar.shared_prefix_choice(id)
                                && self.builder.supports_shared_prefix(spec.wrapper)
                                && !self.builder.observes_rule(spec.wrapper)
                                && self.shared_prefix_valid(spec, id, left, right, env)
                            {
                                let start = self.checkpoint();
                                self.tasks.push(Task::SharedPrefix(Box::new(SharedPrefixFrame { spec, env, step: 0, start, prefix: None, capture: false, generation: self.capture_generation })));
                                dispatch!(Task::Eval(spec.prior, env));
                                return Ok(());
                            }
                            if HYBRID && self.grammar.native_hybrid_choices() && self.grammar.fast_rule_failures()
                                && let Some(prefix) = self.grammar.native_hybrid_choice_prefix(id)
                                && self.source.as_bytes().get(self.offset).is_none_or(|byte|
                                    prefix.bytes[usize::from(*byte / 64)] & (1u64 << (*byte % 64)) == 0)
                                && self.try_native_choice_prefix(prefix, env) {
                                dispatch!(Task::Eval(right, env));
                                return Ok(());
                            }
                            self.tasks.push(Task::Choice { right, env });
                            if self.grammar.tail_dispatch() {
                                self.charge_step()?;
                                task = Task::Eval(left, env);
                                continue;
                            }
                            self.tasks.push(Task::Eval(left, env));
                        }
                        Instruction::Repeat {
                            expression,
                            min,
                            max,
                        } => {
                            if HYBRID && CAPTURE && !ACCOUNT && !DIAGNOSTICS && min==0 && max==Some(1)
                                && self.grammar.empty_suffix_candidate(id) && self.try_empty_suffix(id,expression,env){self.matched=true;return Ok(());}
                            if self.grammar.fast_ascii_repetitions()
                                && env.mode() != Atomicity::Normal
                                && let Instruction::Builtin(builtin) =
                                    self.grammar.instruction(expression)
                                && let Some(accepts) = crate::lexical::ascii_predicate(builtin)
                            {
                                self.ascii_repeat(builtin, accepts, env, min, max)?;
                            } else {
                                let frame = self.repetitions.len();
                                self.repetitions.push(RepeatFrame {
                                    repetition: Repetition {
                                        child: expression,
                                        env,
                                        min,
                                        max,
                                        count: 0,
                                    },
                                    // Assigned by the original Repeat task before
                                    // any continuation reads it; do not add a
                                    // semantic builder checkpoint callback here.
                                    checkpoint: Checkpoint {
                                        offset: 0,
                                        undo: E::undo_mark(0),
                                        mark: 0,
                                    },
                                });
                                self.tasks.push(Task::Repeat(frame));
                            }
                        }
                        Instruction::Predicate {
                            expression,
                            positive,
                        } => {
                            self.tasks.push(Task::Predicate {
                                checkpoint: self.passive_checkpoint(),
                                positive,
                            });
                            self.tasks
                                .push(Task::Eval(expression, env.with_predicate(true)));
                        }
                        Instruction::Group(child) => {
                            if self.grammar.tail_dispatch() {
                                self.charge_step()?;
                                task = Task::Eval(child, env);
                                continue;
                            }
                            self.tasks.push(Task::Eval(child, env));
                        }
                        Instruction::Push(child) => {
                            self.tasks.push(Task::Push { start: self.offset });
                            self.tasks.push(Task::Eval(child, env));
                        }
                        Instruction::PushLiteral(_) => {
                            self.push(StackValue::Literal(id));
                            self.matched = true;
                        }
                        Instruction::PeekSlice { start, end } => {
                            self.peek_slice(start, end, false, env)?
                        }
                    }
                }
                Task::FinishExpr(checkpoint) => {
                    if !self.matched {
                        self.restore(checkpoint);
                    }
                }
                Task::EnterRule(rule, incoming) => {
                    if HYBRID && CAPTURE && !ACCOUNT && !DIAGNOSTICS
                        && self.grammar.empty_success_candidate(rule)
                        && self.try_empty_success(rule,incoming){
                        self.matched=true;return Ok(());}
                    // Root EnterRule was already charged. Only an explicit hybrid
                    // builder and generated certificate can skip this pure failure.
                    if HYBRID && incoming.semantic() && !incoming.predicate() && !incoming.skipping()
                        && self.builder.supports_unobserved_failure_prefixes()
                        && self.grammar.match_observers_are_noop()
                        && let Some(prefix) = self.grammar.nullable_failure_prefix(rule, incoming.flags() & 3, self.source.as_bytes().get(self.offset).copied())
                        && prefix.steps > 0
                        && prefix.unobserved_rules.iter().all(|&rule| (!self.builder.observes_rule(rule) && !self.builder.observes_capture(rule)) || self.builder.failure_prefix_hooks_are_inert(rule))
                        && self.rule_depth.checked_add(prefix.depth).is_some_and(|depth|depth <= self.options.max_rule_depth)
                        && self.steps.checked_add(prefix.steps - 1).is_some_and(|steps|steps <= self.options.max_steps)
                    {
                        if ACCOUNT { self.steps += prefix.steps - 1; }
                        for &expected in prefix.expected {
                            self.failure(incoming, match expected {
                                crate::program::FragmentExpected::Literal(text)=>Expected::Literal(text),
                                crate::program::FragmentExpected::Range(a,b)=>Expected::Range(a,b),
                                crate::program::FragmentExpected::Builtin(Builtin::Newline)=>Expected::Named("NEWLINE"),
                                crate::program::FragmentExpected::Builtin(b)=>Expected::Builtin(b),
                            });
                        }
                        self.matched=false;
                        return Ok(());
                    }
                    self.grammar.observe_rule_attempt(
                        rule,
                        self.offset,
                        incoming.profile_flags(CAPTURE),
                    );
                    let mut owned_prefix = None;
                    // The compact first-byte proof records matcher work, not
                    // descendant semantic callbacks. Hybrid normal matching must
                    // enter the original rules so checked begin errors and hook
                    // attempts cannot disappear on a lexical miss. The explicit
                    // nullable-head proof above already checks every reached hook.
                    let general_failure = CAPTURE && self.grammar.fast_rule_failures()
                        && (!HYBRID || !incoming.semantic()
                            || incoming.predicate() || incoming.skipping());
                    let diagnostics = DIAGNOSTICS && !incoming.skipping()
                        && !incoming.predicate()
                        && self.offset >= self.farthest;
                    if (incoming.skipping()
                        || general_failure
                            && (!diagnostics || !self.grammar.failure_terminals(rule).is_empty())
                        || (CAPTURE
                            && incoming.predicate()
                            && self.grammar.fast_predicate_failures()))
                        && let Some(prefix) = failure_prefix(self.grammar, rule, &mut owned_prefix)
                        && self.source.as_bytes().get(self.offset).is_none_or(|byte| {
                            prefix.bytes[usize::from(*byte / 64)] & (1u64 << (*byte % 64)) == 0
                        })
                        && self
                            .rule_depth
                            .checked_add(prefix.depth + 1)
                            .is_some_and(|depth| depth <= self.options.max_rule_depth)
                        && self
                            .steps
                            .checked_add(prefix.steps + 1)
                            .is_some_and(|steps| steps <= self.options.max_steps)
                    {
                        // Only a certified suppressed failure is bypassed. Reserve the
                        // full original cost/depth; tight limits use real tasks so
                        // work-versus-depth error ordering is unchanged.
                        if ACCOUNT { self.steps += prefix.steps + 1; } // expression + FinishRule
                        if diagnostics {
                            for &id in self.grammar.failure_terminals(rule) {
                                let expected = match self.grammar.instruction(id) {
                                    Instruction::Literal { text, .. } => Expected::Literal(text),
                                    Instruction::Range { start, end } => {
                                        Expected::Range(start, end)
                                    }
                                    Instruction::Builtin(Builtin::Newline) => {
                                        Expected::Named("NEWLINE")
                                    }
                                    Instruction::Builtin(builtin) => Expected::Builtin(builtin),
                                    _ => unreachable!("invalid failure certificate terminal"),
                                };
                                self.failure(incoming, expected);
                            }
                        }
                        self.matched = false;
                        return Ok(());
                    }
                    if self.rule_depth >= self.options.max_rule_depth {
                        return Err(self.error(
                            ParseErrorKind::DepthLimit,
                            "source rule nesting limit exceeded",
                        ));
                    }
                    self.rule_depth += 1;
                    let checkpoint = self.checkpoint();
                    let definition = self.grammar.rule(rule);
                    let mode = match definition.mode {
                        RuleMode::Atomic => Atomicity::Atomic,
                        RuleMode::CompoundAtomic => Atomicity::Compound,
                        RuleMode::NonAtomic => Atomicity::Normal,
                        _ => incoming.mode(),
                    };
                    let capture_mode = if matches!(
                        definition.mode,
                        RuleMode::CompoundAtomic | RuleMode::NonAtomic
                    ) {
                        mode
                    } else {
                        incoming.mode()
                    };
                    let capture = if CAPTURE
                        && (!HYBRID || !incoming.captures_suppressed())
                        && !incoming.predicate()
                        && definition.mode != RuleMode::Silent
                        && capture_mode != Atomicity::Atomic
                    {
                        let index = self.nodes.len();
                        self.nodes.push(Node {
                            rule: CaptureRule::User(rule).encode(),
                            span: Span {
                                start: self.offset,
                                end: self.offset,
                            },
                            subtree_end: index + 1,
                        });
                        Some(index)
                    } else {
                        None
                    };
                    // The rule's node was appended at checkpoint.mark. Its
                    // presence fits the spare source-offset bit, so hybrid
                    // FinishRule tasks need only the semantic completion ID.
                    let checkpoint = if HYBRID {
                        checkpoint.with_rule_capture(capture.is_some())
                    } else { checkpoint };
                    let builder_rule =
                        (!incoming.predicate() && !incoming.skipping() && incoming.semantic() && self.builder.observes_rule(rule))
                            .then_some(rule);
                    if builder_rule.is_some() {
                        self.builder.try_begin(rule, self.offset)?;
                        self.builder_mark = self.builder.checkpoint();
                        if HYBRID { self.semantic_rules.push(rule); }
                    }
                    // Designated trivia rules have atomic bodies, including
                    // explicit calls. Their own capture uses the declared rule
                    // mode above; explicit !/$ child rules can enable captures.
                    let scope = self.grammar.scope(definition.scope);
                    let mode = if scope.whitespace == Some(rule) || scope.comment == Some(rule) {
                        if definition.mode == RuleMode::CompoundAtomic {
                            Atomicity::Compound
                        } else {
                            Atomicity::Atomic
                        }
                    } else {
                        mode
                    };
                    let env = E::enter(
                        self.grammar,
                        rule,
                        definition.scope,
                        (incoming.flags() & !3) | mode as u8,
                    );
                    let pratt = (!CAPTURE || HYBRID)
                        && incoming.semantic()
                        && !incoming.skipping()
                        && self.grammar.source_pratt(rule).is_some()
                        && self.builder.supports_pratt_at(rule, self.offset, self.source);
                    let completion=TaskMode::<CAPTURE,HYBRID>::completion(if CAPTURE{capture}else{builder_rule},builder_rule);
                    // Ordinary/finite/observational paths never query ownership.
                    let owned_end=if CAPTURE&&HYBRID&&!ACCOUNT&&!DIAGNOSTICS&&!pratt
                        &&capture.is_some()&&builder_rule.is_some()&&!incoming.captures_suppressed()
                        &&incoming.semantic()&&!incoming.predicate()&&!incoming.skipping()
                        &&self.builder.supports_owned_capture_boundary(rule)
                        &&definition.mode==RuleMode::Normal
                        &&![scope.whitespace,scope.comment].contains(&Some(rule))
                        &&self.grammar.match_observers_are_noop()
                        &&self.grammar.owned_capture_boundary_metadata_is_current()
                        &&let Some(cert)=self.grammar.owned_capture_boundary(rule)
                        &&cert.expression==definition.expression&&cert.scope==definition.scope
                        &&cert.capture_rules.contains(&rule)
                        &&cert.capture_rules.iter().all(|&r|r==rule||!self.builder.observes_capture(r)) {
                            self.builder.owned_capture_end(rule,self.offset,self.source)
                                .filter(|&end|end>=self.offset&&end<=self.source.len()&&self.source.is_char_boundary(end))
                        }else{None};
                    let env=if let Some(end)=owned_end{
                        self.tasks.push(Task::OwnedCapture(Box::new(OwnedCaptureFrame{checkpoint,body_checkpoint:self.checkpoint(),completion,expression:definition.expression,env,end})));
                        env.with_captures_suppressed(true)
                    }else{self.tasks.push(Task::FinishRule{checkpoint,completion});env};
                    if ((CAPTURE && !HYBRID)
                        || (incoming.semantic() && self.builder.opaque_rule(rule))
                        // Trivia and lookahead suppress semantic hooks. Their
                        // certified atomic bodies can use the same matcher as
                        // capture-only parsing without skipping builder work.
                        || (!ACCOUNT && !DIAGNOSTICS
                            && (!incoming.semantic() || incoming.predicate() || incoming.skipping())))
                        && env.mode() == Atomicity::Atomic
                        && self.grammar.atomic_expression(definition.expression)
                    {
                        if self.scan_lexical_rule(rule) { return Ok(()); }
                        if FRAGMENT
                            && let Some((fragment, partial)) = self.grammar.partial_atomic_fragment(rule)
                                .map(|fragment| (fragment, true))
                                .or_else(|| self.grammar.atomic_fragment(rule).map(|fragment| (fragment, false)))
                            && let Some(outcome) = self.scan_fragment(fragment, partial)
                        {
                            if let FragmentOutcome::Rejected(cost, expected) = outcome {
                                if let Some(guard) = fragment.guard {
                                    self.grammar.observe_rule_attempt(
                                        guard.rule,
                                        self.offset,
                                        env.with_predicate(true).profile_flags(CAPTURE),
                                    );
                                }
                                if ACCOUNT { self.steps += cost; }
                                for &expected in expected {
                                    let expected = match expected {
                                        crate::program::FragmentExpected::Builtin(b) => {
                                            Expected::Builtin(b)
                                        }
                                        crate::program::FragmentExpected::Literal(s) => {
                                            Expected::Literal(s)
                                        }
                                        crate::program::FragmentExpected::Range(a, b) => {
                                            Expected::Range(a, b)
                                        }
                                    };
                                    self.failure(env, expected);
                                }
                                self.matched = false;
                                return Ok(());
                            }
                            let FragmentOutcome::Matched(end, cost, diagnostic, suffix_prior) =
                                outcome
                            else {
                                unreachable!()
                            };
                            if let Some(guard) = fragment.guard {
                                self.grammar.observe_rule_attempt(
                                    guard.rule,
                                    self.offset,
                                    env.with_predicate(true).profile_flags(CAPTURE),
                                );
                            }
                            if ACCOUNT { self.steps += cost; }
                            self.offset = diagnostic;
                            for &expected in fragment.expected {
                                let expected = match expected {
                                    crate::program::FragmentExpected::Builtin(b) => {
                                        Expected::Builtin(b)
                                    }
                                    crate::program::FragmentExpected::Literal(s) => {
                                        Expected::Literal(s)
                                    }
                                    crate::program::FragmentExpected::Range(a, b) => {
                                        Expected::Range(a, b)
                                    }
                                };
                                self.failure(env, expected);
                            }
                            if let Some(suffix) = fragment.suffix {
                                for &expected in &suffix.expected[..suffix_prior] {
                                    let crate::program::FragmentExpected::Literal(text) = expected
                                    else {
                                        unreachable!("certified literal suffix")
                                    };
                                    self.failure(env, Expected::Literal(text));
                                }
                            }
                            self.offset = end;
                            self.matched = true;
                            return Ok(());
                        }
                        self.atomic_run(definition.expression, env)?;
                        return Ok(());
                    }
                    if pratt {
                        let enclosing_captures_suppressed = env.captures_suppressed();
                        let env = if HYBRID && self.builder.suppress_pratt_captures(rule) {
                            env.with_captures_suppressed(true)
                        } else { env };
                        self.tasks.push(Task::Pratt(TaskMode::<CAPTURE, HYBRID>::wrap(Box::new(PrattFrame {
                                target: rule,
                                env,
                                enclosing_captures_suppressed,
                                step: PrattStep::Operand,
                                candidate: 0,
                                operators: Vec::new(),
                                undo: Vec::new(),
                                attempt: None,
                                following: None,
                            }))));
                    } else {
                        if HYBRID && self.grammar.tail_dispatch() && self.grammar.native_hybrid_choices()
                            && let Some(head) = self.grammar.native_hybrid_rule_head(rule) {
                            for (index, &(expression, right)) in head.sequences.iter().enumerate() {
                                self.charge_step()?;
                                self.grammar.observe_expression(expression, self.offset);
                                if index == 0 || ACCOUNT || DIAGNOSTICS {
                                    self.tasks.push(Task::FinishExpr(self.checkpoint()));
                                }
                                self.tasks.push(Task::Sequence { right, env });
                            }
                            self.charge_step()?;
                            self.grammar.observe_expression(head.call_expression, self.offset);
                            self.charge_step()?;
                            task = Task::EnterRule(head.callee, env);
                            continue;
                        }
                        if self.grammar.tail_dispatch() {
                            self.charge_step()?;
                            task = Task::Eval(definition.expression, env);
                            continue;
                        }
                        self.tasks.push(Task::Eval(definition.expression, env));
                    }
                }
                Task::OwnedCapture(frame)=>{
                    let OwnedCaptureFrame{checkpoint,body_checkpoint,completion,expression,env,end}=*frame;
                    if self.matched&&self.offset!=end{
                        self.restore(body_checkpoint);
                        self.tasks.push(Task::FinishRule{checkpoint,completion});
                        self.tasks.push(Task::Eval(expression,env));
                    }else{
                        let semantic=if TaskMode::<CAPTURE,HYBRID>::observed_completion(&completion){Some(self.semantic_rules.pop().expect("pending owned rule"))}else{None};
                        self.finish_rule(checkpoint,TaskMode::<CAPTURE,HYBRID>::completion_value(completion,checkpoint),semantic)?;
                    }
                }
                Task::FinishRule {
                    checkpoint,
                    completion,
                } => {
                    let semantic = if TaskMode::<CAPTURE, HYBRID>::observed_completion(&completion) {
                        Some(self.semantic_rules.pop().expect("pending observed rule"))
                    } else { None };
                    self.finish_rule(
                        checkpoint,
                        TaskMode::<CAPTURE, HYBRID>::completion_value(completion, checkpoint),
                        semantic,
                    )?;
                }
                Task::Sequence { right, env } => {
                    if self.matched {
                        if env.allows_skip() {
                            // Include Skip and AfterSkip in the certificate's cost.
                            // Tight budgets retain their original task ordering.
                            if self.try_fast_trivia_failure(env, false, 2) {
                                dispatch!(Task::Eval(right, env));
                            } else {
                                self.tasks.push(Task::AfterSkip { right, env });
                                dispatch!(Task::Skip(env));
                            }
                        } else {
                            dispatch!(Task::Eval(right, env));
                        }
                    }
                }
                Task::AfterSkip { right, env } => {
                    if self.matched {
                        dispatch!(Task::Eval(right, env));
                    }
                }
                Task::SharedInner(mut f)=>{match f.step {
                    1=>{
                        if !self.matched || self.offset==f.start.offset(){self.restore(f.start);self.tasks.push(Task::Choice{right:f.spec.right,env:f.env});dispatch!(Task::Eval(f.spec.left,f.env));}
                        else{f.prefix=Some(self.checkpoint());f.step=2;let right=f.spec.first;let env=f.env;self.tasks.push(Task::SharedInner(f));dispatch!(Task::Sequence{right,env});}
                    }
                    2=>{
                        if f.spec.kind==0 {
                            if !self.matched{self.restore(f.prefix.unwrap());self.matched=true;f.step=3;let right=f.spec.second;let env=f.env;self.tasks.push(Task::SharedInner(f));dispatch!(Task::Sequence{right,env});}
                        }else if !self.matched{self.restore(f.prefix.unwrap());self.matched=true;}
                        else{f.step=3;let right=f.spec.second;let env=f.env;self.tasks.push(Task::SharedInner(f));dispatch!(Task::Sequence{right,env});}
                    }
                    3=>{if self.matched{f.step=4;let right=f.spec.third;let env=f.env;self.tasks.push(Task::SharedInner(f));dispatch!(Task::Sequence{right,env});}else{self.restore(f.start);}}
                    4=>{if !self.matched{self.restore(f.start);}}
                    _=>unreachable!(),}}
                Task::SharedGroup(mut f)=>{match f.step {
                    1=>{
                        if !self.matched {self.restore(f.start);dispatch!(Task::Eval(f.spec.right,f.env));}
                        else {f.step=2;let env=f.env;self.tasks.push(Task::SharedGroup(f));if env.allows_skip(){dispatch!(Task::Skip(env));}}
                    }
                    2=>{
                        if !self.matched {self.restore(f.start);dispatch!(Task::Eval(f.spec.right,f.env));}
                        else {
                            if self.rule_depth>=self.options.max_rule_depth{return Err(self.error(ParseErrorKind::DepthLimit,"source rule nesting limit exceeded"));}
                            self.rule_depth+=1;f.wrapper=Some(self.checkpoint());f.capture=!f.env.captures_suppressed()&&f.env.mode()!=Atomicity::Atomic;
                            if f.capture{self.nodes.push(Node{rule:CaptureRule::User(f.spec.branch).encode(),span:Span{start:self.offset,end:self.offset},subtree_end:self.nodes.len()+1});}
                            f.step=3;let expr=f.spec.prefix;let env=f.env;self.tasks.push(Task::SharedGroup(f));dispatch!(Task::Eval(expr,env));
                        }
                    }
                    3=>{
                        if !self.matched{self.rule_depth-=1;self.restore(f.start);dispatch!(Task::Eval(f.spec.right,f.env));}
                        else{f.prefix=Some(self.checkpoint());f.step=4;let right=f.spec.branch_tail;let env=f.env;self.tasks.push(Task::SharedGroup(f));dispatch!(Task::Sequence{right,env});}
                    }
                    4=>{
                        if self.matched{
                            self.rule_depth-=1;if f.capture{self.nodes.finish(f.wrapper.unwrap().mark,self.offset);}
                            f.step=5;let right=f.spec.close;let env=f.env;self.tasks.push(Task::SharedGroup(f));dispatch!(Task::Sequence{right,env});
                        }else{
                            self.restore(f.prefix.unwrap());if f.capture{self.nodes.retag_wrapper(f.wrapper.unwrap().mark,CaptureRule::User(f.spec.fallback_rule).encode());}
                            self.matched=true;f.step=7;let right=f.spec.commas;let env=f.env;self.tasks.push(Task::SharedGroup(f));dispatch!(Task::Sequence{right,env});
                        }
                    }
                    5=>{
                        if !self.matched{
                            self.restore(f.prefix.unwrap());self.rule_depth+=1;if f.capture{self.nodes.retag_wrapper(f.wrapper.unwrap().mark,CaptureRule::User(f.spec.fallback_rule).encode());}
                            self.matched=true;f.step=7;let right=f.spec.commas;let env=f.env;self.tasks.push(Task::SharedGroup(f));dispatch!(Task::Sequence{right,env});
                        }
                    }
                    7=>{
                        f.step=8;let right=f.spec.trailing;let env=f.env;self.tasks.push(Task::SharedGroup(f));dispatch!(Task::Sequence{right,env});
                    }
                    8=>{
                        self.rule_depth-=1;if f.capture{self.nodes.finish(f.wrapper.unwrap().mark,self.offset);}
                        f.step=9;let right=f.spec.fallback_close;let env=f.env;self.tasks.push(Task::SharedGroup(f));dispatch!(Task::Sequence{right,env});
                    }
                    9=>{if !self.matched{self.restore(f.start);}}
                    _=>unreachable!(),}}
                Task::SharedWrapper(index)=>{
                    debug_assert_eq!(index+1,self.shared_wrapper_frames.len(),"wrapper frames are nested LIFO");
                    let mut f=self.shared_wrapper_frames.pop().expect("live wrapper frame");
                    match f.step {
                    0=>{if self.rule_depth>=self.options.max_rule_depth{return Err(self.error(ParseErrorKind::DepthLimit,"source rule nesting limit exceeded"));}
                        self.rule_depth+=1;f.capture=!f.env.captures_suppressed()&&f.env.mode()!=Atomicity::Atomic;
                        if f.capture{self.nodes.push(Node{rule:CaptureRule::User(f.spec.branch).encode(),span:Span{start:self.offset,end:self.offset},subtree_end:self.nodes.len()+1});}
                        f.step=1;let prefix=f.spec.prefix;let env=f.env;self.shared_wrapper_frames.push(f);self.tasks.push(Task::SharedWrapper(index));dispatch!(Task::Eval(prefix,env));},
                    1=>{if !self.matched{self.rule_depth-=1;self.restore(f.start);dispatch!(Task::Eval(f.spec.fallback,f.env));}else{f.prefix=Some(self.checkpoint());f.step=2;let right=f.spec.suffix;let env=f.env;self.shared_wrapper_frames.push(f);self.tasks.push(Task::SharedWrapper(index));dispatch!(Task::Sequence{right,env});}},
                    2=>{self.rule_depth-=1;if !self.matched{self.restore(f.prefix.unwrap());if f.capture{self.nodes.retag_wrapper(f.start.mark,CaptureRule::User(f.spec.fallback_rule).encode());}self.matched=true;}if f.capture{self.nodes.finish(f.start.mark,self.offset);}},
                    _=>unreachable!(),}}
                Task::SharedPrefix(mut frame) => {
                    match frame.step {
                        0 => {
                            if self.matched { return Ok(()); }
                            self.restore(frame.start);
                            if self.rule_depth >= self.options.max_rule_depth {
                                return Err(self.error(ParseErrorKind::DepthLimit, "source rule nesting limit exceeded"));
                            }
                            frame.generation = self.capture_generation;
                            self.rule_depth += 1;
                            frame.capture = !frame.env.captures_suppressed() && frame.env.mode() != Atomicity::Atomic;
                            if frame.capture { self.nodes.push(Node { rule: CaptureRule::User(frame.spec.wrapper).encode(), span: Span { start: self.offset, end: self.offset }, subtree_end: self.nodes.len()+1 }); }
                            frame.step = 1;
                            let prefix=frame.spec.prefix; let env=frame.env;
                            self.tasks.push(Task::SharedPrefix(frame));
                            dispatch!(Task::Eval(prefix, env));
                        }
                        1 => {
                            if !self.matched {
                                self.rule_depth -= 1;
                                self.restore(frame.start);
                                dispatch!(Task::Eval(frame.spec.prefix, frame.env));
                            } else {
                                frame.prefix=Some(self.checkpoint()); frame.step=2;
                                let suffix=frame.spec.suffix; let env=frame.env;
                                self.tasks.push(Task::SharedPrefix(frame));
                                dispatch!(Task::Sequence { right:suffix, env });
                            }
                        }
                        2 => {
                            self.rule_depth -= 1;
                            if self.matched {
                                if frame.capture { self.nodes.finish(frame.start.mark, self.offset); }
                            } else {
                                if frame.capture && (self.capture_generation != frame.generation || self.capture_generation == usize::MAX) {
                                    // Prefix tokens may point after this wrapper. Roll back the
                                    // speculative branch and rebuild the ordinary fallback instead.
                                    self.restore(frame.start);
                                    dispatch!(Task::Eval(frame.spec.right, frame.env));
                                } else {
                                    self.restore(frame.prefix.unwrap());
                                    if frame.capture { self.nodes.collapse_wrapper(frame.start.mark); }
                                    self.matched=true;
                                }
                            }
                        }
                        _ => unreachable!(),
                    }
                }
                Task::ChoiceHeads(mut frame)=>{
                    if !self.matched {
                        let i=frame.remaining.trailing_zeros() as usize;frame.remaining&=frame.remaining-1;
                        let expression=frame.branches[i].expression;let env=frame.env;
                        if frame.remaining!=0{self.tasks.push(Task::ChoiceHeads(frame));}
                        dispatch!(Task::Eval(expression,env));
                    }
                }
                Task::Choice { right, env } => {
                    if !self.matched {
                        dispatch!(Task::Eval(right, env));
                    }
                }
                Task::Repeat(frame) => {
                    let repetition = self.repetitions[frame].repetition;
                    if repetition.max.is_some_and(|max| repetition.count >= max) {
                        self.matched = true;
                        self.finish_repeat(frame);
                    } else {
                        self.repetitions[frame].checkpoint = self.checkpoint();
                        self.tasks.push(Task::RepeatAfterSkip(frame));
                        if repetition.count > 0 && repetition.env.allows_skip() {
                            dispatch!(Task::Skip(repetition.env));
                        } else {
                            self.matched = true;
                        }
                    }
                }
                Task::RepeatAfterSkip(frame) => {
                    if self.matched {
                        let repetition = self.repetitions[frame].repetition;
                        // Preserve the old checkpoint boundary for bounded optional
                        // items and the first attempt after a required plus tail.
                        if (repetition.max.is_some() && repetition.count >= repetition.min)
                            || (repetition.max.is_none()
                                && repetition.min > 0
                                && repetition.count == repetition.min)
                        {
                            self.repetitions[frame].checkpoint = self.checkpoint();
                        }
                        self.tasks.push(Task::RepeatAfterChild(frame));
                        dispatch!(Task::Eval(repetition.child, repetition.env));
                    } else {
                        self.finish_repeat(frame);
                    }
                }
                Task::RepeatAfterChild(frame) => {
                    let repetition = self.repetitions[frame].repetition;
                    let checkpoint = self.repetitions[frame].checkpoint;
                    if !self.matched {
                        self.restore(checkpoint);
                        if repetition.max.is_some() && repetition.count >= repetition.min {
                            self.repetitions[frame].repetition.count += 1;
                            dispatch!(Task::Repeat(frame));
                        } else {
                            self.matched = repetition.count >= repetition.min;
                            self.finish_repeat(frame);
                        }
                    } else {
                        if self.offset == checkpoint.offset() && repetition.max.is_none() {
                            return Err(self.error(
                                ParseErrorKind::NonProgress,
                                "unbounded repetition succeeded without consuming input",
                            ));
                        }
                        self.repetitions[frame].repetition.count += 1;
                        dispatch!(Task::Repeat(frame));
                    }
                }
                Task::Predicate {
                    checkpoint,
                    positive,
                } => {
                    let matched = self.matched;
                    self.restore_passive(checkpoint);
                    self.matched = matched == positive;
                }
                Task::Push { start } => {
                    if self.matched {
                        self.push(StackValue::Source(Span {
                            start,
                            end: self.offset,
                        }));
                    }
                }
                Task::Skip(env) => {
                    if !ACCOUNT && self.rule_depth < self.options.max_rule_depth {
                        if let Some(class) = self.grammar.whitespace_run_class(env.scope()) {
                            self.offset += self.source.as_bytes()[self.offset..].iter()
                                .take_while(|&&byte| class.contains(byte)).count();
                        }
                    }

                    if let Some(class) = self.grammar.scope(env.scope()).whitespace_prefix_class {
                        // Preserve depth guards even though the certified rule's
                        // tasks are removed. Charge each scanned byte to the work
                        // budget; large trivia cannot bypass parser limits.
                        if self.rule_depth >= self.options.max_rule_depth {
                            return Err(self.error(
                                ParseErrorKind::DepthLimit,
                                "source rule nesting limit exceeded",
                            ));
                        }
                        while self
                            .source
                            .as_bytes()
                            .get(self.offset)
                            .is_some_and(|byte| class.contains(*byte))
                        {
                            if self.steps >= self.options.max_steps {
                                return Err(self.error(
                                    ParseErrorKind::WorkLimit,
                                    "source parse work limit exceeded",
                                ));
                            }
                            if ACCOUNT { self.steps += 1; }
                            self.offset += 1;
                        }
                        if self.try_fast_trivia_failure(
                            env,
                            self.grammar.scope(env.scope()).whitespace_class.is_some(),
                            0,
                        ) {
                            return Ok(());
                        }
                        if self.grammar.scope(env.scope()).whitespace_class.is_none() {
                            // Prefix scan consumed only successful leading terminal
                            // alternatives. Evaluate the unchanged full rule at
                            // the first miss, retaining all fallback effects.
                            let checkpoint = self.passive_checkpoint();
                            self.tasks.push(Task::AfterWhitespace { checkpoint, env });
                            self.tasks.push(Task::EnterRule(
                                self.grammar.scope(env.scope()).whitespace.unwrap(),
                                env.with_skipping(),
                            ));
                            return Ok(());
                        }
                        let checkpoint = self.passive_checkpoint();
                        self.tasks.push(Task::AfterComment { checkpoint, env });
                        if let Some(rule) = self.grammar.scope(env.scope()).comment {
                            self.tasks.push(Task::EnterRule(rule, env.with_skipping()));
                        } else {
                            self.matched = false;
                        }
                        return Ok(());
                    }
                    if self.try_fast_trivia_failure(env, false, 0) {
                        return Ok(());
                    }
                    let checkpoint = self.passive_checkpoint();
                    let skip_env = env.with_skipping();
                    self.tasks.push(Task::AfterWhitespace { checkpoint, env });
                    if let Some(rule) = self.grammar.scope(env.scope()).whitespace {
                        self.tasks.push(Task::EnterRule(rule, skip_env));
                    } else {
                        self.matched = false;
                    }
                }
                Task::AfterWhitespace { checkpoint, env } => {
                    if self.matched {
                        if self.offset == checkpoint.offset() {
                            return Err(self.error(
                                ParseErrorKind::NonProgress,
                                "WHITESPACE succeeded without consuming input",
                            ));
                        }
                        self.tasks.push(Task::Skip(env));
                    } else {
                        self.restore_passive(checkpoint);
                        self.tasks.push(Task::AfterComment { checkpoint, env });
                        if let Some(rule) = self.grammar.scope(env.scope()).comment {
                            self.tasks.push(Task::EnterRule(rule, env.with_skipping()));
                        } else {
                            self.matched = false;
                        }
                    }
                }
                Task::AfterComment { checkpoint, env } => {
                    if self.matched {
                        if self.offset == checkpoint.offset() {
                            return Err(self.error(
                                ParseErrorKind::NonProgress,
                                "COMMENT succeeded without consuming input",
                            ));
                        }
                        self.tasks.push(Task::Skip(env));
                    } else {
                        self.restore_passive(checkpoint);
                        self.matched = true;
                    }
                }
            }
            return Ok(());
        }
    }

    fn builtin(&mut self, builtin: Builtin, env: E) -> Result<(), ParseError> {
        use Builtin::*;
        match builtin {
            Soi => {
                if self.offset == 0 {
                    self.matched = true;
                } else {
                    self.failure(env, Expected::Named("SOI"));
                }
            }
            Eoi => {
                if self.offset == self.source.len() {
                    self.matched = true;
                    if CAPTURE && (!HYBRID || !env.captures_suppressed())
                        && !env.predicate() && env.mode() != Atomicity::Atomic {
                        let index = self.nodes.len();
                        self.nodes.push(Node {
                            rule: CaptureRule::EndOfInput.encode(),
                            span: Span {
                                start: self.offset,
                                end: self.offset,
                            },
                            subtree_end: index + 1,
                        });
                    }
                } else {
                    self.failure(env, Expected::Named("EOI"));
                }
            }
            Newline => {
                let remaining = &self.source[self.offset..];
                if remaining.starts_with("\r\n") {
                    self.offset += 2;
                    self.matched = true;
                } else if remaining.starts_with(['\r', '\n']) {
                    self.offset += 1;
                    self.matched = true;
                } else {
                    self.failure(env, Expected::Named("NEWLINE"));
                }
            }
            Peek | Pop => {
                let value = *self.stack.last().ok_or_else(|| {
                    self.error(
                        ParseErrorKind::InvalidStack,
                        "PEEK/POP requires a nonempty grammar stack",
                    )
                })?;
                self.match_stack_value(value, env);
                if self.matched && builtin == Pop {
                    self.pop();
                }
            }
            PeekAll | PopAll => {
                for index in (0..self.stack.len()).rev() {
                    self.match_stack_value(self.stack[index], env);
                    if !self.matched {
                        return Ok(());
                    }
                }
                self.matched = true;
                if builtin == PopAll {
                    while self.pop().is_some() {}
                }
            }
            Drop => {
                self.matched = self.pop().is_some();
                if !self.matched {
                    self.failure(env, Expected::Named("DROP"));
                }
            }
            _ => {
                let ch = self.source[self.offset..].chars().next();
                let matched = ch.is_some_and(|ch| match builtin {
                    Any => true,
                    Ascii => ch.is_ascii(),
                    AsciiDigit => ch.is_ascii_digit(),
                    AsciiNonzeroDigit => ('1'..='9').contains(&ch),
                    AsciiBinDigit => matches!(ch, '0' | '1'),
                    AsciiOctDigit => ('0'..='7').contains(&ch),
                    AsciiHexDigit => ch.is_ascii_hexdigit(),
                    AsciiAlphaLower => ch.is_ascii_lowercase(),
                    AsciiAlphaUpper => ch.is_ascii_uppercase(),
                    AsciiAlpha => ch.is_ascii_alphabetic(),
                    AsciiAlphanumeric => ch.is_ascii_alphanumeric(),
                    XidStart => unicode_ident::is_xid_start(ch),
                    XidContinue => unicode_ident::is_xid_continue(ch),
                    _ => unreachable!(),
                });
                if matched {
                    self.offset += ch.unwrap().len_utf8();
                    self.matched = true;
                } else {
                    self.failure(env, Expected::Builtin(builtin));
                }
            }
        }
        Ok(())
    }
    fn match_stack_value(&mut self, value: StackValue, env: E) {
        let length = self.stack_text(value).len();
        if self.source[self.offset..].starts_with(self.stack_text(value)) {
            self.offset += length;
            self.matched = true;
        } else {
            self.failure(env, Expected::Named("stack text"));
        }
    }
    fn peek_slice(
        &mut self,
        start: i32,
        end: Option<i32>,
        reverse: bool,
        env: E,
    ) -> Result<(), ParseError> {
        let length = self.stack.len() as i64;
        let normalize = |index: i32| {
            let index = i64::from(index);
            if index < 0 { length + index } else { index }
        };
        let start = normalize(start);
        let end = end.map(normalize).unwrap_or(length);
        if start < 0 || end < 0 || start > length || end > length {
            return Err(self.error(
                ParseErrorKind::InvalidStack,
                "PEEK slice index is out of range",
            ));
        }
        self.matched = true;
        if reverse {
            for index in (start as usize..end as usize).rev() {
                self.match_stack_value(self.stack[index], env);
                if !self.matched {
                    break;
                }
            }
        } else {
            for index in start as usize..end as usize {
                self.match_stack_value(self.stack[index], env);
                if !self.matched {
                    break;
                }
            }
        }
        Ok(())
    }
}

#[cfg(all(test, target_pointer_width = "64"))]
mod task_layout_experiment {
    #[test]
    fn compact_tasks() {
        assert_eq!(std::mem::size_of::<super::ExecutionTask<super::Env, true>>(), 48);
        assert_eq!(std::mem::size_of::<super::ExecutionTask<super::Env, false>>(), 48);
    }
}

#[cfg(test)]
mod compact_layout_tests {
    use super::*;
    #[test]
    fn hybrid_rule_capture_tag_preserves_full_width_cursors() {
        type E = HybridEnv<StackFreeEnv>;
        for offset in [0, 1, isize::MAX as usize] {
            for mark in [0, usize::MAX] {
                let checkpoint = Checkpoint::<E> {
                    offset, mark,
                    undo: E::semantic_mark(E::undo_mark(0), usize::MAX),
                };
                for capture in [false, true] {
                    let tagged = checkpoint.with_rule_capture(capture);
                    assert_eq!(tagged.offset(), offset);
                    assert_eq!(tagged.captures_rule(), capture);
                    assert_eq!(E::semantic_value(tagged.undo), usize::MAX);
                    let completion = <TaskMode<true, true> as PrattTaskMode<E>>::completion(
                        capture.then_some(mark), Some(7),
                    );
                    assert_eq!(<TaskMode<true, true> as PrattTaskMode<E>>::observed_completion(&completion), true);
                    assert_eq!(<TaskMode<true, true> as PrattTaskMode<E>>::completion_value(completion, tagged), capture.then_some(mark));
                }
            }
        }
    }
    type WideCaptureTask = Task<
        <TaskMode<true> as PrattTaskMode<Env>>::Frame,
        Env,
        <TaskMode<true> as PrattTaskMode<Env>>::Completion,
    >;
    type CompactCaptureTask = Task<
        <TaskMode<true> as PrattTaskMode<CompactEnv>>::Frame,
        CompactEnv,
        <TaskMode<true> as PrattTaskMode<CompactEnv>>::Completion,
    >;
    type BuilderTask = Task<
        <TaskMode<false> as PrattTaskMode<Env>>::Frame,
        Env,
        <TaskMode<false> as PrattTaskMode<Env>>::Completion,
    >;
    #[test]
    #[cfg(target_pointer_width = "64")]
    fn sizes() {
        assert_eq!(std::mem::size_of::<Env>(), 16);
        assert_eq!(std::mem::size_of::<CompactEnv>(), 8);
        assert_eq!(std::mem::size_of::<WideCaptureTask>(), 48);
        assert_eq!(std::mem::size_of::<CompactCaptureTask>(), 40);
        assert_eq!(std::mem::size_of::<BuilderTask>(), 48);
        assert_eq!(std::mem::size_of::<ExecutionTask<HybridEnv<Env>, true, true>>(), 48);
        assert_eq!(std::mem::size_of::<ExecutionTask<HybridEnv<CompactEnv>, true, true>>(), 40);
        assert_eq!(std::mem::size_of::<ExecutionTask<HybridEnv<StackFreeEnv>, true, true>>(), 32);
    }
    #[test]
    fn boundary_scopes_and_completions() {
        let env = Env::enter(
            &crate::compile("entry={ANY}").unwrap(),
            usize::MAX,
            usize::MAX,
            63,
        );
        assert_eq!(env.scope(), usize::MAX);
        assert_eq!(
            Env::capture_completion_value(Env::capture_completion(Some(usize::MAX))),
            Some(usize::MAX)
        );
        for scope in [0, 1, 17, usize::MAX >> 6] {
            for flags in 0..64 {
                let word = crate::program::CompactScope::new(scope);
                let env = CompactEnv(word.word() | flags as usize);
                assert_eq!(env.scope(), scope);
                assert_eq!(env.flags(), flags);
                for predicate in [true, false] {
                    assert_eq!(env.with_predicate(predicate).scope(), scope);
                }
            }
        }
        for index in [0, 1, 17, usize::MAX - 1] {
            assert_eq!(
                CompactEnv::capture_completion_value(CompactEnv::capture_completion(Some(index))),
                Some(index)
            );
        }
    }
}

#[cfg(all(test, target_pointer_width = "64"))]
mod stackfree_layout_tests {
    use super::*;
    type StackFreeTask = Task<
        <TaskMode<true> as PrattTaskMode<StackFreeEnv>>::Frame,
        StackFreeEnv,
        <TaskMode<true> as PrattTaskMode<StackFreeEnv>>::Completion,
    >;
    #[test]
    fn actual_stackfree_layout() {
        assert_eq!(std::mem::size_of::<Checkpoint<StackFreeEnv>>(), 16);
        assert_eq!(std::mem::size_of::<Checkpoint<CompactEnv>>(), 24);
        assert_eq!(std::mem::size_of::<StackFreeEnv>(), 8);
        assert_eq!(std::mem::size_of::<StackFreeTask>(), 32);
    }
}

#[cfg(test)]
#[test]
fn fragment_specialization_preserves_wide_and_compact_state_layout() {
    type WideDisabled = State<'static, 'static, CompiledGrammar, NoBuilder, true, Env, false>;
    type WideEnabled = State<'static, 'static, CompiledGrammar, NoBuilder, true, Env, true>;
    type CompactDisabled =
        State<'static, 'static, CompiledGrammar, NoBuilder, true, CompactEnv, false>;
    type CompactEnabled =
        State<'static, 'static, CompiledGrammar, NoBuilder, true, CompactEnv, true>;
    assert_eq!(
        std::mem::size_of::<WideDisabled>(),
        std::mem::size_of::<WideEnabled>()
    );
    assert_eq!(
        std::mem::align_of::<WideDisabled>(),
        std::mem::align_of::<WideEnabled>()
    );
    assert_eq!(
        std::mem::size_of::<CompactDisabled>(),
        std::mem::size_of::<CompactEnabled>()
    );
    assert_eq!(
        std::mem::align_of::<CompactDisabled>(),
        std::mem::align_of::<CompactEnabled>()
    );
}
