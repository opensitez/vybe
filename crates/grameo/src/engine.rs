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
    pub max_steps: usize,
    pub max_rule_depth: usize,
}
impl Default for MatchOptions {
    fn default() -> Self {
        Self {
            max_steps: 10_000_000,
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
    let state =
        State::<P, NoBuilder, true, E, FRAGMENT>::run(program, source, entry, options, NoBuilder)?;
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

impl CompiledGrammar {
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
        let state = State::<Self, NoBuilder, true>::run(self, source, entry, options, NoBuilder)?;
        Ok(ParseTree {
            grammar: self,
            source,
            nodes: state.nodes,
            consumed: state.offset,
        })
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
    fn undo_mark(len: usize) -> Self::UndoMark;
    fn undo_value(mark: Self::UndoMark) -> usize;
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
    // Capture parsing uses NoBuilder; semantic builders never emit captures.
    mark: usize,
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
struct PrattFrame<E: Environment> {
    target: RuleId,
    env: E,
    operators: Vec<PendingOperator>,
    undo: Vec<OperatorUndo>,
    attempt: Option<(Checkpoint<E>, usize)>,
}

// Capture mode never constructs Pratt continuations. An uninhabited payload
// keeps those variants out of its task layout while builder mode keeps its box.
struct TaskMode<const CAPTURE: bool>;
trait PrattTaskMode<E: Environment> {
    type Completion;
    type AfterFollowing;
    fn completion(value: Option<usize>) -> Self::Completion;
    fn completion_value(value: Self::Completion) -> Option<usize>;
    fn following(value: (Box<PrattFrame<E>>, usize, Checkpoint<E>, usize)) -> Self::AfterFollowing;
    fn following_value(
        value: Self::AfterFollowing,
    ) -> (Box<PrattFrame<E>>, usize, Checkpoint<E>, usize);
    type Frame;
    fn wrap(frame: Box<PrattFrame<E>>) -> Self::Frame;
    fn unwrap(frame: Self::Frame) -> Box<PrattFrame<E>>;
}
impl<E: Environment> PrattTaskMode<E> for TaskMode<true> {
    type Completion = E::CaptureCompletion;
    type AfterFollowing = std::convert::Infallible;
    fn completion(value: Option<usize>) -> Self::Completion {
        E::capture_completion(value)
    }
    fn completion_value(value: Self::Completion) -> Option<usize> {
        E::capture_completion_value(value)
    }
    fn following(_: (Box<PrattFrame<E>>, usize, Checkpoint<E>, usize)) -> Self::AfterFollowing {
        unreachable!("capture mode cannot create Pratt continuations")
    }
    fn following_value(
        value: Self::AfterFollowing,
    ) -> (Box<PrattFrame<E>>, usize, Checkpoint<E>, usize) {
        match value {}
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
    type AfterFollowing = (Box<PrattFrame<E>>, usize, Checkpoint<E>, usize);
    fn completion(value: Option<usize>) -> Self::Completion {
        value
    }
    fn completion_value(value: Self::Completion) -> Option<usize> {
        value
    }
    fn following(value: Self::AfterFollowing) -> Self::AfterFollowing {
        value
    }
    fn following_value(value: Self::AfterFollowing) -> Self::AfterFollowing {
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

enum Task<F, E: Environment, C, A> {
    PrattOperand {
        frame: F,
        candidate: usize,
    },
    PrattAfterPrefix {
        frame: F,
        candidate: usize,
        checkpoint: Checkpoint<E>,
    },
    PrattAfterAtom(F),
    PrattFollowing {
        frame: F,
        candidate: usize,
        checkpoint: Checkpoint<E>,
    },
    PrattAfterFollowing(A),
    RunBuiltin(Builtin, E),
    Eval(ExprId, E),
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
        checkpoint: Checkpoint<E>,
        positive: bool,
    },
    Push {
        start: usize,
    },
    Skip(E),
    AfterWhitespace {
        checkpoint: Checkpoint<E>,
        env: E,
    },
    AfterComment {
        checkpoint: Checkpoint<E>,
        env: E,
    },
}

type ExecutionTask<E, const CAPTURE: bool> = Task<
    <TaskMode<CAPTURE> as PrattTaskMode<E>>::Frame,
    E,
    <TaskMode<CAPTURE> as PrattTaskMode<E>>::Completion,
    <TaskMode<CAPTURE> as PrattTaskMode<E>>::AfterFollowing,
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
> where
    TaskMode<CAPTURE>: PrattTaskMode<E>,
{
    grammar: &'g P,
    builder: B,
    source: &'i str,
    offset: usize,
    nodes: Nodes,
    stack: Vec<StackValue>,
    undo: Vec<Undo>,
    tasks: Vec<ExecutionTask<E, CAPTURE>>,
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

impl<'g, 'i, P: Program, B: Builder, const CAPTURE: bool, E: Environment, const FRAGMENT: bool>
    State<'g, 'i, P, B, CAPTURE, E, FRAGMENT>
where
    TaskMode<CAPTURE>: PrattTaskMode<E>,
{
    fn run(
        grammar: &'g P,
        source: &'i str,
        entry: &str,
        options: MatchOptions,
        builder: B,
    ) -> Result<Self, ParseError> {
        let entry = grammar
            .rule_id(entry)
            .map(Reference::Rule)
            .or_else(|| Builtin::from_name(entry).map(Reference::Builtin))
            .ok_or_else(|| ParseError {
                kind: ParseErrorKind::UnknownEntry,
                offset: 0,
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
            let hint = (source.len() / 4).clamp(4, 32).min(options.max_steps);
            if hint != 0 {
                nodes.reserve(1usize << (usize::BITS - 1 - hint.leading_zeros()));
            }
        }
        let mut state = Self {
            grammar,
            builder,
            source,
            offset: 0,
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
            farthest: 0,
            expected: Vec::new(),
            expected_unique_len: 0,
        };
        while let Some(task) = state.tasks.pop() {
            if state.steps >= options.max_steps {
                return Err(state.error(
                    ParseErrorKind::WorkLimit,
                    "source parse work limit exceeded",
                ));
            }
            state.steps += 1;
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
        Ok(state)
    }
    fn checkpoint(&self) -> Checkpoint<E> {
        Checkpoint {
            offset: self.offset,
            undo: E::undo_mark(self.undo.len()),
            mark: if CAPTURE {
                self.nodes.len()
            } else {
                self.builder.checkpoint()
            },
        }
    }
    fn try_fast_trivia_failure(
        &mut self,
        env: E,
        mut whitespace_done: bool,
        continuation_steps: usize,
    ) -> bool {
        // Capture parsing always uses NoBuilder. Other modes retain builder
        // checkpoint/rollback callbacks, even when trivia makes no changes.
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
        self.steps += cost;
        self.matched = true;
        true
    }
    fn restore(&mut self, checkpoint: Checkpoint<E>) {
        if CAPTURE {
            self.nodes.truncate(checkpoint.mark);
        } else {
            self.builder.rollback(checkpoint.mark);
        }
        self.offset = checkpoint.offset;
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
        if env.predicate() || env.skipping() || self.offset < self.farthest {
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
        }
        Ok(())
    }
    fn pratt_finish(&mut self, mut frame: PrattFrame<E>) -> Result<(), ParseError> {
        while let Some(operator) = frame.operators.pop() {
            self.pratt_reduce(frame.env, operator)?;
        }
        self.matched = true;
        Ok(())
    }
    fn pratt_following(&mut self, frame: Box<PrattFrame<E>>) {
        let checkpoint = self.checkpoint();
        let env = frame.env;
        self.tasks.push(Task::PrattFollowing {
            frame: TaskMode::<CAPTURE>::wrap(frame),
            candidate: 0,
            checkpoint,
        });
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
        let candidates = if env.predicate() || env.skipping() || self.offset < self.farthest {
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
                self.steps += alternative.success_steps - 1;
                self.offset += length;
                self.matched = true;
                return true;
            }
            self.failure(env, expected);
        }
        self.steps += choice.failure_steps - 1;
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
    fn charge_step(&mut self) -> Result<(), ParseError> {
        if self.steps >= self.options.max_steps {
            return Err(self.error(
                ParseErrorKind::WorkLimit,
                "source parse work limit exceeded",
            ));
        }
        self.steps += 1;
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
    ) -> Result<(), ParseError> {
        self.rule_depth -= 1;
        if !self.matched {
            self.restore(checkpoint);
        } else if CAPTURE && let Some(index) = completion {
            self.nodes.finish(index, self.offset);
        }
        if !CAPTURE
            && self.matched
            && let Some(rule) = completion
        {
            self.builder.try_finish(
                rule,
                Span {
                    start: checkpoint.offset,
                    end: self.offset,
                },
                self.source,
            )?;
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
        self.steps += scanned * 4;
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
    fn scan_fragment(
        &self,
        fragment: &crate::program::AtomicFragment<'g>,
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
                            self.steps += prefix.steps + 1;
                            self.matched = false;
                            break;
                        }
                        if self.rule_depth >= self.options.max_rule_depth {
                            return Err(self.error(
                                ParseErrorKind::DepthLimit,
                                "source rule nesting limit exceeded",
                            ));
                        }
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
                        if let Some(prefix) = self.grammar.leading_bytes(repetition.child) {
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
                            self.steps += count * cost;
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
    fn execute(&mut self, mut task: ExecutionTask<E, CAPTURE>) -> Result<(), ParseError> {
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
            match task {
                Task::PrattOperand { frame, candidate } => {
                    let frame = TaskMode::<CAPTURE>::unwrap(frame);
                    let spec = self.grammar.source_pratt(frame.target).unwrap();
                    if let Some((candidate, operator)) = spec
                        .operators
                        .iter()
                        .enumerate()
                        .skip(candidate)
                        .find(|(_, op)| op.fixity == crate::pratt::Fixity::Prefix)
                    {
                        let operator = *operator;
                        let env = frame.env.without_semantic();
                        self.tasks.push(Task::PrattAfterPrefix {
                            frame: TaskMode::<CAPTURE>::wrap(frame),
                            candidate,
                            checkpoint: self.checkpoint(),
                        });
                        self.tasks.push(Task::EnterRule(operator.rule, env));
                    } else {
                        let atom = spec.atom;
                        let env = frame.env;
                        self.tasks
                            .push(Task::PrattAfterAtom(TaskMode::<CAPTURE>::wrap(frame)));
                        self.tasks.push(Task::EnterRule(atom, env));
                    }
                }
                Task::PrattAfterPrefix {
                    frame,
                    candidate,
                    checkpoint,
                } => {
                    let mut frame = TaskMode::<CAPTURE>::unwrap(frame);
                    if self.matched {
                        if self.offset == checkpoint.offset {
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
                                    start: checkpoint.offset,
                                    end: self.offset,
                                },
                            },
                        )?;
                        let env = frame.env;
                        self.tasks.push(Task::PrattOperand {
                            frame: TaskMode::<CAPTURE>::wrap(frame),
                            candidate: 0,
                        });
                        if env.allows_skip() {
                            self.tasks.push(Task::Skip(env));
                        }
                    } else {
                        self.tasks.push(Task::PrattOperand {
                            frame: TaskMode::<CAPTURE>::wrap(frame),
                            candidate: candidate + 1,
                        });
                    }
                }
                Task::PrattAfterAtom(frame) => {
                    let mut frame = TaskMode::<CAPTURE>::unwrap(frame);
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
                        self.tasks.push(Task::PrattFollowing {
                            frame: TaskMode::<CAPTURE>::wrap(frame),
                            candidate: candidate + 1,
                            checkpoint,
                        });
                        if env.allows_skip() {
                            self.tasks.push(Task::Skip(env));
                        }
                    }
                    // Initial operand failure leaves matched=false. FinishRule
                    // restores the expression's complete transactional checkpoint.
                }
                Task::PrattFollowing {
                    frame,
                    candidate,
                    mut checkpoint,
                } => {
                    let frame = TaskMode::<CAPTURE>::unwrap(frame);
                    let spec = self.grammar.source_pratt(frame.target).unwrap();
                    if candidate == 0
                        && frame.env.semantic()
                        && !frame.env.predicate()
                        && !frame.env.skipping()
                    {
                        self.builder.try_pratt_boundary(frame.target, self.offset)?;
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
                        .find(|(_, op)| op.fixity != crate::pratt::Fixity::Prefix)
                    {
                        let operator = *operator;
                        let env = frame.env.without_semantic();
                        self.tasks
                            .push(Task::PrattAfterFollowing(TaskMode::<CAPTURE>::following((
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
                Task::PrattAfterFollowing(value) => {
                    let (mut frame, candidate, checkpoint, start) =
                        TaskMode::<CAPTURE>::following_value(value);
                    if !self.matched {
                        self.tasks.push(Task::PrattFollowing {
                            frame: TaskMode::<CAPTURE>::wrap(frame),
                            candidate: candidate + 1,
                            checkpoint,
                        });
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
                            self.tasks.push(Task::PrattOperand {
                                frame: TaskMode::<CAPTURE>::wrap(frame),
                                candidate: 0,
                            });
                            if env.allows_skip() {
                                self.tasks.push(Task::Skip(env));
                            }
                        }
                    }
                }
                Task::RunBuiltin(builtin, env) => self.builtin(builtin, env)?,
                Task::Eval(id, env) => {
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
                        self.steps += cost - 1; // root Eval already charged
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
                    ) {
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
                            if self.grammar.tail_dispatch() {
                                self.charge_step()?;
                                task = Task::Eval(left, env);
                                continue;
                            }
                            self.tasks.push(Task::Eval(left, env));
                        }
                        Instruction::Choice { left, right } => {
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
                                checkpoint: self.checkpoint(),
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
                    self.grammar.observe_rule_attempt(
                        rule,
                        self.offset,
                        incoming.profile_flags(CAPTURE),
                    );
                    let mut owned_prefix = None;
                    let general_failure = CAPTURE && self.grammar.fast_rule_failures();
                    let diagnostics = !incoming.skipping()
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
                        self.steps += prefix.steps + 1; // expression + FinishRule
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
                    let builder_rule =
                        (!incoming.predicate() && !incoming.skipping() && incoming.semantic())
                            .then_some(rule);
                    if builder_rule.is_some() {
                        self.builder.try_begin(rule, self.offset)?;
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
                    let pratt = !CAPTURE
                        && incoming.semantic()
                        && !incoming.skipping()
                        && self.builder.supports_pratt(rule)
                        && self.grammar.source_pratt(rule).is_some();
                    self.tasks.push(Task::FinishRule {
                        checkpoint,
                        completion: TaskMode::<CAPTURE>::completion(if CAPTURE {
                            capture
                        } else {
                            builder_rule
                        }),
                    });
                    if CAPTURE
                        && env.mode() == Atomicity::Atomic
                        && self.grammar.atomic_expression(definition.expression)
                    {
                        if FRAGMENT
                            && let Some(fragment) = self.grammar.atomic_fragment(rule)
                            && let Some(outcome) = self.scan_fragment(fragment)
                        {
                            if let FragmentOutcome::Rejected(cost, expected) = outcome {
                                if let Some(guard) = fragment.guard {
                                    self.grammar.observe_rule_attempt(
                                        guard.rule,
                                        self.offset,
                                        env.with_predicate(true).profile_flags(CAPTURE),
                                    );
                                }
                                self.steps += cost;
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
                            self.steps += cost;
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
                        self.tasks.push(Task::PrattOperand {
                            frame: TaskMode::<CAPTURE>::wrap(Box::new(PrattFrame {
                                target: rule,
                                env,
                                operators: Vec::new(),
                                undo: Vec::new(),
                                attempt: None,
                            })),
                            candidate: 0,
                        });
                    } else {
                        if self.grammar.tail_dispatch() {
                            self.charge_step()?;
                            task = Task::Eval(definition.expression, env);
                            continue;
                        }
                        self.tasks.push(Task::Eval(definition.expression, env));
                    }
                }
                Task::FinishRule {
                    checkpoint,
                    completion,
                } => {
                    self.finish_rule(
                        checkpoint,
                        TaskMode::<CAPTURE>::completion_value(completion),
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
                        if self.offset == checkpoint.offset && repetition.max.is_none() {
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
                    self.restore(checkpoint);
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
                            self.steps += 1;
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
                            let checkpoint = self.checkpoint();
                            self.tasks.push(Task::AfterWhitespace { checkpoint, env });
                            self.tasks.push(Task::EnterRule(
                                self.grammar.scope(env.scope()).whitespace.unwrap(),
                                env.with_skipping(),
                            ));
                            return Ok(());
                        }
                        let checkpoint = self.checkpoint();
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
                    let checkpoint = self.checkpoint();
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
                        if self.offset == checkpoint.offset {
                            return Err(self.error(
                                ParseErrorKind::NonProgress,
                                "WHITESPACE succeeded without consuming input",
                            ));
                        }
                        self.tasks.push(Task::Skip(env));
                    } else {
                        self.restore(checkpoint);
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
                        if self.offset == checkpoint.offset {
                            return Err(self.error(
                                ParseErrorKind::NonProgress,
                                "COMMENT succeeded without consuming input",
                            ));
                        }
                        self.tasks.push(Task::Skip(env));
                    } else {
                        self.restore(checkpoint);
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
                    if CAPTURE && !env.predicate() && env.mode() != Atomicity::Atomic {
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
        eprintln!(
            "Task={} RepeatFrame={} Repetition={} Checkpoint={}",
            std::mem::size_of::<
                super::Task<
                    std::convert::Infallible,
                    super::Env,
                    Option<usize>,
                    std::convert::Infallible,
                >,
            >(),
            std::mem::size_of::<super::RepeatFrame<super::Env>>(),
            std::mem::size_of::<super::Repetition<super::Env>>(),
            std::mem::size_of::<super::Checkpoint>()
        );
        assert_eq!(
            std::mem::size_of::<
                super::Task<
                    std::convert::Infallible,
                    super::Env,
                    Option<usize>,
                    std::convert::Infallible,
                >,
            >(),
            48
        );
        assert_eq!(
            std::mem::size_of::<
                super::Task<
                    Box<super::PrattFrame<super::Env>>,
                    super::Env,
                    Option<usize>,
                    (
                        Box<super::PrattFrame<super::Env>>,
                        usize,
                        super::Checkpoint,
                        usize
                    ),
                >,
            >(),
            56
        );
    }
}

#[cfg(test)]
mod compact_layout_tests {
    use super::*;
    type WideCaptureTask = Task<
        <TaskMode<true> as PrattTaskMode<Env>>::Frame,
        Env,
        <TaskMode<true> as PrattTaskMode<Env>>::Completion,
        <TaskMode<true> as PrattTaskMode<Env>>::AfterFollowing,
    >;
    type CompactCaptureTask = Task<
        <TaskMode<true> as PrattTaskMode<CompactEnv>>::Frame,
        CompactEnv,
        <TaskMode<true> as PrattTaskMode<CompactEnv>>::Completion,
        <TaskMode<true> as PrattTaskMode<CompactEnv>>::AfterFollowing,
    >;
    type BuilderTask = Task<
        <TaskMode<false> as PrattTaskMode<Env>>::Frame,
        Env,
        <TaskMode<false> as PrattTaskMode<Env>>::Completion,
        <TaskMode<false> as PrattTaskMode<Env>>::AfterFollowing,
    >;
    #[test]
    #[cfg(target_pointer_width = "64")]
    fn sizes() {
        assert_eq!(std::mem::size_of::<Env>(), 16);
        assert_eq!(std::mem::size_of::<CompactEnv>(), 8);
        assert_eq!(std::mem::size_of::<WideCaptureTask>(), 48);
        assert_eq!(std::mem::size_of::<CompactCaptureTask>(), 40);
        assert_eq!(std::mem::size_of::<BuilderTask>(), 56);
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
        <TaskMode<true> as PrattTaskMode<StackFreeEnv>>::AfterFollowing,
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
