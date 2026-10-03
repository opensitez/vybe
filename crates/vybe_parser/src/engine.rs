//! Iterative scannerless PEG reference engine. No host/Rust recursion is used
//! for source nesting. Generated recognition/direct AST backends follow later.
use crate::builder::{Builder, NoBuilder};
use crate::grammar::{ExprId, RuleId};
use crate::program::{Instruction, Program};
use crate::tree::{CaptureRule, Node, ParseTree};
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
    let state = State::<P, NoBuilder, true>::run(program, source, entry, options, NoBuilder)?;
    Ok(ParseTree {
        grammar: program,
        source,
        nodes: state.nodes,
        consumed: state.offset,
    })
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

#[derive(Clone, Copy, PartialEq, Eq)]
enum Atomicity {
    Normal,
    Atomic,
    Compound,
}
#[derive(Clone, Copy)]
struct Env {
    semantic: bool,
    scope: usize,
    mode: Atomicity,
    predicate: bool,
    skipping: bool,
}
impl Env {
    fn allows_skip(self) -> bool {
        self.mode == Atomicity::Normal && !self.skipping
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
struct Checkpoint {
    offset: usize,
    nodes: usize,
    undo: usize,
    builder: usize,
}
#[derive(Clone, Copy)]
struct Repetition {
    child: ExprId,
    env: Env,
    min: u32,
    max: Option<u32>,
    count: u32,
}

#[derive(Clone, Copy)]
struct PendingOperator { operator: crate::program::BoundOperator, span: Span }
enum OperatorUndo { Remove, Restore(PendingOperator) }
struct PrattFrame {
    target: RuleId,
    env: Env,
    operators: Vec<PendingOperator>,
    undo: Vec<OperatorUndo>,
    attempt: Option<Checkpoint>,
}

enum Task {
    PrattOperand { frame: Box<PrattFrame>, candidate: usize },
    PrattAfterPrefix { frame: Box<PrattFrame>, candidate: usize, checkpoint: Checkpoint },
    PrattAfterAtom(Box<PrattFrame>),
    PrattFollowing { frame: Box<PrattFrame>, candidate: usize, checkpoint: Checkpoint },
    PrattAfterFollowing { frame: Box<PrattFrame>, candidate: usize, checkpoint: Checkpoint, start: usize },
    RunBuiltin(Builtin, Env),
    Eval(ExprId, Env),
    FinishExpr(Checkpoint),
    EnterRule(RuleId, Env),
    FinishRule {
        checkpoint: Checkpoint,
        capture: Option<usize>,
        builder_rule: Option<RuleId>,
    },
    Sequence {
        right: ExprId,
        env: Env,
    },
    AfterSkip {
        right: ExprId,
        env: Env,
    },
    Choice {
        right: ExprId,
        env: Env,
    },
    Repeat(Repetition),
    RepeatAfterSkip {
        repetition: Repetition,
        checkpoint: Checkpoint,
    },
    RepeatAfterChild {
        repetition: Repetition,
        checkpoint: Checkpoint,
    },
    Predicate {
        checkpoint: Checkpoint,
        positive: bool,
    },
    Push {
        start: usize,
    },
    Skip(Env),
    AfterWhitespace {
        checkpoint: Checkpoint,
        env: Env,
    },
    AfterComment {
        checkpoint: Checkpoint,
        env: Env,
    },
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

struct State<'g, 'i, P: Program, B: Builder, const CAPTURE: bool> {
    grammar: &'g P,
    builder: B,
    source: &'i str,
    offset: usize,
    nodes: Vec<Node>,
    stack: Vec<StackValue>,
    undo: Vec<Undo>,
    tasks: Vec<Task>,
    matched: bool,
    rule_depth: usize,
    steps: usize,
    options: MatchOptions,
    farthest: usize,
    expected: Vec<Expected<'g>>,
}

impl<'g, 'i, P: Program, B: Builder, const CAPTURE: bool> State<'g, 'i, P, B, CAPTURE> {
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
        let env = Env {
            semantic: true,
            scope: 0,
            mode: Atomicity::Normal,
            predicate: false,
            skipping: false,
        };
        let mut state = Self {
            grammar,
            builder,
            source,
            offset: 0,
            nodes: Vec::new(),
            stack: Vec::new(),
            undo: Vec::new(),
            tasks: vec![match entry {
                Reference::Rule(rule) => Task::EnterRule(rule, env),
                Reference::Builtin(builtin) => Task::RunBuiltin(builtin, env),
            }],
            matched: true,
            rule_depth: 0,
            steps: 0,
            options,
            farthest: 0,
            expected: Vec::new(),
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
            return Err(ParseError {
                kind: ParseErrorKind::Syntax,
                offset: state.farthest,
                expected: state.expected.into_iter().map(Expected::render).collect(),
                message: "source does not match the grammar".into(),
            });
        }
        Ok(state)
    }
    fn checkpoint(&self) -> Checkpoint {
        Checkpoint {
            offset: self.offset,
            nodes: self.nodes.len(),
            undo: self.undo.len(),
            builder: self.builder.checkpoint(),
        }
    }
    fn restore(&mut self, checkpoint: Checkpoint) {
        self.builder.rollback(checkpoint.builder);
        self.offset = checkpoint.offset;
        self.nodes.truncate(checkpoint.nodes);
        while self.undo.len() > checkpoint.undo {
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
    fn failure(&mut self, env: Env, expected: Expected<'g>) {
        self.matched = false;
        if env.predicate || env.skipping || self.offset < self.farthest {
            return;
        }
        if self.offset > self.farthest {
            self.farthest = self.offset;
            self.expected.clear();
        }
        if !self.expected.contains(&expected) {
            self.expected.push(expected);
        }
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
    fn pratt_push(&self, frame: &mut PrattFrame, operator: PendingOperator) -> Result<(), ParseError> {
        if frame.operators.len() >= self.options.max_rule_depth {
            return Err(self.error(ParseErrorKind::DepthLimit, "source expression operator nesting limit exceeded"));
        }
        frame.operators.push(operator);
        frame.undo.push(OperatorUndo::Remove);
        Ok(())
    }
    fn pratt_reduce(&mut self, env: Env, operator: PendingOperator) {
        if env.semantic && !env.predicate && !env.skipping {
            self.builder.reduce(operator.operator.rule, operator.operator.fixity, operator.span);
        }
    }
    fn pratt_finish(&mut self, mut frame: Box<PrattFrame>) {
        while let Some(operator) = frame.operators.pop() { self.pratt_reduce(frame.env, operator); }
        self.matched = true;
    }
    fn pratt_following(&mut self, frame: Box<PrattFrame>) {
        let checkpoint = self.checkpoint();
        let env = frame.env;
        self.tasks.push(Task::PrattFollowing { frame, candidate: 0, checkpoint });
        if env.allows_skip() { self.tasks.push(Task::Skip(env)); }
    }
    fn literal(&mut self, text: &'g str, insensitive: bool, env: Env) {
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
    fn execute(&mut self, task: Task) -> Result<(), ParseError> {
        match task {
            Task::PrattOperand { frame, candidate } => {
                let spec = self.grammar.source_pratt(frame.target).unwrap();
                if let Some((candidate, operator)) = spec.operators.iter().enumerate().skip(candidate).find(|(_, op)| op.fixity == crate::pratt::Fixity::Prefix) {
                    let operator = *operator;
                    let env = Env { semantic: false, ..frame.env };
                    self.tasks.push(Task::PrattAfterPrefix { frame, candidate, checkpoint: self.checkpoint() });
                    self.tasks.push(Task::EnterRule(operator.rule, env));
                } else {
                    let atom = spec.atom;
                    let env = frame.env;
                    self.tasks.push(Task::PrattAfterAtom(frame));
                    self.tasks.push(Task::EnterRule(atom, env));
                }
            }
            Task::PrattAfterPrefix { mut frame, candidate, checkpoint } => {
                if self.matched {
                    if self.offset == checkpoint.offset { return Err(self.error(ParseErrorKind::NonProgress, "Pratt prefix succeeded without consuming input")); }
                    let operator = self.grammar.source_pratt(frame.target).unwrap().operators[candidate];
                    self.pratt_push(&mut frame, PendingOperator { operator, span: Span { start: checkpoint.offset, end: self.offset } })?;
                    let env = frame.env;
                    self.tasks.push(Task::PrattOperand { frame, candidate: 0 });
                    if env.allows_skip() { self.tasks.push(Task::Skip(env)); }
                } else { self.tasks.push(Task::PrattOperand { frame, candidate: candidate + 1 }); }
            }
            Task::PrattAfterAtom(mut frame) => {
                if self.matched {
                    frame.attempt = None;
                    frame.undo.clear();
                    self.pratt_following(frame);
                } else if let Some(checkpoint) = frame.attempt.take() {
                    // A grammar repetition rolls back a consumed infix whose
                    // RHS failed. Restore popped operator metadata as well as
                    // builder values/input; eager reductions may have happened.
                    self.restore(checkpoint);
                    while let Some(undo) = frame.undo.pop() {
                        match undo { OperatorUndo::Remove => { frame.operators.pop().unwrap(); }, OperatorUndo::Restore(op) => frame.operators.push(op) }
                    }
                    self.pratt_finish(frame);
                }
                // Initial operand failure leaves matched=false. FinishRule
                // restores the expression's complete transactional checkpoint.
            }
            Task::PrattFollowing { frame, candidate, checkpoint } => {
                let spec = self.grammar.source_pratt(frame.target).unwrap();
                if let Some((candidate, operator)) = spec.operators.iter().enumerate().skip(candidate).find(|(_, op)| op.fixity != crate::pratt::Fixity::Prefix) {
                    let operator = *operator;
                    let env = Env { semantic: false, ..frame.env };
                    self.tasks.push(Task::PrattAfterFollowing { frame, candidate, checkpoint, start: self.offset });
                    self.tasks.push(Task::EnterRule(operator.rule, env));
                } else {
                    self.restore(checkpoint); // keep trailing trivia for caller
                    self.pratt_finish(frame);
                }
            }
            Task::PrattAfterFollowing { mut frame, candidate, checkpoint, start } => {
                if !self.matched {
                    self.tasks.push(Task::PrattFollowing { frame, candidate: candidate + 1, checkpoint });
                } else {
                    if self.offset == start { return Err(self.error(ParseErrorKind::NonProgress, "Pratt operator succeeded without consuming input")); }
                    let operator = self.grammar.source_pratt(frame.target).unwrap().operators[candidate];
                    while frame.operators.last().is_some_and(|top| top.operator.precedence > operator.precedence || (top.operator.precedence == operator.precedence && operator.fixity != crate::pratt::Fixity::InfixRight)) {
                        let top = frame.operators.pop().unwrap();
                        frame.undo.push(OperatorUndo::Restore(top));
                        self.pratt_reduce(frame.env, top);
                    }
                    let pending = PendingOperator { operator, span: Span { start, end: self.offset } };
                    if operator.fixity == crate::pratt::Fixity::Postfix {
                        self.pratt_reduce(frame.env, pending);
                        frame.undo.clear();
                        self.pratt_following(frame);
                    } else {
                        self.pratt_push(&mut frame, pending)?;
                        frame.attempt = Some(checkpoint);
                        let env = frame.env;
                        self.tasks.push(Task::PrattOperand { frame, candidate: 0 });
                        if env.allows_skip() { self.tasks.push(Task::Skip(env)); }
                    }
                }
            }
            Task::RunBuiltin(builtin, env) => self.builtin(builtin, env)?,
            Task::Eval(id, env) => {
                let instruction = self.grammar.instruction(id);
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
                    Instruction::Call(rule) => self.tasks.push(Task::EnterRule(rule, env)),
                    Instruction::Builtin(builtin) => self.builtin(builtin, env)?,
                    Instruction::Sequence { left, right } => {
                        self.tasks.push(Task::Sequence { right, env });
                        self.tasks.push(Task::Eval(left, env));
                    }
                    Instruction::Choice { left, right } => {
                        self.tasks.push(Task::Choice { right, env });
                        self.tasks.push(Task::Eval(left, env));
                    }
                    Instruction::Repeat {
                        expression,
                        min,
                        max,
                    } => self.tasks.push(Task::Repeat(Repetition {
                        child: expression,
                        env,
                        min,
                        max,
                        count: 0,
                    })),
                    Instruction::Predicate {
                        expression,
                        positive,
                    } => {
                        self.tasks.push(Task::Predicate {
                            checkpoint: self.checkpoint(),
                            positive,
                        });
                        self.tasks.push(Task::Eval(
                            expression,
                            Env {
                                predicate: true,
                                ..env
                            },
                        ));
                    }
                    Instruction::Group(child) => self.tasks.push(Task::Eval(child, env)),
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
                    _ => incoming.mode,
                };
                let capture_mode = if matches!(
                    definition.mode,
                    RuleMode::CompoundAtomic | RuleMode::NonAtomic
                ) {
                    mode
                } else {
                    incoming.mode
                };
                let capture = if CAPTURE
                    && !incoming.predicate
                    && definition.mode != RuleMode::Silent
                    && capture_mode != Atomicity::Atomic
                {
                    let index = self.nodes.len();
                    self.nodes.push(Node {
                        rule: CaptureRule::User(rule),
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
                let builder_rule = (!incoming.predicate && !incoming.skipping && incoming.semantic).then_some(rule);
                if builder_rule.is_some() {
                    self.builder.begin(rule, self.offset);
                }
                self.tasks.push(Task::FinishRule {
                    checkpoint,
                    capture,
                    builder_rule,
                });
                let env = Env {
                        mode,
                        scope: definition.scope,
                        ..incoming
                    };
                if !CAPTURE && incoming.semantic && self.builder.supports_pratt(rule) && self.grammar.source_pratt(rule).is_some() {
                    self.tasks.push(Task::PrattOperand { frame: Box::new(PrattFrame { target: rule, env, operators: Vec::new(), undo: Vec::new(), attempt: None }), candidate: 0 });
                } else { self.tasks.push(Task::Eval(definition.expression, env)); }
            }
            Task::FinishRule {
                checkpoint,
                capture,
                builder_rule,
            } => {
                self.rule_depth -= 1;
                if !self.matched {
                    self.restore(checkpoint);
                } else if let Some(index) = capture {
                    self.nodes[index].span.end = self.offset;
                    self.nodes[index].subtree_end = self.nodes.len();
                }
                if self.matched
                    && let Some(rule) = builder_rule
                {
                    self.builder.finish(
                        rule,
                        Span {
                            start: checkpoint.offset,
                            end: self.offset,
                        },
                        self.source,
                    );
                }
            }
            Task::Sequence { right, env } => {
                if self.matched {
                    if env.allows_skip() {
                        self.tasks.push(Task::AfterSkip { right, env });
                        self.tasks.push(Task::Skip(env));
                    } else {
                        self.tasks.push(Task::Eval(right, env));
                    }
                }
            }
            Task::AfterSkip { right, env } => {
                if self.matched {
                    self.tasks.push(Task::Eval(right, env));
                }
            }
            Task::Choice { right, env } => {
                if !self.matched {
                    self.tasks.push(Task::Eval(right, env));
                }
            }
            Task::Repeat(repetition) => {
                if repetition.max.is_some_and(|max| repetition.count >= max) {
                    self.matched = true;
                } else {
                    let checkpoint = self.checkpoint();
                    self.tasks.push(Task::RepeatAfterSkip {
                        repetition,
                        checkpoint,
                    });
                    if repetition.count > 0 && repetition.env.allows_skip() {
                        self.tasks.push(Task::Skip(repetition.env));
                    } else {
                        self.matched = true;
                    }
                }
            }
            Task::RepeatAfterSkip {
                repetition,
                checkpoint,
            } => {
                if self.matched {
                    self.tasks.push(Task::RepeatAfterChild {
                        repetition,
                        checkpoint,
                    });
                    self.tasks
                        .push(Task::Eval(repetition.child, repetition.env));
                }
            }
            Task::RepeatAfterChild {
                mut repetition,
                checkpoint,
            } => {
                if !self.matched {
                    self.restore(checkpoint);
                    self.matched = repetition.count >= repetition.min;
                } else {
                    if self.offset == checkpoint.offset && repetition.max.is_none() {
                        return Err(self.error(
                            ParseErrorKind::NonProgress,
                            "unbounded repetition succeeded without consuming input",
                        ));
                    }
                    repetition.count += 1;
                    self.tasks.push(Task::Repeat(repetition));
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
                if let Some(class) = self.grammar.scope(env.scope).whitespace_prefix_class {
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
                    if self.grammar.scope(env.scope).whitespace_class.is_none() {
                        // Prefix scan consumed only successful leading terminal
                        // alternatives. Evaluate the unchanged full rule at
                        // the first miss, retaining all fallback effects.
                        let checkpoint = self.checkpoint();
                        self.tasks.push(Task::AfterWhitespace { checkpoint, env });
                        self.tasks.push(Task::EnterRule(
                            self.grammar.scope(env.scope).whitespace.unwrap(),
                            Env {
                                skipping: true,
                                ..env
                            },
                        ));
                        return Ok(());
                    }
                    let checkpoint = self.checkpoint();
                    self.tasks.push(Task::AfterComment { checkpoint, env });
                    if let Some(rule) = self.grammar.scope(env.scope).comment {
                        self.tasks.push(Task::EnterRule(
                            rule,
                            Env {
                                skipping: true,
                                ..env
                            },
                        ));
                    } else {
                        self.matched = false;
                    }
                    return Ok(());
                }
                let checkpoint = self.checkpoint();
                let skip_env = Env {
                    skipping: true,
                    ..env
                };
                self.tasks.push(Task::AfterWhitespace { checkpoint, env });
                if let Some(rule) = self.grammar.scope(env.scope).whitespace {
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
                    if let Some(rule) = self.grammar.scope(env.scope).comment {
                        self.tasks.push(Task::EnterRule(
                            rule,
                            Env {
                                skipping: true,
                                ..env
                            },
                        ));
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
        Ok(())
    }

    fn builtin(&mut self, builtin: Builtin, env: Env) -> Result<(), ParseError> {
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
                    if CAPTURE && !env.predicate && env.mode != Atomicity::Atomic {
                        let index = self.nodes.len();
                        self.nodes.push(Node {
                            rule: CaptureRule::EndOfInput,
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
    fn match_stack_value(&mut self, value: StackValue, env: Env) {
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
        env: Env,
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
