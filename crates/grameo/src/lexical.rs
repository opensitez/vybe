//! Conservative lexical certificates. Unknown, stateful, nullable or capturing
//! constructions stay on the full engine; no speculative heuristic is used.
use crate::{
    Builtin, CompiledGrammar, RuleMode,
    program::{Instruction, Program},
};

pub(crate) fn ascii_predicate(builtin: Builtin) -> Option<fn(u8) -> bool> {
    Some(match builtin {
        Builtin::Ascii => |b| b < 128,
        Builtin::AsciiDigit => |b| b.is_ascii_digit(),
        Builtin::AsciiAlpha => |b| b.is_ascii_alphabetic(),
        Builtin::AsciiAlphaLower => |b| b.is_ascii_lowercase(),
        Builtin::AsciiAlphaUpper => |b| b.is_ascii_uppercase(),
        Builtin::AsciiAlphanumeric => |b| b.is_ascii_alphanumeric(),
        Builtin::AsciiHexDigit => |b| b.is_ascii_hexdigit(),
        Builtin::AsciiNonzeroDigit => |b| (b'1'..=b'9').contains(&b),
        Builtin::AsciiBinDigit => |b| matches!(b, b'0' | b'1'),
        Builtin::AsciiOctDigit => |b| (b'0'..=b'7').contains(&b),
        _ => return None,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ByteClass(pub [u64; 2]);
impl ByteClass {
    pub fn contains(self, byte: u8) -> bool {
        byte < 128 && self.0[usize::from(byte / 64)] & (1u64 << (byte % 64)) != 0
    }
    fn insert(&mut self, byte: u8) {
        self.0[usize::from(byte / 64)] |= 1u64 << (byte % 64);
    }
}

/// A necessary first-byte set with exact cost/depth for failure on a miss.
/// Trivia and capture-mode lookahead use this certificate: their failures emit
/// no diagnostics or semantic hooks. Possible matches execute the grammar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FailurePrefix {
    pub bytes: [u64; 4],
    /// Number of expression tasks on the certified failing path.
    pub steps: usize,
    /// Maximum additional rule depth inside the expression.
    pub depth: usize,
}
impl FailurePrefix {
    pub fn contains(self, byte: u8) -> bool {
        self.bytes[usize::from(byte / 64)] & (1u64 << (byte % 64)) != 0
    }
    fn insert(&mut self, byte: u8) {
        self.bytes[usize::from(byte / 64)] |= 1u64 << (byte % 64);
    }
    fn add(mut self, steps: usize, depth: usize) -> Option<Self> {
        self.steps = self.steps.checked_add(steps)?;
        self.depth = self.depth.checked_add(depth)?;
        (self.steps <= 131_072).then_some(self)
    }
}

pub(crate) fn trivia_failure_prefixes(grammar: &CompiledGrammar) -> Vec<Option<FailurePrefix>> {
    let count = grammar.syntax().expressions.len();
    let mut expressions = vec![None; count];
    let mut status = vec![0u8; count];
    let mut needed = vec![true; grammar.syntax().rules.len()];
    let mut tasks: Vec<_> = grammar
        .syntax()
        .rules
        .iter()
        .map(|rule| (rule.expression, false))
        .collect();
    for scope in grammar.scopes() {
        for rule in scope.whitespace.into_iter().chain(scope.comment) {
            needed[rule] = true;
            tasks.push((grammar.rule(rule).expression, false));
        }
    }
    // A predicate's directly reachable rule calls can use the same exact
    // failure proof. Follow only the paths supported by the conservative
    // analysis below; unknown/effectful prefixes stay on the ordinary engine.
    for id in 0..count {
        if let Instruction::Predicate { expression, .. } = grammar.instruction(id) {
            tasks.push((expression, false));
        }
    }
    while let Some((id, finish)) = tasks.pop() {
        if !finish {
            if status[id] != 0 {
                continue;
            }
            status[id] = 1;
            tasks.push((id, true));
            match grammar.instruction(id) {
                Instruction::Call(rule) => {
                    needed[rule] = true;
                    tasks.push((grammar.rule(rule).expression, false));
                }
                Instruction::Sequence { left, .. } => tasks.push((left, false)),
                Instruction::Choice { left, right } => {
                    tasks.push((right, false));
                    tasks.push((left, false));
                }
                Instruction::Group(child) | Instruction::Push(child) => tasks.push((child, false)),
                Instruction::Repeat {
                    expression, min, ..
                } if min > 0 => tasks.push((expression, false)),
                _ => {}
            }
            continue;
        }
        let mut terminal = FailurePrefix {
            bytes: [0; 4],
            steps: 1,
            depth: 0,
        };
        expressions[id] = match grammar.instruction(id) {
            Instruction::Literal { text, insensitive } if !text.is_empty() => {
                let byte = text.as_bytes()[0];
                terminal.insert(byte);
                if insensitive {
                    terminal.insert(byte.to_ascii_lowercase());
                    terminal.insert(byte.to_ascii_uppercase());
                }
                Some(terminal)
            }
            Instruction::Range { start, end } => {
                for byte in 0u8..128 {
                    if (start..=end).contains(&char::from(byte)) {
                        terminal.insert(byte);
                    }
                }
                // A conservative superset for all non-ASCII UTF-8 scalars.
                if end > '\u{7f}' {
                    for byte in 0xc2..=0xf4 {
                        terminal.insert(byte);
                    }
                }
                Some(terminal)
            }
            Instruction::Builtin(builtin) => {
                if let Some(accepts) = ascii_predicate(builtin) {
                    for byte in 0u8..128 {
                        if accepts(byte) {
                            terminal.insert(byte);
                        }
                    }
                    Some(terminal)
                } else if builtin == Builtin::Newline {
                    terminal.insert(b'\r');
                    terminal.insert(b'\n');
                    Some(terminal)
                } else {
                    None
                }
            }
            Instruction::Sequence { left, .. } => {
                expressions[left].and_then(|p: FailurePrefix| p.add(3, 0))
            }
            Instruction::Choice { left, right } => match (expressions[left], expressions[right]) {
                (Some(a), Some(b)) => {
                    let mut prefix = a;
                    for (word, other) in prefix.bytes.iter_mut().zip(b.bytes) {
                        *word |= other;
                    }
                    prefix.depth = a.depth.max(b.depth);
                    prefix.add(b.steps + 2, 0)
                }
                _ => None,
            },
            Instruction::Call(rule) => {
                expressions[grammar.rule(rule).expression].and_then(|p| p.add(3, 1))
            }
            Instruction::Group(child) => expressions[child].and_then(|p| p.add(1, 0)),
            Instruction::Push(child) => expressions[child].and_then(|p| p.add(2, 0)),
            Instruction::Repeat {
                expression, min, ..
            } if min > 0 => expressions[expression].and_then(|p| p.add(5, 0)),
            _ => None,
        };
        status[id] = 2;
    }
    needed
        .into_iter()
        .enumerate()
        .map(|(rule, needed)| {
            needed
                .then(|| expressions[grammar.rule(rule).expression])
                .flatten()
        })
        .collect()
}

/// Ordered terminal failures on a certified first-byte miss. These have the
/// same offset and no effects; diagnostics retain first-occurrence ordering.
pub(crate) fn failure_terminals(grammar: &CompiledGrammar) -> Vec<Vec<usize>> {
    (0..grammar.syntax().rules.len())
        .map(|rule| {
            if grammar.trivia_failure_prefix(rule).is_none() {
                return Vec::new();
            }
            let mut pending = vec![grammar.rule(rule).expression];
            let mut terminals = Vec::new();
            let mut visits = 0;
            while let Some(id) = pending.pop() {
                visits += 1;
                if visits > 131_072 {
                    return Vec::new();
                }
                match grammar.instruction(id) {
                    Instruction::Literal { .. }
                    | Instruction::Range { .. }
                    | Instruction::Builtin(_) => {
                        if !terminals.contains(&id) {
                            terminals.push(id);
                        }
                    }
                    Instruction::Call(rule) => pending.push(grammar.rule(rule).expression),
                    Instruction::Sequence { left, .. } => pending.push(left),
                    Instruction::Choice { left, right } => {
                        pending.push(right);
                        pending.push(left);
                    }
                    Instruction::Group(child) | Instruction::Push(child) => pending.push(child),
                    Instruction::Repeat {
                        expression, min, ..
                    } if min > 0 => pending.push(expression),
                    _ => return Vec::new(),
                }
            }
            terminals
        })
        .collect()
}

pub(crate) fn whitespace_classes(
    grammar: &CompiledGrammar,
    whitespace: Option<usize>,
) -> (Option<ByteClass>, Option<ByteClass>) {
    let Some(id) = whitespace else {
        return (None, None);
    };
    let rule = id;
    let rule = grammar.rule(rule);
    if rule.mode != RuleMode::Silent {
        return (None, None);
    }
    let mut classes: Vec<Option<ByteClass>> =
        Vec::with_capacity(grammar.syntax().expressions.len());
    for id in 0..grammar.syntax().expressions.len() {
        let mut class = ByteClass([0, 0]);
        let result = match grammar.instruction(id) {
            Instruction::Literal { text, insensitive } if text.len() == 1 && text.is_ascii() => {
                let byte = text.as_bytes()[0];
                class.insert(byte);
                if insensitive {
                    class.insert(byte.to_ascii_lowercase());
                    class.insert(byte.to_ascii_uppercase());
                }
                Some(class)
            }
            Instruction::Range { start, end } if start.is_ascii() && end.is_ascii() => {
                for byte in start as u8..=end as u8 {
                    class.insert(byte);
                }
                Some(class)
            }
            Instruction::Choice { left, right } => match (classes[left], classes[right]) {
                (Some(a), Some(b)) => Some(ByteClass([a.0[0] | b.0[0], a.0[1] | b.0[1]])),
                _ => None,
            },
            Instruction::Group(child) => classes[child],
            Instruction::Builtin(builtin) => {
                let Some(predicate) = ascii_predicate(builtin) else {
                    classes.push(None);
                    continue;
                };
                for byte in 0..128 {
                    if predicate(byte) {
                        class.insert(byte);
                    }
                }
                Some(class)
            }
            _ => None,
        };
        classes.push(result);
    }
    // Only leading ordered alternatives are eligible. Stop at the first
    // unknown alternative: later terminals must not overtake a comment/stack
    // branch that could accept them or fail with a resource error.
    let mut pending = vec![rule.expression];
    let mut prefix = ByteClass([0, 0]);
    while let Some(id) = pending.pop() {
        if let Some(class) = classes[id] {
            prefix.0[0] |= class.0[0];
            prefix.0[1] |= class.0[1];
        } else {
            match grammar.instruction(id) {
                Instruction::Choice { left, right } => {
                    pending.push(right);
                    pending.push(left);
                }
                Instruction::Group(child) => pending.push(child),
                _ => break,
            }
        }
    }
    (
        classes[rule.expression],
        (prefix.0 != [0, 0]).then_some(prefix),
    )
}

/// A maximal choice/group containing only nonempty literals. Alternative IDs
/// retain grammar priority; successful choices always execute ordinary tasks.
#[derive(Debug, Clone)]
pub struct LiteralChoice {
    pub alternatives: Vec<usize>,
    pub steps: usize,
}

/// Exactly the engine's literal comparison, including ASCII-only folding.
#[inline]
pub fn literal_matches(source: &str, text: &str, insensitive: bool) -> bool {
    if insensitive {
        source
            .as_bytes()
            .get(..text.len())
            .is_some_and(|bytes| bytes.eq_ignore_ascii_case(text.as_bytes()))
    } else {
        source.starts_with(text)
    }
}

pub(crate) fn literal_choices(
    grammar: &CompiledGrammar,
) -> std::collections::HashMap<usize, LiteralChoice> {
    let count = grammar.syntax().expressions.len();
    // Compute costs/counts first, then flatten only maximal roots. Flattening
    // every nested choice would make long alternative lists quadratic.
    let mut proofs: Vec<Option<(usize, usize)>> = Vec::with_capacity(count);
    let mut parent = vec![false; count];
    for id in 0..count {
        let proof = match grammar.instruction(id) {
            Instruction::Literal { text, .. } if !text.is_empty() => Some((1, 1)),
            Instruction::Group(child) => proofs[child]
                .and_then(|(steps, leaves)| steps.checked_add(1).map(|steps| (steps, leaves))),
            Instruction::Choice { left, right } => match (proofs[left], proofs[right]) {
                (Some((a, x)), Some((b, y))) => a
                    .checked_add(b)
                    .and_then(|sum| sum.checked_add(2))
                    .zip(x.checked_add(y)),
                _ => None,
            },
            _ => None,
        }
        .filter(|(steps, _)| *steps <= 131_072);
        if proof.is_some() {
            match grammar.instruction(id) {
                Instruction::Group(child) => parent[child] = true,
                Instruction::Choice { left, right } => {
                    parent[left] = true;
                    parent[right] = true;
                }
                _ => {}
            }
        }
        proofs.push(proof);
    }
    let mut choices = std::collections::HashMap::new();
    for id in 0..count {
        let Some((steps, leaves)) = proofs[id] else {
            continue;
        };
        if parent[id] || leaves < 2 {
            continue;
        }
        let mut pending = vec![id];
        let mut alternatives = Vec::with_capacity(leaves);
        while let Some(child) = pending.pop() {
            match grammar.instruction(child) {
                Instruction::Literal { .. } => alternatives.push(child),
                Instruction::Group(child) => pending.push(child),
                Instruction::Choice { left, right } => {
                    pending.push(right);
                    pending.push(left);
                }
                _ => unreachable!("certified pure literal choice"),
            }
        }
        choices.insert(
            id,
            LiteralChoice {
                alternatives,
                steps,
            },
        );
    }
    choices
}

/// Pure terminal alternatives in grammar order, with exact dispatch costs.
#[derive(Debug, Clone, Copy)]
pub struct TerminalAlternative {
    pub expression: usize,
    pub success_steps: usize,
}
#[derive(Debug, Clone, Copy)]
pub struct TerminalChoice<'a> {
    pub alternatives: &'a [TerminalAlternative],
    pub failure_steps: usize,
}
#[derive(Debug, Clone)]
pub(crate) struct OwnedTerminalChoice {
    pub alternatives: Vec<TerminalAlternative>,
    pub failure_steps: usize,
}
pub(crate) fn terminal_choices(grammar: &CompiledGrammar) -> Vec<Option<OwnedTerminalChoice>> {
    let count = grammar.syntax().expressions.len();
    let mut costs: Vec<Option<usize>> = Vec::with_capacity(count);
    let mut parent = vec![false; count];
    for id in 0..count {
        let cost = match grammar.instruction(id) {
            Instruction::Literal { text, .. } if !text.is_empty() => Some(1usize),
            Instruction::Range { .. } => Some(1),
            Instruction::Builtin(builtin) if ascii_predicate(builtin).is_some() => Some(1),
            Instruction::Group(child) => costs[child].and_then(|cost: usize| cost.checked_add(1)),
            Instruction::Choice { left, right } => match (costs[left], costs[right]) {
                (Some(a), Some(b)) => a.checked_add(b).and_then(|cost| cost.checked_add(2)),
                _ => None,
            },
            _ => None,
        }
        .filter(|cost| *cost <= 131_072);
        if cost.is_some() {
            match grammar.instruction(id) {
                Instruction::Group(child) => parent[child] = true,
                Instruction::Choice { left, right } => {
                    parent[left] = true;
                    parent[right] = true;
                }
                _ => {}
            }
        }
        costs.push(cost);
    }
    (0..count)
        .map(|id| {
            let failure_steps = costs[id]?;
            if parent[id]
                || !matches!(
                    grammar.instruction(id),
                    Instruction::Choice { .. } | Instruction::Group(_)
                )
            {
                return None;
            }
            let mut pending = vec![(id, 0)];
            let mut alternatives = Vec::new();
            while let Some((id, overhead)) = pending.pop() {
                match grammar.instruction(id) {
                    Instruction::Group(child) => pending.push((child, overhead + 1)),
                    Instruction::Choice { left, right } => {
                        pending.push((right, overhead + 2 + costs[left].unwrap()));
                        pending.push((left, overhead + 2));
                    }
                    _ => alternatives.push(TerminalAlternative {
                        expression: id,
                        success_steps: overhead + 1,
                    }),
                }
            }
            if alternatives.len() < 8 {
                return None;
            }
            Some(OwnedTerminalChoice {
                alternatives,
                failure_steps,
            })
        })
        .collect()
}

/// Expressions proven to have only input/diagnostic effects in atomic mode.
/// Rule calls must retain atomic mode; recursive/stackful paths remain general.
pub(crate) fn atomic_expressions(grammar: &CompiledGrammar) -> Vec<bool> {
    let count = grammar.syntax().expressions.len();
    let mut status = vec![0u8; count];
    let mut pure = vec![false; count];
    let mut pending = Vec::new();
    for root in 0..count {
        pending.push((root, false));
        while let Some((id, finish)) = pending.pop() {
            if !finish {
                if status[id] != 0 {
                    continue;
                }
                status[id] = 1;
                pending.push((id, true));
                match grammar.instruction(id) {
                    Instruction::Sequence { left, right } | Instruction::Choice { left, right } => {
                        pending.push((right, false));
                        pending.push((left, false));
                    }
                    Instruction::Group(child)
                    | Instruction::Predicate {
                        expression: child, ..
                    }
                    | Instruction::Repeat {
                        expression: child, ..
                    } => pending.push((child, false)),
                    Instruction::Call(rule)
                        if !matches!(
                            grammar.rule(rule).mode,
                            RuleMode::NonAtomic | RuleMode::CompoundAtomic
                        ) =>
                    {
                        pending.push((grammar.rule(rule).expression, false))
                    }
                    _ => {}
                }
                continue;
            }
            pure[id] = match grammar.instruction(id) {
                Instruction::Literal { .. } | Instruction::Range { .. } => true,
                Instruction::Builtin(builtin) => !matches!(
                    builtin,
                    Builtin::Peek
                        | Builtin::PeekAll
                        | Builtin::Pop
                        | Builtin::PopAll
                        | Builtin::Drop
                ),
                Instruction::Sequence { left, right } | Instruction::Choice { left, right } => {
                    pure[left] && pure[right]
                }
                Instruction::Group(child)
                | Instruction::Predicate {
                    expression: child, ..
                }
                | Instruction::Repeat {
                    expression: child, ..
                } => pure[child],
                Instruction::Call(rule) => {
                    !matches!(
                        grammar.rule(rule).mode,
                        RuleMode::NonAtomic | RuleMode::CompoundAtomic
                    ) && pure[grammar.rule(rule).expression]
                }
                _ => false,
            };
            status[id] = 2;
        }
    }
    pure
}

/// Successful first-alternative single-ASCII-byte paths. No failed alternative
/// diagnostics or side effects precede these matches, so repeats can scan runs.
#[derive(Debug, Clone, Copy)]
pub struct LeadingBytes {
    pub class: ByteClass,
    pub steps: usize,
}
pub(crate) fn leading_bytes(grammar: &CompiledGrammar) -> Vec<Option<LeadingBytes>> {
    let mut proofs: Vec<Option<LeadingBytes>> =
        Vec::with_capacity(grammar.syntax().expressions.len());
    let mut failures: Vec<Option<LeadingBytes>> =
        Vec::with_capacity(grammar.syntax().expressions.len());
    let mut guards: Vec<Option<LeadingBytes>> =
        Vec::with_capacity(grammar.syntax().expressions.len());
    for id in 0..grammar.syntax().expressions.len() {
        let instruction = grammar.instruction(id);
        let mut class = ByteClass([0; 2]);
        let proof = match instruction {
            Instruction::Literal { text, insensitive } if text.len() == 1 && text.is_ascii() => {
                let byte = text.as_bytes()[0];
                class.insert(byte);
                if insensitive {
                    class.insert(byte.to_ascii_lowercase());
                    class.insert(byte.to_ascii_uppercase());
                }
                Some(LeadingBytes { class, steps: 1 })
            }
            Instruction::Range { start, end } => {
                for byte in 0u8..128 {
                    if (start..=end).contains(&char::from(byte)) {
                        class.insert(byte);
                    }
                }
                (class.0 != [0; 2]).then_some(LeadingBytes { class, steps: 1 })
            }
            Instruction::Builtin(builtin) => {
                for byte in 0u8..128 {
                    let matched = if let Some(accepts) = ascii_predicate(builtin) {
                        accepts(byte)
                    } else {
                        match builtin {
                            Builtin::Any => true,
                            Builtin::XidStart => unicode_ident::is_xid_start(char::from(byte)),
                            Builtin::XidContinue => {
                                unicode_ident::is_xid_continue(char::from(byte))
                            }
                            _ => false,
                        }
                    };
                    if matched {
                        class.insert(byte);
                    }
                }
                (class.0 != [0; 2]).then_some(LeadingBytes { class, steps: 1 })
            }
            Instruction::Group(child) => proofs[child].map(|p| LeadingBytes {
                steps: p.steps + 1,
                ..p
            }),
            Instruction::Choice { left, .. } => proofs[left].map(|p| LeadingBytes {
                steps: p.steps + 2,
                ..p
            }),
            _ => None,
        }
        .filter(|p| p.steps <= 131_072);
        let failure = match instruction {
            Instruction::Literal { text, insensitive } if !text.is_empty() => {
                let first = text.as_bytes()[0];
                let mut class = ByteClass([u64::MAX; 2]);
                if first < 128 {
                    class.0[usize::from(first / 64)] &= !(1u64 << (first % 64));
                    if insensitive {
                        for byte in [first.to_ascii_lowercase(), first.to_ascii_uppercase()] {
                            class.0[usize::from(byte / 64)] &= !(1u64 << (byte % 64));
                        }
                    }
                }
                Some(LeadingBytes { class, steps: 1 })
            }
            Instruction::Range { .. } | Instruction::Builtin(_) => proof.map(|p| LeadingBytes {
                class: ByteClass([!p.class.0[0], !p.class.0[1]]),
                ..p
            }),
            Instruction::Group(child) => failures[child].map(|p| LeadingBytes {
                steps: p.steps + 1,
                ..p
            }),
            Instruction::Choice { left, right } => match (failures[left], failures[right]) {
                (Some(a), Some(b)) => Some(LeadingBytes {
                    class: ByteClass([a.class.0[0] & b.class.0[0], a.class.0[1] & b.class.0[1]]),
                    steps: a.steps + b.steps + 2,
                }),
                _ => None,
            },
            _ => None,
        }
        .filter(|p| p.steps <= 131_072);
        let guard = match instruction {
            Instruction::Predicate {
                expression,
                positive: false,
            } => failures[expression].map(|p| LeadingBytes {
                steps: p.steps + 2,
                ..p
            }),
            Instruction::Group(child) => guards[child].map(|p| LeadingBytes {
                steps: p.steps + 1,
                ..p
            }),
            Instruction::Sequence { left, right } => match (guards[left], guards[right]) {
                (Some(a), Some(b)) => Some(LeadingBytes {
                    class: ByteClass([a.class.0[0] & b.class.0[0], a.class.0[1] & b.class.0[1]]),
                    steps: a.steps + b.steps + 3,
                }),
                _ => None,
            },
            _ => None,
        }
        .filter(|p| p.steps <= 131_072);
        let proof = proof
            .or_else(|| match instruction {
                Instruction::Sequence { left, right } => match (guards[left], proofs[right]) {
                    (Some(a), Some(b)) => Some(LeadingBytes {
                        class: ByteClass([
                            a.class.0[0] & b.class.0[0],
                            a.class.0[1] & b.class.0[1],
                        ]),
                        steps: a.steps + b.steps + 3,
                    }),
                    _ => None,
                },
                _ => None,
            })
            .filter(|p| p.steps <= 131_072);
        failures.push(failure);
        guards.push(guard);
        proofs.push(proof);
    }
    proofs
}
