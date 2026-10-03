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
/// Only trivia uses this certificate: failed trivia emits no diagnostics or
/// semantic hooks. A possible match always executes the original grammar.
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
    let mut needed = vec![false; grammar.syntax().rules.len()];
    let mut tasks = Vec::new();
    for scope in grammar.scopes() {
        for rule in scope.whitespace.into_iter().chain(scope.comment) {
            needed[rule] = true;
            tasks.push((grammar.rule(rule).expression, false));
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
