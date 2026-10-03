//! Conservative lexical certificates. Unknown, stateful, nullable or capturing
//! constructions stay on the full engine; no speculative heuristic is used.
use crate::{
    Builtin, CompiledGrammar, RuleMode,
    program::{Instruction, Program},
};

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
                let predicate: fn(u8) -> bool = match builtin {
                    Builtin::Ascii => |_| true,
                    Builtin::AsciiDigit => |b| b.is_ascii_digit(),
                    Builtin::AsciiAlpha => |b| b.is_ascii_alphabetic(),
                    Builtin::AsciiAlphaLower => |b| b.is_ascii_lowercase(),
                    Builtin::AsciiAlphaUpper => |b| b.is_ascii_uppercase(),
                    Builtin::AsciiAlphanumeric => |b| b.is_ascii_alphanumeric(),
                    Builtin::AsciiHexDigit => |b| b.is_ascii_hexdigit(),
                    Builtin::AsciiNonzeroDigit => |b| (b'1'..=b'9').contains(&b),
                    Builtin::AsciiBinDigit => |b| matches!(b, b'0' | b'1'),
                    Builtin::AsciiOctDigit => |b| (b'0'..=b'7').contains(&b),
                    _ => {
                        classes.push(None);
                        continue;
                    }
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
