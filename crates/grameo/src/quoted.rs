//! Certificates for pure, single-byte-delimited lexical tokens. Unknown
//! constructions retain the PEG matcher, including all resource accounting.
use crate::{Builtin, program::{Instruction, Program}};

#[derive(Debug, Clone, Copy)]
pub struct Fragment {
    pub delimiter: u8,
    pub escapes: [u64; 2],
    pub any_escape: bool,
    pub doubled_delimiter: bool,
    pub forbidden_newline: bool,
    /// The guard rejects LF and CRLF but permits a bare CR.
    pub newline_crlf_only: bool,
}

/// Pure `open ~ (!close ~ ANY)* ~ close` lexical body.
#[derive(Debug, Clone, Copy)]
pub struct DelimitedFragment<'a> {
    pub open: &'a str,
    pub close: &'a str,
}

pub fn certify_delimited(program: &impl Program, rule: usize) -> Option<DelimitedFragment<'_>> {
    let parts = flatten(program, program.rule(rule).expression, false);
    if parts.len() != 3 { return None; }
    let open = literal(program, parts[0])?;
    let close = literal(program, parts[2])?;
    if open.is_empty() || close.is_empty() { return None; }
    let Instruction::Repeat { expression, min: 0, max: None } = program.instruction(parts[1]) else { return None; };
    let body = flatten(program, expression, false);
    if body.len() != 2 || !matches!(program.instruction(body[1]), Instruction::Builtin(Builtin::Any)) { return None; }
    let Instruction::Predicate { expression, positive: false } = program.instruction(ungroup(program, body[0])) else { return None; };
    if literal(program, expression)? != close { return None; }
    Some(DelimitedFragment { open, close })
}

/// Success-only scan; an absent terminator leaves diagnostics to the PEG engine.
pub fn scan_delimited(fragment: DelimitedFragment<'_>, source: &str, start: usize) -> Option<usize> {
    if fragment.open.is_empty() || fragment.close.is_empty() { return None; }
    let rest = source.get(start..)?.strip_prefix(fragment.open)?;
    let end = rest.find(fragment.close)?;
    Some(start + fragment.open.len() + end + fragment.close.len())
}

fn ungroup(program: &impl Program, mut id: usize) -> usize {
    while let Instruction::Group(child) = program.instruction(id) { id = child; }
    id
}
fn flatten(program: &impl Program, root: usize, choice: bool) -> Vec<usize> {
    let mut pending = vec![root];
    let mut out = Vec::new();
    while let Some(id) = pending.pop() {
        let id = ungroup(program, id);
        match program.instruction(id) {
            Instruction::Choice { left, right } if choice => { pending.push(right); pending.push(left); }
            Instruction::Sequence { left, right } if !choice => { pending.push(right); pending.push(left); }
            _ => out.push(id),
        }
    }
    out
}
fn literal<'a>(program: &'a impl Program, id: usize) -> Option<&'a str> {
    match program.instruction(ungroup(program, id)) {
        Instruction::Literal { text, insensitive: false } => Some(text),
        _ => None,
    }
}

/// Build-time proof: delimiter, an unbounded repeat of escapes followed by
/// `!delimiter ~ ANY`, and the same closing delimiter. No rule calls, stack
/// operations, trivia, semantic callbacks, or stateful predicates are omitted.
pub fn certify(program: &impl Program, rule: usize) -> Option<Fragment> {
    let parts = flatten(program, program.rule(rule).expression, false);
    if parts.len() != 3 { return None; }
    let open = literal(program, parts[0])?;
    if open.len() != 1 || !open.is_ascii() || open == "\\" || literal(program, parts[2])? != open { return None; }
    let Instruction::Repeat { expression, min: 0, max: None } = program.instruction(parts[1]) else { return None; };
    let alternatives = flatten(program, expression, true);
    let (&ordinary, escapes) = alternatives.split_last()?;
    let ordinary = flatten(program, ordinary, false);
    if !(2..=3).contains(&ordinary.len()) || !matches!(program.instruction(*ordinary.last()?), Instruction::Builtin(Builtin::Any)) { return None; }
    let Instruction::Predicate { expression, positive: false } = program.instruction(ungroup(program, ordinary[0])) else { return None; };
    if literal(program, expression)? != open { return None; }
    let forbidden_newline = ordinary.len() == 3;
    let mut newline_crlf_only = false;
    if forbidden_newline {
        let Instruction::Predicate { expression, positive: false } = program.instruction(ungroup(program, ordinary[1])) else { return None; };
        match program.instruction(ungroup(program, expression)) {
            Instruction::Builtin(Builtin::Newline) => {},
            Instruction::Call(rule) => {
                // A pure literal choice has no nested calls, stack effects or
                // sequence trivia. Predicate captures are discarded by PEG.
                let alternatives = flatten(program, program.rule(rule).expression, true);
                let mut lf = false;
                let mut cr = false;
                let mut crlf = false;
                for alternative in alternatives {
                    match literal(program, alternative)? {
                        "\n" => lf = true,
                        "\r" => cr = true,
                        "\r\n" => crlf = true,
                        _ => return None,
                    }
                }
                if !lf || !(cr || crlf) { return None; }
                newline_crlf_only = !cr;
            },
            _ => return None,
        }
    }
    let mut fragment = Fragment { delimiter: open.as_bytes()[0], escapes: [0; 2], any_escape: false, doubled_delimiter: false, forbidden_newline, newline_crlf_only };
    for &escape in escapes {
        if let Some(text) = literal(program, escape) {
            let bytes = text.as_bytes();
            if bytes.len() != 2 || !text.is_ascii() { return None; }
            if bytes[0] == b'\\' { fragment.escapes[usize::from(bytes[1] / 64)] |= 1u64 << (bytes[1] % 64); }
            else if bytes[0] == fragment.delimiter && bytes[1] == fragment.delimiter { fragment.doubled_delimiter = true; }
            else { return None; }
        } else {
            let parts = flatten(program, escape, false);
            if parts.len() != 2 || literal(program, parts[0])? != "\\"
                || !matches!(program.instruction(parts[1]), Instruction::Builtin(Builtin::Any)) { return None; }
            fragment.any_escape = true;
        }
    }
    Some(fragment)
}

fn special(bytes: &[u8], delimiter: u8, forbidden_newline: bool) -> usize {
    const ONES: u64 = 0x0101_0101_0101_0101;
    const HIGHS: u64 = 0x8080_8080_8080_8080;
    let delimiter_word = u64::from(delimiter) * ONES;
    let slash_word = u64::from(b'\\') * ONES;
    let zero = |word: u64| word.wrapping_sub(ONES) & !word & HIGHS != 0;
    let mut offset = 0;
    while bytes.len() - offset >= 8 {
        let word = u64::from_ne_bytes(bytes[offset..offset + 8].try_into().unwrap());
        if zero(word ^ delimiter_word) || zero(word ^ slash_word)
            || forbidden_newline && (zero(word ^ (u64::from(b'\r') * ONES)) || zero(word ^ (u64::from(b'\n') * ONES))) { break; }
        offset += 8;
    }
    offset + bytes[offset..].iter().position(|&byte| byte == delimiter || byte == b'\\'
        || forbidden_newline && matches!(byte, b'\r' | b'\n')).unwrap_or(bytes.len() - offset)
}

/// Success-only scan. Failures leave parser state untouched and use ordinary
/// matching for exact expectations. ASCII boundaries are valid UTF-8 boundaries.
pub fn scan(fragment: Fragment, source: &str, start: usize) -> Option<usize> {
    if !fragment.delimiter.is_ascii() || fragment.delimiter == b'\\' { return None; }
    let bytes = source.as_bytes();
    if bytes.get(start) != Some(&fragment.delimiter) { return None; }
    let mut cursor = start + 1;
    loop {
        cursor += special(&bytes[cursor..], fragment.delimiter, fragment.forbidden_newline);
        let byte = *bytes.get(cursor)?;
        if fragment.forbidden_newline && matches!(byte, b'\r' | b'\n') {
            if byte == b'\n' || !fragment.newline_crlf_only || bytes.get(cursor + 1) == Some(&b'\n') { return None; }
            cursor += 1;
            continue;
        }
        if byte == fragment.delimiter {
            if fragment.doubled_delimiter && bytes.get(cursor + 1) == Some(&byte) { cursor += 2; }
            else { return Some(cursor + 1); }
        } else {
            let next = bytes.get(cursor + 1).copied();
            let escaped = next.is_some_and(|next| fragment.any_escape
                || next < 128 && fragment.escapes[usize::from(next / 64)] & (1u64 << (next % 64)) != 0);
            cursor += if escaped { 1 + source[cursor + 1..].chars().next().unwrap().len_utf8() } else { 1 };
        }
    }
}
