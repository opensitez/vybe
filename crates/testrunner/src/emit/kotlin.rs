//! Kotlin emitter: one extracted case → a standalone `.kt` test.
//!
//! `println(x)` takes exactly one argument, so each call produces one line and
//! pairs directly — no argument joining like Go's `Println` or JS's
//! `console.log`.
//!
//! Sources arrive as one-liners with `;` separators and are split back into
//! statements. Only on `;`, never on braces: a brace may open a lambda
//! (`map { it * 2 }`) rather than a block, and Kotlin treats a newline as a
//! statement separator anyway, so replacing an explicit `;` is safe.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("// vybe-test: {slug}\n// origin: {origin}\n");

    let Some(expected) = case.expected.as_ref() else {
        return Emitted {
            text: format!("{header}// vybe-test-mode: compile\n\n{}\n", reflow(&case.source)),
            pairing: Pairing::Direct,
        };
    };

    let prints = find_prints(&case.source);
    if let Some(reason) = unpairable(&case.source, &prints, expected.len()) {
        return Emitted {
            text: format!("{header}\n{}\n", reflow(&case.source)),
            pairing: Pairing::Unpairable(reason),
        };
    }

    let mut body = case.source.clone();
    for (i, span) in prints.iter().enumerate().rev() {
        let args = case.source[span.1..span.2].trim();
        let call = format!("__check(({args}).toString(), {})", kt_string(&expected[i]));
        body.replace_range(span.0..span.3, &call);
    }

    Emitted {
        text: format!("{header}\n{}", splice_harness(&reflow(&body), harness)),
        pairing: Pairing::Direct,
    }
}

fn splice_harness(src: &str, harness: &str) -> String {
    match src.find("fun main(") {
        Some(at) => format!("{}{harness}\n\n{}", &src[..at], &src[at..]),
        None => format!("{src}\n{harness}\n"),
    }
}

fn unpairable(src: &str, prints: &[Span], expected: usize) -> Option<String> {
    if prints.is_empty() {
        return Some("no println to pair".into());
    }
    if has_word(src, "for") || has_word(src, "while") || src.contains("forEach") {
        return Some("loop — print count is not static".into());
    }
    // `print` writes no newline, so consecutive calls share a line.
    if has_call(src, "print") {
        return Some("print without newline — output is not one line per call".into());
    }
    if prints.len() != expected {
        return Some(format!(
            "{} println call(s) but {expected} expected line(s)",
            prints.len()
        ));
    }
    None
}

/// (call start, args start, args end, call end)
type Span = (usize, usize, usize, usize);

fn find_prints(src: &str) -> Vec<Span> {
    let bytes = src.as_bytes();
    let mut spans = Vec::new();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i)
            && src[i..].starts_with("println(")
            && !is_ident(if i == 0 { b' ' } else { bytes[i - 1] })
        {
            let args_start = i + "println(".len();
            if let Some(end) = close_paren(bytes, args_start) {
                spans.push((i, args_start, end, end + 1));
                i = end + 1;
                continue;
            }
        }
        i += 1;
    }
    spans
}

fn close_paren(bytes: &[u8], from: usize) -> Option<usize> {
    let mut depth = 1usize;
    let mut i = from;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        match bytes[i] {
            b'(' => depth += 1,
            b')' => {
                depth -= 1;
                if depth == 0 {
                    return Some(i);
                }
            }
            _ => {}
        }
        i += 1;
    }
    None
}

/// Kotlin string and char literals, including raw `"""…"""`.
fn skip_literal(bytes: &[u8], at: usize) -> Option<usize> {
    match bytes.get(at)? {
        b'"' => {
            let triple = bytes.get(at + 1) == Some(&b'"') && bytes.get(at + 2) == Some(&b'"');
            let mut i = at + if triple { 3 } else { 1 };
            while i < bytes.len() {
                if !triple && bytes[i] == b'\\' {
                    i += 2;
                    continue;
                }
                if triple {
                    if bytes[i] == b'"' && bytes.get(i + 1) == Some(&b'"') && bytes.get(i + 2) == Some(&b'"') {
                        return Some(i + 3);
                    }
                } else if bytes[i] == b'"' {
                    return Some(i + 1);
                }
                i += 1;
            }
            Some(bytes.len())
        }
        b'\'' => {
            let mut i = at + 1;
            while i < bytes.len() {
                match bytes[i] {
                    b'\\' => i += 2,
                    b'\'' => return Some(i + 1),
                    _ => i += 1,
                }
            }
            Some(bytes.len())
        }
        _ => None,
    }
}

fn has_word(src: &str, word: &str) -> bool {
    let bytes = src.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i) && src[i..].starts_with(word) {
            let before = if i == 0 { b' ' } else { bytes[i - 1] };
            let after = bytes.get(i + word.len()).copied().unwrap_or(b' ');
            if !is_ident(before) && !is_ident(after) {
                return true;
            }
        }
        i += 1;
    }
    false
}

/// `print(` but not `println(`.
fn has_call(src: &str, name: &str) -> bool {
    let bytes = src.as_bytes();
    let needle = format!("{name}(");
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i)
            && src[i..].starts_with(&needle)
            && !is_ident(if i == 0 { b' ' } else { bytes[i - 1] })
        {
            return true;
        }
        i += 1;
    }
    false
}

fn is_ident(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_' || byte == b'.'
}

fn kt_string(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 2);
    out.push('"');
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '$' => out.push_str("\\$"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out.push('"');
    out
}

/// Split one-liner sources back into statements, on `;` only.
fn reflow(src: &str) -> String {
    let bytes = src.as_bytes();
    let mut lines: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut paren = 0usize;
    let mut i = 0usize;

    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            current.push_str(&src[i..next]);
            i = next;
            continue;
        }
        let ch = bytes[i] as char;
        match ch {
            '(' => paren += 1,
            ')' => paren = paren.saturating_sub(1),
            _ => {}
        }
        if ch == ';' && paren == 0 {
            let text = current.trim().to_string();
            if !text.is_empty() {
                lines.push(text);
            }
            current.clear();
            i += 1;
            continue;
        }
        current.push(ch);
        i += 1;
    }
    let text = current.trim().to_string();
    if !text.is_empty() {
        lines.push(text);
    }

    let mut out = String::new();
    for line in &lines {
        out.push_str(line);
        out.push('\n');
    }
    out
}
