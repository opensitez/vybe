//! JS emitter: one extracted case → a standalone `.js` test.
//!
//! Same rewrite as the Go emitter, against `console.log` instead of
//! `fmt.Println`. JS needs no reflow — the sources are already multi-line
//! `r#"…"#` blocks, so they are copied verbatim.

use crate::extract::Case;
use crate::emit::go::Pairing;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let mut header = format!("// vybe-test: {slug}\n// origin: {origin}\n\n");

    let Some(expected) = case.expected.as_ref() else {
        header.push_str("// vybe-test-mode: compile\n\n");
        return Emitted { text: header + case.source.trim(), pairing: Pairing::Direct };
    };

    let logs = find_logs(&case.source);
    if let Some(reason) = unpairable(&case.source, &logs, expected.len(), case.single_line) {
        return Emitted {
            text: header + harness + "\n\n" + case.source.trim() + "\n",
            pairing: Pairing::Unpairable(reason),
        };
    }

    let mut body = case.source.clone();
    for (i, span) in logs.iter().enumerate().rev() {
        let args = case.source[span.1..span.2].trim();
        let replacement = format!("__check(__line({args}), {})", js_string(&expected[i]));
        body.replace_range(span.0..span.3, &replacement);
    }

    Emitted {
        text: header + harness + "\n\n" + body.trim() + "\n",
        pairing: Pairing::Direct,
    }
}

fn unpairable(src: &str, logs: &[Span], expected: usize, single: bool) -> Option<String> {
    if logs.is_empty() {
        return Some("no console.log to pair".into());
    }
    // A `_one` helper compares a single value whose meaning differs by module —
    // the first line in some, every line joined in others. With exactly one
    // print both readings agree; with more they don't, so don't guess.
    if single && logs.len() != 1 {
        return Some(format!("`_one` helper with {} prints — ambiguous", logs.len()));
    }
    if !single && logs.len() != expected {
        return Some(format!("{} console.log call(s) but {expected} expected line(s)", logs.len()));
    }
    // Output order stops matching source order once anything defers work.
    for word in ["setTimeout", "queueMicrotask", "await", "then", "Promise"] {
        if src.contains(word) {
            return Some(format!("async ({word}) — output order is not source order"));
        }
    }
    if has_word(src, "for") || has_word(src, "while") || src.contains("forEach") {
        return Some("loop — print count is not static".into());
    }
    None
}

/// (call start, args start, args end, call end)
type Span = (usize, usize, usize, usize);

fn find_logs(src: &str) -> Vec<Span> {
    const NEEDLE: &str = "console.log(";
    let bytes = src.as_bytes();
    let mut spans = Vec::new();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        // Sources contain non-ASCII outside literals (an em dash in a comment,
        // for one), and slicing into the middle of a code point panics.
        if !src.is_char_boundary(i) {
            i += 1;
            continue;
        }
        if src[i..].starts_with(NEEDLE) {
            let args_start = i + NEEDLE.len();
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

/// JS string, template and regex literals all hide delimiters.
fn skip_literal(bytes: &[u8], at: usize) -> Option<usize> {
    let quote = match bytes.get(at)? {
        c @ (b'"' | b'\'' | b'`') => *c,
        _ => return None,
    };
    let mut i = at + 1;
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => i += 2,
            c if c == quote => return Some(i + 1),
            _ => i += 1,
        }
    }
    Some(bytes.len())
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
            if !before.is_ascii_alphanumeric() && before != b'_'
                && !after.is_ascii_alphanumeric() && after != b'_'
            {
                return true;
            }
        }
        i += 1;
    }
    false
}

fn js_string(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 2);
    out.push('"');
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out.push('"');
    out
}
