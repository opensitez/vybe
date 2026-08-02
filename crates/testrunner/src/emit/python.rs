//! Python emitter: one extracted case → a standalone `.py` test.
//!
//! `print(a, b)` composes one line from its arguments joined by a space, so it
//! pairs the same way `fmt.Println` and `console.log` do. Sources are already
//! multi-line and indentation-sensitive, so nothing is reflowed — each print is
//! replaced in place, preserving its column.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("# vybe-test: {slug}\n# origin: {origin}\n");

    let Some(raw) = case.expected.as_ref() else {
        return Emitted {
            text: format!("{header}# vybe-test-mode: compile\n\n{}\n", case.source.trim()),
            pairing: Pairing::Direct,
        };
    };

    // `run_python_one` JOINS every line with "\n" (unlike JS's `_one`, which
    // keeps only the first), so a single expected value splits straight back
    // into the line list.
    let expected: Vec<String> = if case.single_line {
        raw.first().map(|s| s.split('\n').map(str::to_string).collect()).unwrap_or_default()
    } else {
        raw.clone()
    };

    let prints = find_prints(&case.source);
    if let Some(reason) = unpairable(&case.source, &prints, expected.len()) {
        return Emitted {
            text: format!("{header}\n{}\n", case.source.trim()),
            pairing: Pairing::Unpairable(reason),
        };
    }

    let mut body = case.source.clone();
    for (i, span) in prints.iter().enumerate().rev() {
        let args = case.source[span.1..span.2].trim();
        let call = if args.is_empty() {
            format!("__check(\"\", {})", py_string(&expected[i]))
        } else {
            format!("__check(__line({args}), {})", py_string(&expected[i]))
        };
        body.replace_range(span.0..span.3, &call);
    }

    Emitted {
        text: format!("{header}\n{harness}\n\n{}\n", body.trim()),
        pairing: Pairing::Direct,
    }
}

fn unpairable(src: &str, prints: &[Span], expected: usize) -> Option<String> {
    if prints.is_empty() {
        return Some("no print() to pair".into());
    }
    if has_word(src, "for") || has_word(src, "while") {
        return Some("loop — print count is not static".into());
    }
    for span in prints {
        let args = &src[span.1..span.2];
        // `end=` suppresses the newline so prints share a line; `sep=` changes
        // the join; `*xs` hides how many values there are.
        if args.contains("end=") || args.contains("sep=") {
            return Some("print(end=/sep=) — output is not one line per call".into());
        }
        if args.trim_start().starts_with('*') {
            return Some("unpacked print args — count is not static".into());
        }
    }
    if prints.len() != expected {
        return Some(format!(
            "{} print() call(s) but {expected} expected line(s)",
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
        if let Some(next) = skip_atom(src, bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i)
            && src[i..].starts_with("print(")
            && !is_ident(if i == 0 { b' ' } else { bytes[i - 1] })
        {
            let args_start = i + "print(".len();
            if let Some(end) = close_paren(src, bytes, args_start) {
                spans.push((i, args_start, end, end + 1));
                i = end + 1;
                continue;
            }
        }
        i += 1;
    }
    spans
}

fn close_paren(src: &str, bytes: &[u8], from: usize) -> Option<usize> {
    let mut depth = 1usize;
    let mut i = from;
    while i < bytes.len() {
        if let Some(next) = skip_atom(src, bytes, i) {
            i = next;
            continue;
        }
        match bytes[i] {
            b'(' | b'[' | b'{' => depth += 1,
            b')' if depth == 1 => return Some(i),
            b')' | b']' | b'}' => depth -= 1,
            _ => {}
        }
        i += 1;
    }
    None
}

/// Step over a Python string literal or a `#` comment.
///
/// Literals carry prefixes (`f`, `r`, `b`, `rb`, `fr`, …) and come in single
/// and triple-quoted forms; a comment can contain anything, including a
/// convincing `print(`.
fn skip_atom(src: &str, bytes: &[u8], at: usize) -> Option<usize> {
    if bytes.get(at) == Some(&b'#') {
        let mut i = at;
        while i < bytes.len() && bytes[i] != b'\n' {
            i += 1;
        }
        return Some(i);
    }

    // A prefix only counts when a quote follows it directly.
    let mut start = at;
    let mut prefix = 0usize;
    while prefix < 2 && matches!(bytes.get(start), Some(b'f' | b'F' | b'r' | b'R' | b'b' | b'B' | b'u' | b'U'))
    {
        start += 1;
        prefix += 1;
    }
    let quote = *bytes.get(start)?;
    if quote != b'"' && quote != b'\'' {
        return None;
    }
    // An identifier ending in one of those letters is not a prefix.
    if prefix > 0 && at > 0 && is_ident(bytes[at - 1]) {
        return None;
    }
    if prefix > 0 && at > 0 && start != at {
        // `print(` guard: `f"…"` after an identifier char is part of a name.
    }

    let triple = src.is_char_boundary(start)
        && (src[start..].starts_with("\"\"\"") || src[start..].starts_with("'''"));
    let mut i = start + if triple { 3 } else { 1 };
    while i < bytes.len() {
        if bytes[i] == b'\\' {
            i += 2;
            continue;
        }
        if triple {
            if bytes[i] == quote && bytes.get(i + 1) == Some(&quote) && bytes.get(i + 2) == Some(&quote) {
                return Some(i + 3);
            }
        } else if bytes[i] == quote {
            return Some(i + 1);
        } else if bytes[i] == b'\n' {
            // An unterminated single-quoted string cannot span lines.
            return Some(i);
        }
        i += 1;
    }
    Some(bytes.len())
}

fn has_word(src: &str, word: &str) -> bool {
    let bytes = src.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_atom(src, bytes, i) {
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

fn is_ident(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_'
}

fn py_string(text: &str) -> String {
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
