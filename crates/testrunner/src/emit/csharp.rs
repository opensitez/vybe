//! C# emitter: one extracted case → a standalone `.cs` test.
//!
//! `Console.WriteLine(x)` writes one line, so calls pair positionally. The
//! corpus is written as top-level statements, so the harness is spliced in
//! front as a local function and nothing is reflowed.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing }

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("// vybe-test: {slug}\n// origin: {origin}\n");

    let Some(expected) = case.expected.as_ref() else {
        return Emitted {
            text: format!("{header}// vybe-test-mode: compile\n\n{}\n", case.source.trim()),
            pairing: Pairing::Direct };
    };

    let prints = find_prints(&case.source);
    if let Some(reason) = unpairable(&case.source, &prints, expected.len()) {
        return Emitted {
            text: format!("{header}\n{}\n", case.source.trim()),
            pairing: Pairing::Unpairable(reason) };
    }

    let mut body = case.source.clone();
    for (i, span) in prints.iter().enumerate().rev() {
        let args = case.source[span.1..span.2].trim();
        // `.ToString()`, NOT `Convert.ToString(...)`. Measured against vybex:
        // `WriteLine(b)` and `b.ToString()` both give `True`, while
        // `Convert.ToString(b)` gives `true` and `"" + b` gives `0` — the
        // wrong one cost 3,323 false failures. (Real C# gives `True` for all
        // three, so those two are Vybe bugs in their own right.)
        let call = format!("__Check(({args}).ToString(), {})", cs_string(&expected[i]));
        body.replace_range(span.0..span.3, &call);
    }

    Emitted {
        text: format!("{header}\n{harness}\n\n{}\n", body.trim()),
        pairing: Pairing::Direct }
}

fn unpairable(src: &str, prints: &[Span], expected: usize) -> Option<String> {
    if prints.is_empty() {
        return Some("no Console.WriteLine to pair".into());
    }
    if has_word(src, "for") || has_word(src, "foreach") || has_word(src, "while") {
        return Some("loop — print count is not static".into());
    }
    // `Console.Write` emits no newline, so consecutive calls share a line.
    if src.contains("Console.Write(") {
        return Some("Console.Write without newline — not one line per call".into());
    }
    if prints.len() != expected {
        return Some(format!(
            "{} WriteLine call(s) but {expected} expected line(s)",
            prints.len()
        ));
    }
    None
}

/// (call start, args start, args end, call end)
type Span = (usize, usize, usize, usize);

fn find_prints(src: &str) -> Vec<Span> {
    const NEEDLE: &str = "Console.WriteLine(";
    let bytes = src.as_bytes();
    let mut spans = Vec::new();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_atom(src, bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i) && src[i..].starts_with(NEEDLE) {
            let args_start = i + NEEDLE.len();
            if let Some(end) = close_paren(src, bytes, args_start) {
                // `Console.WriteLine()` writes a blank line — nothing to pair.
                if src[args_start..end].trim().is_empty() {
                    i = end + 1;
                    continue;
                }
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

/// C# strings (`"…"` with backslash escapes, verbatim `@"…"` where `""` escapes
/// a quote), char literals, and both comment forms.
fn skip_atom(src: &str, bytes: &[u8], at: usize) -> Option<usize> {
    if bytes.get(at) == Some(&b'/') {
        match bytes.get(at + 1) {
            Some(b'/') => {
                let mut i = at;
                while i < bytes.len() && bytes[i] != b'\n' {
                    i += 1;
                }
                return Some(i);
            }
            Some(b'*') => {
                let mut i = at + 2;
                while i + 1 < bytes.len() && !(bytes[i] == b'*' && bytes[i + 1] == b'/') {
                    i += 1;
                }
                return Some((i + 2).min(bytes.len()));
            }
            _ => return None }
    }

    let verbatim = bytes.get(at) == Some(&b'@') && bytes.get(at + 1) == Some(&b'"');
    let start = if verbatim { at + 1 } else { at };
    match bytes.get(start)? {
        b'"' => {
            let mut i = start + 1;
            while i < bytes.len() {
                if verbatim {
                    if bytes[i] == b'"' {
                        if bytes.get(i + 1) == Some(&b'"') {
                            i += 2;
                            continue;
                        }
                        return Some(i + 1);
                    }
                } else {
                    match bytes[i] {
                        b'\\' => {
                            i += 2;
                            continue;
                        }
                        b'"' => return Some(i + 1),
                        _ => {}
                    }
                }
                i += 1;
            }
            Some(bytes.len())
        }
        b'\'' if !verbatim => {
            let mut i = start + 1;
            while i < bytes.len() {
                match bytes[i] {
                    b'\\' => i += 2,
                    b'\'' => return Some(i + 1),
                    _ => i += 1 }
            }
            Some(bytes.len())
        }
        _ => {
            let _ = src;
            None
        }
    }
}

fn has_word(src: &str, word: &str) -> bool {
    let bytes = src.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_atom(src, bytes, i) {
            i = next;
            continue;
        }
        // BOTH ends must sit on a code-point boundary — slicing into a
        // multi-byte character panics and aborts the whole extraction.
        if src.is_char_boundary(i)
            && src.len() >= i + word.len()
            && src.is_char_boundary(i + word.len())
            && &src[i..i + word.len()] == word
        {
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

fn cs_string(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 2);
    out.push('"');
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '"' => out.push_str("\\\""),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch) }
    }
    out.push('"');
    out
}
