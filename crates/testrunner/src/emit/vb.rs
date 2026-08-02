//! VB emitter: one extracted case → a standalone `.vb` test.
//!
//! `Console.WriteLine(x)` writes one line, so calls pair positionally like Go's
//! `Println` and Kotlin's `println`. VB is line-oriented and has no statement
//! separator to unfold, so nothing is reflowed — each call is replaced in
//! place.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("' vybe-test: {slug}\n' origin: {origin}\n");

    let Some(raw) = case.expected.as_ref() else {
        return Emitted {
            text: format!("{header}' vybe-test-mode: compile\n\n{}\n", case.source.trim()),
            pairing: Pairing::Direct,
        };
    };

    // The `vb_*_spec!`/`vb_case!` macros pass expectations through
    // `dotnet_expected_one`, which rewrites the Rust-style `true`/`false` to
    // the `True`/`False` .NET actually prints.
    let expected: Vec<String> = raw
        .iter()
        .map(|e| match e.as_str() {
            "true" => "True".to_string(),
            "false" => "False".to_string(),
            other => other.to_string(),
        })
        .collect();

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
        let call = format!("__Check(CStr({args}), {})", vb_string(&expected[i]));
        body.replace_range(span.0..span.3, &call);
    }

    Emitted {
        text: format!("{header}\n{harness}\n\n{}\n", body.trim()),
        pairing: Pairing::Direct,
    }
}

fn unpairable(src: &str, prints: &[Span], expected: usize) -> Option<String> {
    if prints.is_empty() {
        return Some("no Console.WriteLine to pair".into());
    }
    if has_word(src, "For") || has_word(src, "While") || has_word(src, "Do") {
        return Some("loop — print count is not static".into());
    }
    // `Console.Write` leaves the cursor on the same line, so consecutive calls
    // share one output line.
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
                // `Console.WriteLine()` with no argument writes a blank line;
                // there is nothing to convert.
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

/// A VB string (`""` escapes a quote) or a `'` comment to end of line.
fn skip_atom(_src: &str, bytes: &[u8], at: usize) -> Option<usize> {
    match bytes.get(at)? {
        b'"' => {
            let mut i = at + 1;
            while i < bytes.len() {
                if bytes[i] == b'"' {
                    if bytes.get(i + 1) == Some(&b'"') {
                        i += 2;
                        continue;
                    }
                    return Some(i + 1);
                }
                i += 1;
            }
            Some(bytes.len())
        }
        b'\'' => {
            let mut i = at;
            while i < bytes.len() && bytes[i] != b'\n' {
                i += 1;
            }
            Some(i)
        }
        _ => None,
    }
}

/// VB keywords are case-insensitive.
fn has_word(src: &str, word: &str) -> bool {
    let bytes = src.as_bytes();
    let lower = word.to_ascii_lowercase();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_atom(src, bytes, i) {
            i = next;
            continue;
        }
        // Both ends must land on a code-point boundary — a `ß` in a VB source
        // string sits inside the window otherwise and slicing panics.
        if src.is_char_boundary(i)
            && src.len() >= i + word.len()
            && src.is_char_boundary(i + word.len())
        {
            let candidate = &src[i..i + word.len()];
            if candidate.to_ascii_lowercase() == lower {
                let before = if i == 0 { b' ' } else { bytes[i - 1] };
                let after = bytes.get(i + word.len()).copied().unwrap_or(b' ');
                if !is_ident(before) && !is_ident(after) {
                    return true;
                }
            }
        }
        i += 1;
    }
    false
}

fn is_ident(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_'
}

/// VB has no backslash escapes; a quote is doubled.
fn vb_string(text: &str) -> String {
    format!("\"{}\"", text.replace('"', "\"\""))
}
