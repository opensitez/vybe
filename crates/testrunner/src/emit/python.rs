//! Python emitter: one extracted case → a standalone `.py` test.
//!
//! **Output is COLLECTED, not paired.** Every `print(a, b)` becomes
//! `__p(__line(a, b))`, appending to a buffer the harness compares once at the
//! end of the file. Pairing the i-th print with the i-th expected line cannot
//! assert anything about a loop — 936 of Python's cases — nor about an
//! `if`/`else` where only one branch prints, which is most of the 209 cases
//! whose print count simply differed from the line count.
//!
//! `print(..., end='')` becomes `__pr`, which appends without the newline, so
//! prints sharing a line are ordinary now too.
//!
//! Sources are already multi-line and indentation-sensitive, so nothing is
//! reflowed — each print is replaced in place, preserving its column.

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
            text: format!(
                "{header}# vybe-test-mode: compile\n\n{}\n",
                case.source.trim()
            ),
            pairing: Pairing::Direct,
        };
    };

    // `run_python_one` JOINS every line with "\n" (unlike JS's `_one`, which
    // keeps only the first), so a single expected value splits straight back
    // into the line list.
    let expected: Vec<String> = if case.single_line {
        raw.first()
            .map(|s| s.split('\n').map(str::to_string).collect())
            .unwrap_or_default()
    } else {
        raw.clone()
    };

    let prints = find_prints(&case.source);
    if let Some(reason) = unpairable(&case.source, &prints) {
        return Emitted {
            text: format!("{header}\n{}\n", case.source.trim()),
            pairing: Pairing::Unpairable(reason),
        };
    }

    let mut body = case.source.clone();
    for span in prints.iter().rev() {
        let args = case.source[span.1..span.2].trim();
        body.replace_range(span.0..span.3, &collect_call(args));
    }

    let want = py_string(&expected.join("\n"));
    let body = format!("{}\n__check(__buf, {want})", body.trim_end());

    Emitted {
        text: format!("{header}\n{harness}\n\n{}\n", body.trim()),
        pairing: Pairing::Direct,
    }
}

/// One print, rewritten to append to the buffer.
///
/// `end=''` is what makes consecutive prints share a line, so it selects `__pr`
/// (append without a newline) rather than being a reason to give up. Any other
/// `end=` value is passed through as the terminator.
fn collect_call(args: &str) -> String {
    let (values, end) = split_end_kwarg(args);
    let values = values.trim();
    let line = if values.is_empty() {
        "\"\"".to_string()
    } else {
        format!("__line({values})")
    };
    match end {
        // The default terminator is a newline, which `__p` supplies.
        None => format!("__p({line})"),
        Some(e) if e.trim() == "''" || e.trim() == "\"\"" => format!("__pr({line})"),
        Some(e) => format!("__pr({line} + {})", e.trim()),
    }
}

/// Split `a, b, end='x'` into (`a, b`, Some(`'x'`)). Only a TOP-LEVEL `end=`
/// counts — one inside a nested call belongs to that call.
fn split_end_kwarg(args: &str) -> (&str, Option<&str>) {
    let bytes = args.as_bytes();
    let mut depth = 0i32;
    let mut i = 0usize;
    let mut last_comma: Option<usize> = None;
    while i < bytes.len() {
        match bytes[i] {
            b'"' | b'\'' => {
                let quote = bytes[i];
                i += 1;
                while i < bytes.len() && bytes[i] != quote {
                    i += if bytes[i] == b'\\' { 2 } else { 1 };
                }
            }
            b'(' | b'[' | b'{' => depth += 1,
            b')' | b']' | b'}' => depth -= 1,
            b',' if depth == 0 => last_comma = Some(i),
            b'e' if depth == 0 && args[i..].starts_with("end=") => {
                let before_ok = i == 0 || last_comma.is_some_and(|c| c < i);
                if before_ok {
                    let cut = last_comma.unwrap_or(0);
                    let values = if last_comma.is_some() {
                        &args[..cut]
                    } else {
                        ""
                    };
                    return (values, Some(&args[i + 4..]));
                }
            }
            _ => {}
        }
        i += 1;
    }
    (args, None)
}

/// Under collection the only thing that defeats the check is output that never
/// reaches the buffer. Loops, count mismatches and `end=` are all ordinary now.
fn unpairable(src: &str, prints: &[Span]) -> Option<String> {
    if prints.is_empty() {
        return Some("no print() to collect".into());
    }
    for span in prints {
        let args = &src[span.1..span.2];
        // `sep=` changes the join, and `__line` hardcodes a single space.
        if args.contains("sep=") {
            return Some("print(sep=) — the join is not a single space".into());
        }
        // `*xs` hides how many values there are, so `__line` cannot spread it.
        if args.trim_start().starts_with('*') {
            return Some("unpacked print args — count is not static".into());
        }
    }
    // Writing to a stream of its own bypasses the buffer entirely.
    if src.contains("sys.stdout") || src.contains("sys.stderr") {
        return Some("writes to sys.stdout/stderr directly".into());
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
    while prefix < 2
        && matches!(
            bytes.get(start),
            Some(b'f' | b'F' | b'r' | b'R' | b'b' | b'B' | b'u' | b'U')
        )
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
            if bytes[i] == quote
                && bytes.get(i + 1) == Some(&quote)
                && bytes.get(i + 2) == Some(&quote)
            {
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

#[allow(dead_code)]
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
