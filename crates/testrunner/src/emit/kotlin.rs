//! Kotlin emitter: one extracted case → a standalone `.kt` test.
//!
//! **Output is COLLECTED, not paired.** Every `println(x)` becomes `__p(x)`,
//! appending to a buffer the harness compares once at the end of `main`.
//! Pairing the i-th print with the i-th expected line cannot assert anything
//! about a loop, and that alone left 517 of 5,619 cases without an assertion —
//! plus 110 more where the print count simply differed from the line count,
//! which under collection is not a mismatch at all. `print` (no newline) is no
//! longer a blocker either: it becomes `__pr`.
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
            text: format!(
                "{header}// vybe-test-mode: compile\n\n{}\n",
                reflow(&case.source)
            ),
            pairing: Pairing::Direct,
        };
    };

    let (body, rewritten) = rewrite_prints(&case.source);
    if let Some(reason) = unpairable(&case.source, rewritten) {
        return Emitted {
            text: format!("{header}\n{}\n", reflow(&case.source)),
            pairing: Pairing::Unpairable(reason),
        };
    }

    let want = kt_string(&expected.join("\n"));
    let body = close_main_with(&body, &format!("__check({want})"));

    Emitted {
        text: format!("{header}\n{}", splice_harness(&reflow(&body), harness)),
        pairing: Pairing::Direct,
    }
}

/// Put `__check(…)` at the END of `main`, where every print has already run.
///
/// Appending it to the file instead would put it at top level, which Kotlin
/// does not allow, and appending it to the wrong function would compare a
/// buffer that is still being filled.
fn close_main_with(src: &str, call: &str) -> String {
    let Some(at) = find_main(src) else {
        // No `main` to close — 6 of 5,619 cases. Keep the previous behaviour
        // and let the file speak for itself rather than inventing a wrapper.
        return format!("{src}\n{call}\n");
    };
    let Some(open) = src[at..].find('{').map(|o| at + o) else {
        return format!("{src}\n{call}\n");
    };
    match matching_brace(src.as_bytes(), open) {
        Some(close) => format!("{}\n{call}\n{}", &src[..close], &src[close..]),
        None => format!("{src}\n{call}\n"),
    }
}

/// `fun main(` outside a string literal.
fn find_main(src: &str) -> Option<usize> {
    let bytes = src.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i) && src[i..].starts_with("fun main(") {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn matching_brace(bytes: &[u8], open: usize) -> Option<usize> {
    let mut depth = 0usize;
    let mut i = open;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        match bytes[i] {
            b'{' => depth += 1,
            b'}' => {
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

/// Rewrite every print into a buffer append. Returns the new source and how
/// many calls were rewritten.
fn rewrite_prints(src: &str) -> (String, usize) {
    let bytes = src.as_bytes();
    let mut out = String::with_capacity(src.len());
    let mut count = 0usize;
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            out.push_str(&src[i..next]);
            i = next;
            continue;
        }
        // Longest first: `println(` also starts with `print(`, and the Java
        // spelling appears in a handful of Kotlin cases.
        let hit = [
            "System.out.println(",
            "System.out.print(",
            "println(",
            "print(",
        ]
        .into_iter()
        .find(|n| src.is_char_boundary(i) && src[i..].starts_with(n))
        // A qualified call — `writer.println(` — writes somewhere else.
        .filter(|n| n.starts_with("System") || !is_ident(if i == 0 { b' ' } else { bytes[i - 1] }));
        if let Some(needle) = hit {
            let args_start = i + needle.len();
            if let Some(end) = close_paren(bytes, args_start) {
                let args = src[args_start..end].trim();
                let target = if needle.ends_with("println(") {
                    "__p"
                } else {
                    "__pr"
                };
                // Render HERE, where the expression still has its static type.
                // `__p(x)` with an `Any?` parameter renders a Boolean as 1/0
                // and cannot resolve `toString` at all — see harness/kotlin.
                // `println()` with no argument writes a bare newline.
                let rendered = if args.is_empty() {
                    "\"\"".to_string()
                } else {
                    format!("({args}).toString()")
                };
                out.push_str(&format!("{target}({rendered})"));
                count += 1;
                i = end + 1;
                continue;
            }
        }
        let ch = src[i..].chars().next().unwrap_or(' ');
        out.push(ch);
        i += ch.len_utf8();
    }
    (out, count)
}

fn splice_harness(src: &str, harness: &str) -> String {
    match src.find("fun main(") {
        Some(at) => format!("{}{harness}\n\n{}", &src[..at], &src[at..]),
        None => format!("{src}\n{harness}\n"),
    }
}

/// Under collection the only thing that defeats the check is output that never
/// reaches the buffer. A loop, a mismatched count and a newline-less `print`
/// are all ordinary now.
fn unpairable(src: &str, rewritten: usize) -> Option<String> {
    if rewritten == 0 {
        return Some("no print call to collect".into());
    }
    // A format call writes many values per call through its own path.
    if src.contains("System.out.printf") || has_call(src, "printf") {
        return Some("printf — format output is not collectable".into());
    }
    // Handing the stream to something else writes past the buffer, so the
    // collected output would be short and the test would fail for the wrong
    // reason.
    if src.contains("System.out.flush") || src.contains("System.out)") {
        return Some("passes or flushes System.out directly".into());
    }
    None
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
                    if bytes[i] == b'"'
                        && bytes.get(i + 1) == Some(&b'"')
                        && bytes.get(i + 2) == Some(&b'"')
                    {
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
