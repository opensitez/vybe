//! C# emitter: one extracted case → a standalone `.cs` test.
//!
//! **Output is COLLECTED, not paired.** Every `Console.WriteLine(x)` becomes
//! `__P((x).ToString())`, appending to a buffer the harness compares once at
//! the end. Pairing the i-th print with the i-th expected line cannot assert
//! anything about a loop, and that alone left 706 of 7,622 cases without an
//! assertion. `Console.Write` (no newline) is no longer a blocker either: it
//! becomes `__Pr`.
//!
//! The corpus is written as top-level statements, so the harness is spliced in
//! front as local functions over a local buffer, nothing is reflowed, and the
//! final check goes at the END OF THE FILE — there is no `Main` to close.

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

    let (body, rewritten) = rewrite_prints(&case.source);
    if let Some(reason) = unpairable(&case.source, rewritten) {
        return Emitted {
            text: format!("{header}\n{}\n", case.source.trim()),
            pairing: Pairing::Unpairable(reason) };
    }

    let want = cs_string(&expected.join("\n"));
    let body = format!("{}\n__Check({want});", body.trim());

    Emitted {
        text: format!("{header}\n{harness}\n\n{body}\n"),
        pairing: Pairing::Direct }
}

/// Rewrite every print into a buffer append. Returns the new source and how
/// many calls were rewritten.
fn rewrite_prints(src: &str) -> (String, usize) {
    let bytes = src.as_bytes();
    let mut out = String::with_capacity(src.len());
    let mut count = 0usize;
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_atom(src, bytes, i) {
            out.push_str(&src[i..next]);
            i = next;
            continue;
        }
        // Longest first: `WriteLine(` also starts with `Write(` once the
        // `Console.` prefix is consumed.
        let hit = ["Console.WriteLine(", "Console.Write("]
            .into_iter()
            .find(|n| src.is_char_boundary(i) && src[i..].starts_with(n));
        if let Some(needle) = hit {
            let args_start = i + needle.len();
            if let Some(end) = close_paren(src, bytes, args_start) {
                let args = src[args_start..end].trim();
                let target = if needle.ends_with("WriteLine(") { "__P" } else { "__Pr" };
                // Render HERE, where the expression still has its static type,
                // and with `.ToString()` — see harness/csharp/check.cs for why
                // `Convert.ToString` and `"" + x` are both wrong under Vybe.
                // `WriteLine()` with no argument writes a bare newline.
                let rendered = if args.is_empty() {
                    "\"\"".to_string()
                } else {
                    format!("({args}).ToString()")
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

/// Under collection the only thing that defeats the check is output that never
/// reaches the buffer. A loop, a mismatched count and a newline-less `Write`
/// are all ordinary now.
fn unpairable(src: &str, rewritten: usize) -> Option<String> {
    if rewritten == 0 {
        return Some("no Console.Write call to collect".into());
    }
    // A format call writes many values per call through its own path.
    if src.contains("Console.Out") || src.contains("Console.Error") {
        return Some("writes to a stream other than Console.Write".into());
    }
    None
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
