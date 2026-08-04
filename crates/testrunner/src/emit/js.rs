//! JS emitter: one extracted case → a standalone `.js` test.
//!
//! **Output is COLLECTED, not paired.** Every `console.log(a)` becomes
//! `__p(__line(a))`, appending to a buffer compared once at the end.
//!
//! That is what makes ASYNC assertable: 967 of the 1,860 refused cases used
//! `await` / `then` / `Promise`, where the i-th log in the source is not the
//! i-th line of output. The buffer records the order things actually ran, so
//! the emitter needs no ordering analysis at all — and the check is deferred
//! through `setTimeout(…, 0)`, a macrotask that fires only once the microtask
//! queue has drained. Loops go the same way.
//!
//! JS needs no reflow — the sources are already multi-line `r#"…"#` blocks, so
//! they are copied verbatim.

use crate::extract::Case;
use crate::emit::go::Pairing;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing }

/// Does the SOURCE open with a `"use strict"` directive?
///
/// Only the program's own directive prologue counts — a `"use strict"` inside a
/// function makes that function strict and nothing else, so hoisting it would
/// change the program.
fn opens_strict(src: &str) -> bool {
    for line in src.lines() {
        let t = line.trim();
        if t.is_empty() || t.starts_with("//") {
            continue;
        }
        // STARTS with, not equals: the corpus writes whole one-liner programs,
        // so the directive shares its line with the first statement
        // (`"use strict"; const o = Object.freeze(…);`).
        return t.starts_with("\"use strict\"") || t.starts_with("'use strict'");
    }
    false
}

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let mut header = format!("// vybe-test: {slug}\n// origin: {origin}\n\n");
    // The harness is spliced ABOVE the program, which pushes the source's own
    // `"use strict"` out of directive-prologue position — it landed around line
    // 47 and stopped being a directive at all, so every strict-mode test ran
    // SLOPPY. It passed anyway while Vybe threw regardless of mode; once the
    // mode was honoured, `test_js_class_getter_only_assignment_throws_in_strict_mode`
    // failed and exposed the whole class.
    //
    // Re-stating it as the file's first statement restores the prologue, and
    // costs nothing when the source was not strict.
    if opens_strict(&case.source) {
        header.push_str("\"use strict\";\n\n");
    }

    let Some(expected) = case.expected.as_ref() else {
        header.push_str("// vybe-test-mode: compile\n\n");
        return Emitted { text: header + case.source.trim(), pairing: Pairing::Direct };
    };

    let logs = find_logs(&case.source);
    if let Some(reason) = unpairable(&case.source, &logs) {
        return Emitted {
            text: header + harness + "\n\n" + case.source.trim() + "\n",
            pairing: Pairing::Unpairable(reason) };
    }

    let mut body = case.source.clone();
    for span in logs.iter().rev() {
        let args = case.source[span.1..span.2].trim();
        let call = if args.is_empty() {
            "__p(\"\")".to_string()
        } else {
            format!("__p(__line({args}))")
        };
        body.replace_range(span.0..span.3, &call);
    }

    // `_one` joins every line with "\n", which is exactly what the buffer
    // holds, so both readings collapse to the same comparison here.
    let want = js_string(&expected.join("\n"));
    let body = format!("{}\n__checkLater({want});", body.trim_end());

    Emitted {
        text: header + harness + "\n\n" + body.trim() + "\n",
        pairing: Pairing::Direct }
}

/// Under collection the only thing that defeats the check is output that never
/// reaches the buffer. Async ordering, loops and count mismatches are all
/// ordinary now — the buffer is filled in RUNTIME order and read after the
/// microtask queue drains.
fn unpairable(src: &str, logs: &[Span]) -> Option<String> {
    if logs.is_empty() {
        return Some("no console.log to collect".into());
    }
    // Writing to a stream of its own bypasses the buffer entirely.
    if src.contains("process.stdout") || src.contains("console.error") {
        return Some("writes to a stream other than console.log".into());
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
        _ => return None };
    let mut i = at + 1;
    while i < bytes.len() {
        match bytes[i] {
            b'\\' => i += 2,
            c if c == quote => return Some(i + 1),
            _ => i += 1 }
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
            _ => out.push(ch) }
    }
    out.push('"');
    out
}
