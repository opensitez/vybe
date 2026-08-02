//! Go emitter: turn one extracted case into a standalone `.go` test.
//!
//! This file holds only the *rewrite* — mapping each `fmt.Println(x)` to the
//! expected line it produces. The assertion itself lives in Go, in
//! `harness/go/check.go`, and is passed alongside the test at run time the way
//! test262 passes `assert.js`.

use crate::extract::Case;

pub enum Pairing {
    /// Every expected line was matched to the print that produces it.
    Direct,
    /// The print-to-line mapping is not static; the case is reported rather
    /// than guessed at.
    Unpairable(String),
}

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let mut header = format!("// vybe-test: {slug}\n// origin: {origin}\n");

    let Some(expected) = case.expected.as_ref() else {
        let mode = if case.expect_failure { "compile-fail" } else { "compile" };
        header.push_str(&format!("// vybe-test-mode: {mode}\n\n"));
        return Emitted {
            text: header + &reflow(&case.source),
            // A compile case never had an expectation to pair.
            pairing: Pairing::Direct,
        };
    };
    header.push('\n');

    let prints = find_prints(&case.source);

    if let Some(reason) = unpairable_reason(&case.source, &prints, expected.len()) {
        return Emitted {
            text: header + &reflow(&case.source),
            pairing: Pairing::Unpairable(reason),
        };
    }

    // Rewrite back-to-front so the earlier spans keep their byte offsets.
    let mut body = case.source.clone();
    for (i, span) in prints.iter().enumerate().rev() {
        let args = &case.source[span.args_start..span.args_end];
        let replacement = format!(
            "__check({}, {})",
            render_println_args(args),
            go_string(&expected[i])
        );
        body.replace_range(span.start..span.end, &replacement);
    }

    Emitted {
        text: header + &splice_harness(&reflow(&body), harness),
        pairing: Pairing::Direct,
    }
}

/// Put the harness in front of `func main`, where a reader looking for the
/// program itself will scroll past it rather than through it.
fn splice_harness(src: &str, harness: &str) -> String {
    match src.find("func main(") {
        Some(at) => format!("{}{harness}\n\n{}", &src[..at], &src[at..]),
        None => format!("{src}\n{harness}\n"),
    }
}

fn unpairable_reason(src: &str, prints: &[Span], expected: usize) -> Option<String> {
    if has_keyword(src, "for") || has_keyword(src, "range") {
        return Some("loop — print count is not static".into());
    }
    if has_call(src, "fmt.Printf") || has_call(src, "fmt.Print") {
        return Some("Printf/Print — output is not one line per call".into());
    }
    if prints.len() != expected {
        return Some(format!(
            "{} Println call(s) but {expected} expected line(s)",
            prints.len()
        ));
    }
    if prints.is_empty() {
        return Some("no output calls to pair".into());
    }
    // `fmt.Println(xs...)` spreads a slice: how many values it prints, and
    // therefore where the spaces fall, is not visible in the source.
    if prints.iter().any(|span| {
        split_top_level(&src[span.args_start..span.args_end])
            .iter()
            .any(|p| p.ends_with("..."))
    }) {
        return Some("variadic spread — Println's spacing is not static".into());
    }
    None
}

/// The expression producing exactly the line `fmt.Println(args)` would print.
///
/// NOT `fmt.Sprint(args)`. `Println` always puts a space between operands;
/// `Sprint` inserts one only "between operands when neither is a string", so
/// `fmt.Println("a", "b")` prints `a b` while `fmt.Sprint("a", "b")` yields
/// `ab`. Emitting the latter mis-reported 53 passing tests as failures with the
/// tell-tale `want [a b] got [ab]`.
///
/// `fmt.Sprintln` would be the exact counterpart but is not in the Go profile,
/// so each operand is rendered on its own and joined with a literal space —
/// which needs no language feature the test didn't already use.
fn render_println_args(args: &str) -> String {
    let parts = split_top_level(args);
    match parts.len() {
        0 => "\"\"".to_string(),
        1 => format!("fmt.Sprint({})", parts[0]),
        _ => parts
            .iter()
            .map(|p| format!("fmt.Sprint({p})"))
            .collect::<Vec<_>>()
            .join(" + \" \" + "),
    }
}

/// Split a call's arguments on the commas that separate them — not on commas
/// inside a nested call, composite literal, index or string.
fn split_top_level(args: &str) -> Vec<String> {
    let bytes = args.as_bytes();
    let mut parts = Vec::new();
    let mut depth = 0usize;
    let mut start = 0usize;
    let mut i = 0usize;

    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        match bytes[i] {
            b'(' | b'[' | b'{' => depth += 1,
            b')' | b']' | b'}' => depth = depth.saturating_sub(1),
            b',' if depth == 0 => {
                parts.push(args[start..i].trim().to_string());
                start = i + 1;
            }
            _ => {}
        }
        i += 1;
    }
    let tail = args[start..].trim();
    if !tail.is_empty() {
        parts.push(tail.to_string());
    }
    parts.retain(|p| !p.is_empty());
    parts
}

struct Span {
    start: usize,
    args_start: usize,
    args_end: usize,
    end: usize,
}

/// Every `fmt.Println(...)` call outside a string literal, in source order.
fn find_prints(src: &str) -> Vec<Span> {
    const NEEDLE: &str = "fmt.Println(";
    let bytes = src.as_bytes();
    let mut spans = Vec::new();
    let mut i = 0usize;

    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src[i..].starts_with(NEEDLE) {
            let args_start = i + NEEDLE.len();
            if let Some(args_end) = close_paren(bytes, args_start) {
                spans.push(Span { start: i, args_start, args_end, end: args_end + 1 });
                i = args_end + 1;
                continue;
            }
        }
        i += 1;
    }
    spans
}

/// Index of the `)` closing a call whose arguments start at `from`.
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

/// If `at` opens a Go literal, the index just past its close. Go has three:
/// interpreted `"..."` (backslash escapes), raw `` `...` `` (none — this is
/// where struct tags live), and rune `'...'`.
fn skip_literal(bytes: &[u8], at: usize) -> Option<usize> {
    match bytes.get(at)? {
        b'"' => {
            let mut i = at + 1;
            while i < bytes.len() {
                match bytes[i] {
                    b'\\' => i += 2,
                    b'"' => return Some(i + 1),
                    _ => i += 1,
                }
            }
            Some(bytes.len())
        }
        b'`' => {
            let mut i = at + 1;
            while i < bytes.len() && bytes[i] != b'`' {
                i += 1;
            }
            Some((i + 1).min(bytes.len()))
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

fn has_keyword(src: &str, word: &str) -> bool {
    let bytes = src.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src[i..].starts_with(word) {
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

fn has_call(src: &str, name: &str) -> bool {
    let bytes = src.as_bytes();
    let needle = format!("{name}(");
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src[i..].starts_with(&needle) {
            return true;
        }
        i += 1;
    }
    false
}

fn is_ident(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_' || byte == b'.'
}

/// A Go interpreted string literal holding `text`.
fn go_string(text: &str) -> String {
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

/// The corpus writes whole programs on one line with `;` separators. Split them
/// back into statements so a person can read — and debug — the emitted file.
///
/// ONLY on `;`, never on braces. Replacing an explicit semicolon with a newline
/// is safe by construction in Go: automatic semicolon insertion puts back
/// exactly the token that was removed. A brace is not safe — `{` opens a block
/// in `func main() {` and a composite literal in `x := Data{A: 1}`, and telling
/// those apart is the composite-literal ambiguity the real grammar resolves
/// with parser context. Guessing at it turned 4,130 of 6,393 emitted files into
/// syntax errors.
///
/// Semicolons inside a control-statement header are the language's own, not
/// separators: `for a; b; c {`, `if x := f(); ok {`, `switch v := i.(type) {`.
/// Those run from the keyword to the `{` that opens the body.
fn reflow(src: &str) -> String {
    let bytes = src.as_bytes();
    let mut lines: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut paren = 0usize;
    let mut in_clause = false;
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
            '{' if in_clause => in_clause = false,
            _ => {}
        }
        if !in_clause && paren == 0 && starts_control_clause(src, bytes, i) {
            in_clause = true;
        }

        if ch == ';' && paren == 0 && !in_clause {
            push_line(&mut lines, &mut current);
            i += 1;
            continue;
        }
        current.push(ch);
        i += 1;
    }
    push_line(&mut lines, &mut current);

    let mut out = String::new();
    for line in &lines {
        out.push_str(line);
        out.push('\n');
    }
    out
}

/// Whether a control-statement header opens at `at` — the keyword must stand as
/// its own token so `format` and `notif` don't trip it.
fn starts_control_clause(src: &str, bytes: &[u8], at: usize) -> bool {
    const KEYWORDS: [&str; 3] = ["for", "if", "switch"];
    let before = if at == 0 { b' ' } else { bytes[at - 1] };
    if is_ident(before) {
        return false;
    }
    KEYWORDS.iter().any(|kw| {
        src[at..].starts_with(kw)
            && !is_ident(bytes.get(at + kw.len()).copied().unwrap_or(b' '))
    })
}

fn push_line(lines: &mut Vec<String>, current: &mut String) {
    let text = current.trim().to_string();
    if !text.is_empty() {
        lines.push(text);
    }
    current.clear();
}

