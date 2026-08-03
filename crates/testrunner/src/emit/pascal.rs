//! Pascal emitter: one extracted case → a standalone `.pas` test.
//!
//! Sources are complete programs (`program T; … begin … end.`), so nothing is
//! wrapped. Every `WriteLn(a, b, c)` becomes `__p(__vs(a) + __vs(b) + __vs(c))`
//! and the whole buffer is compared once, so loops assert like anything else.
//!
//! `__vs` is an OVERLOAD SET, not one conversion: `WriteLn` takes any type, and
//! Vybe has no `WriteStr` (the primitive that would render an argument list
//! exactly as `Write` does). Overloading lets the compiler pick per argument.
//! See `harness/pascal/check.pas`.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing }

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("// vybe-test: {slug}\n// origin: {origin}\n");
    let body = case.source.trim();

    let Some(expected) = case.expected.as_ref() else {
        return Emitted {
            text: format!("{header}{}\n", with_prologue(body, "")),
            pairing: Pairing::Direct };
    };

    if let Some(reason) = unpairable(body) {
        return Emitted {
            text: format!("{header}{}\n", with_prologue(body, "")),
            pairing: Pairing::Unpairable(reason) };
    }

    let collected = rewrite_writes(body);
    let Some(with_check) = insert_check(&collected, &pas_string(&expected.join("\n"))) else {
        return Emitted {
            text: format!("{header}{}\n", with_prologue(body, "")),
            pairing: Pairing::Unpairable("no final `end.` to place the check before".into()) };
    };

    Emitted {
        text: format!("{header}{}\n", with_prologue(&with_check, &harness_decls(harness))),
        pairing: Pairing::Direct }
}

/// `{$mode delphi}` + `uses SysUtils` (fpc needs both for `Format`/`IntToStr`;
/// Vybe accepts and ignores them), then the harness declarations — all placed
/// after the `program` line so the program's own declarations still precede
/// its `begin`.
fn with_prologue(src: &str, decls: &str) -> String {
    let Some(semi) = src.find(';') else {
        return src.to_string();
    };
    let (head, rest) = src.split_at(semi + 1);
    let mut out = String::with_capacity(src.len() + decls.len() + 64);
    out.push_str(head);
    if !src.contains("{$mode") {
        out.push_str("\n{$mode delphi}");
    }
    if !src.to_ascii_lowercase().contains("uses ") {
        out.push_str("\nuses SysUtils;");
    }
    if !decls.is_empty() {
        out.push('\n');
        out.push_str(decls);
    }
    out.push_str(rest);
    out
}

/// The harness file is a complete program; a test needs its declarations —
/// everything between the `uses` clause and the harness's own `begin`.
fn harness_decls(harness: &str) -> String {
    let start = harness
        .find("uses SysUtils;")
        .map(|i| i + "uses SysUtils;".len())
        .unwrap_or(0);
    let tail = &harness[start..];
    // The LAST top-level `begin` opens the harness's own body, which the test
    // does not want.
    let end = tail.rfind("\nbegin").unwrap_or(tail.len());
    tail[..end].trim().to_string()
}

/// `WriteLn(a, b)` → `__p(__vs(a) + __vs(b))`; `Write(...)` → `__pw(...)`.
fn rewrite_writes(src: &str) -> String {
    let mut out = String::with_capacity(src.len());
    let mut rest = src;
    loop {
        let Some((at, name, is_line)) = next_write(rest) else {
            out.push_str(rest);
            return out;
        };
        let open = at + name.len();
        let Some(close) = matching_paren(rest, open) else {
            out.push_str(&rest[..open]);
            rest = &rest[open..];
            continue;
        };
        out.push_str(&rest[..at]);
        let args = split_args(&rest[open + 1..close]);
        let joined = args
            .iter()
            .map(|a| format!("__vs({})", a.trim()))
            .collect::<Vec<_>>()
            .join(" + ");
        out.push_str(&format!("{}({joined})", if is_line { "__p" } else { "__pw" }));
        rest = &rest[close + 1..];
    }
}

/// The next `WriteLn(` / `Write(` at a statement position, case-insensitively.
fn next_write(src: &str) -> Option<(usize, String, bool)> {
    let lower = src.to_ascii_lowercase();
    let mut best: Option<(usize, String, bool)> = None;
    for (needle, is_line) in [("writeln(", true), ("write(", false)] {
        let mut from = 0usize;
        while let Some(off) = lower[from..].find(needle) {
            let at = from + off;
            let before = src[..at].chars().last().unwrap_or(' ');
            // Not `MyWriteLn(`, and not inside a string literal.
            if !before.is_alphanumeric()
                && before != '_'
                && before != '.'
                && src[..at].matches('\'').count() % 2 == 0
            {
                let name = src[at..at + needle.len() - 1].to_string();
                if best.as_ref().is_none_or(|(b, _, _)| at < *b) {
                    best = Some((at, name, is_line));
                }
                break;
            }
            from = at + needle.len();
        }
    }
    best
}

fn matching_paren(src: &str, open: usize) -> Option<usize> {
    let bytes = src.as_bytes();
    let mut depth = 0i32;
    let mut in_str = false;
    for i in open..bytes.len() {
        match bytes[i] {
            b'\'' => in_str = !in_str,
            b'(' if !in_str => depth += 1,
            b')' if !in_str => {
                depth -= 1;
                if depth == 0 {
                    return Some(i);
                }
            }
            _ => {}
        }
    }
    None
}

/// Split an argument list on top-level commas, respecting `'…'` and nesting.
fn split_args(text: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut depth = 0i32;
    let mut in_str = false;
    for ch in text.chars() {
        match ch {
            '\'' => {
                in_str = !in_str;
                cur.push(ch);
            }
            '(' | '[' if !in_str => {
                depth += 1;
                cur.push(ch);
            }
            ')' | ']' if !in_str => {
                depth -= 1;
                cur.push(ch);
            }
            ',' if !in_str && depth == 0 => out.push(std::mem::take(&mut cur)),
            _ => cur.push(ch) }
    }
    if !cur.trim().is_empty() {
        out.push(cur);
    }
    out
}

/// Place the check immediately before the program's final `end.`.
fn insert_check(src: &str, want: &str) -> Option<String> {
    let at = src.rfind("end.")?;
    Some(format!("{}__vybeCheck({want});\n{}", &src[..at], &src[at..]))
}

fn unpairable(src: &str) -> Option<String> {
    let lower = src.to_ascii_lowercase();
    // Writing to a file handle bypasses the buffer. Match a real handle
    // argument — `WriteLn(f, …)` — not merely a call starting with `f`:
    // `writeln(f` also prefixes `WriteLn(Format(…))`, which marked 17 of 20
    // cases unpairable on the first pass.
    if lower.contains("assign(") || lower.contains("rewrite(") || lower.contains("append(") {
        return Some("writes to a file handle — bypasses the buffer".into());
    }
    if has_write_field_width(src) {
        return Some("`Write` with a field width — layout is not plain concatenation".into());
    }
    None
}

/// `Write(x:5)` / `WriteLn(y:8:2)` — a `:` inside a Write argument list that
/// is not the `:=` of an assignment.
fn has_write_field_width(src: &str) -> bool {
    let lower = src.to_ascii_lowercase();
    let bytes = src.as_bytes();
    let mut from = 0usize;
    while let Some(off) = lower[from..].find("write") {
        let at = from + off;
        let Some(open) = src[at..].find('(').map(|o| at + o) else {
            break;
        };
        let Some(close) = matching_paren(src, open) else {
            break;
        };
        let mut in_str = false;
        for i in open..close {
            match bytes[i] {
                b'\'' => in_str = !in_str,
                b':' if !in_str && bytes.get(i + 1) != Some(&b'=') => return true,
                _ => {}
            }
        }
        from = close;
    }
    false
}

fn pas_string(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 2);
    out.push('\'');
    for ch in text.chars() {
        match ch {
            // A quote is doubled in a Pascal string.
            '\'' => out.push_str("''"),
            '\n' => out.push_str("' + #10 + '"),
            '\r' => out.push_str("' + #13 + '"),
            '\t' => out.push_str("' + #9 + '"),
            _ => out.push(ch) }
    }
    out.push('\'');
    out
}
