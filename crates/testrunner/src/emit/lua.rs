//! Lua emitter: one extracted case → a standalone `.lua` test.
//!
//! Checks are at the CALL SITE, not collected: Lua's `print` cannot be
//! intercepted under Vybe — `function print()`, `print = function` and
//! `_G.print = function` all fall through to the real builtin (measured), and
//! `io.write` is undefined. So there is nothing to collect into.
//!
//! Only the FIRST print is compared, because the corpus asserts through
//! `run_lua_one`, which returns only the first line of the run. The counter
//! still advances on every print, which is what lets a print inside a loop be
//! handled — the check fires on print #1 wherever it happens to occur.
//!
//! See `harness/lua/check.lua`.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, _harness: &str) -> Emitted {
    let header = format!("-- vybe-test: {slug}\n-- origin: {origin}\n");
    let body = case.source.trim_end();

    let Some(expected) = case.expected.as_ref() else {
        return Emitted {
            text: format!("{header}\n{body}\n"),
            pairing: Pairing::Direct,
        };
    };

    let prints = find_prints(body);
    if let Some(reason) = unpairable(body, &prints) {
        return Emitted {
            text: format!("{header}\n{body}\n"),
            pairing: Pairing::Unpairable(reason),
        };
    }

    let mut out = body.to_string();
    for p in prints.iter().rev() {
        out.replace_range(p.start..p.end, &check_for(p));
    }

    // `run_lua_one` compares the first line only; an empty expectation there
    // means the run produced nothing.
    let want = lua_string(expected.first().map(String::as_str).unwrap_or(""));
    Emitted {
        text: format!(
            "{header}\nlocal __w1 = {want}\nlocal __i = 0\n\n{out}\n\n\
             if __i == 0 then error(\"FAIL: no output, wanted [\" .. __w1 .. \"]\") end\n"
        ),
        pairing: Pairing::Direct,
    }
}

/// One print, rewritten to render its arguments and check print #1.
fn check_for(p: &Print) -> String {
    format!(
        "do local __t = {}; __i = __i + 1\n  \
         if __i == 1 and __t ~= __w1 then \
         error(\"FAIL: want [\" .. __w1 .. \"] got [\" .. __t .. \"]\") end end",
        p.rendered
    )
}

struct Print {
    start: usize,
    end: usize,
    /// The arguments rendered the way `print` renders them.
    rendered: String,
}

fn find_prints(src: &str) -> Vec<Print> {
    let bytes = src.as_bytes();
    let mut out = Vec::new();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i) && src[i..].starts_with("print(") {
            let before = if i == 0 { b' ' } else { bytes[i - 1] };
            // Not `myprint(` and not `io.print(`.
            if !is_ident(before) {
                let open = i + "print(".len();
                if let Some(close) = close_paren(bytes, open) {
                    let args = split_args(&src[open..close]);
                    // `print` joins its arguments with a TAB and `tostring`s
                    // each — verified identical in `lua` and Vybe.
                    let rendered = if args.is_empty() {
                        "\"\"".to_string()
                    } else {
                        args.iter()
                            .map(|a| format!("tostring({})", a.trim()))
                            .collect::<Vec<_>>()
                            .join(" .. \"\\t\" .. ")
                    };
                    out.push(Print {
                        start: i,
                        end: close + 1,
                        rendered,
                    });
                    i = close + 1;
                    continue;
                }
            }
        }
        i += 1;
    }
    out
}

fn unpairable(src: &str, prints: &[Print]) -> Option<String> {
    if prints.is_empty() {
        return Some("no print to pair".into());
    }
    // A program that reassigns `print` shadows the rewrite target, and one
    // that writes through `io` bypasses it entirely.
    if src.contains("print =") || src.contains("print=") || src.contains("_G.print") {
        return Some("reassigns `print` — the rewrite would be shadowed".into());
    }
    if src.contains("io.write") || src.contains("io.stdout") {
        return Some("writes through `io` — not a print call".into());
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

/// Split on top-level commas, respecting strings, parens, braces and brackets.
fn split_args(text: &str) -> Vec<String> {
    let bytes = text.as_bytes();
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut depth = 0i32;
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            cur.push_str(&text[i..next]);
            i = next;
            continue;
        }
        let ch = bytes[i] as char;
        match ch {
            '(' | '[' | '{' => {
                depth += 1;
                cur.push(ch);
            }
            ')' | ']' | '}' => {
                depth -= 1;
                cur.push(ch);
            }
            ',' if depth == 0 => out.push(std::mem::take(&mut cur)),
            _ => cur.push(ch),
        }
        i += 1;
    }
    if !cur.trim().is_empty() {
        out.push(cur);
    }
    out
}

/// Lua string and long-bracket literals.
fn skip_literal(bytes: &[u8], at: usize) -> Option<usize> {
    match bytes.get(at)? {
        b'"' | b'\'' => {
            let quote = bytes[at];
            let mut i = at + 1;
            while i < bytes.len() {
                match bytes[i] {
                    b'\\' => i += 2,
                    b if b == quote => return Some(i + 1),
                    _ => i += 1,
                }
            }
            Some(bytes.len())
        }
        // `[[ … ]]` long string.
        b'[' if bytes.get(at + 1) == Some(&b'[') => {
            let mut i = at + 2;
            while i + 1 < bytes.len() {
                if bytes[i] == b']' && bytes[i + 1] == b']' {
                    return Some(i + 2);
                }
                i += 1;
            }
            Some(bytes.len())
        }
        _ => None,
    }
}

fn is_ident(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_' || byte == b'.' || byte == b':'
}

fn lua_string(text: &str) -> String {
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
