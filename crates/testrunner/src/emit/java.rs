//! Java emitter: one extracted case → a standalone `.java` test.
//!
//! Two things make this different from the Kotlin/C# emitters it otherwise
//! resembles.
//!
//! **The harness goes INSIDE the class.** Java has no top-level functions, and
//! `run_main` wraps the body in `public class Main { … }`, so `__check` has to
//! be a static member of `Main` rather than something prepended to the file.
//!
//! **A case may carry a second source.** `run_in_main(main_body, type_defs)`
//! puts type declarations beside `main`; 835 cases use it. Those arrive in
//! `case.prelude`.
//!
//! **Output is collected, not paired.** Every `System.out.println(x)` is
//! rewritten to `__p(x)`, which appends to a buffer the harness compares once
//! at the end. Pairing the i-th print with the i-th expected line cannot
//! assert anything about a loop, and loops were 659 of 7,395 cases.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing }

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("// vybe-test: {slug}\n// origin: {origin}\n");
    let members = harness_members(harness);

    let Some(expected) = case.expected.as_ref() else {
        return Emitted {
            text: format!(
                "{header}// vybe-test-mode: compile\n\n{}\n",
                assemble(&case.source, case.prelude.as_deref(), "")
            ),
            pairing: Pairing::Direct };
    };

    // Output is COLLECTED, not paired: every print is rewritten to append to
    // the harness buffer and the whole thing is compared once. Pairing the
    // i-th print with the i-th expected line cannot assert anything about a
    // loop, which was 659 of 7,395 cases — the single largest gap.
    let (body, rewritten) = rewrite_prints(&case.source);
    if let Some(reason) = unpairable(&case.source, rewritten) {
        return Emitted {
            text: format!("{header}\n{}\n", assemble(&case.source, case.prelude.as_deref(), "")),
            pairing: Pairing::Unpairable(reason) };
    }

    let want = java_string(&expected.join("\n"));
    let body = format!("{body}\n__check({want});");

    Emitted {
        text: format!(
            "{header}\n{}\n",
            assemble(&body, case.prelude.as_deref(), &members)
        ),
        pairing: Pairing::Direct }
}

/// `System.out.println(x)` → `__p(x)`, `System.out.print(x)` → `__pr(x)`.
/// Returns the rewritten source and how many calls were redirected.
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
        // Longest first: `println(` also starts with `print`.
        let hit = ["System.out.println(", "System.out.print("]
            .into_iter()
            .find(|n| src.is_char_boundary(i) && src[i..].starts_with(n));
        if let Some(needle) = hit {
            let args_start = i + needle.len();
            if let Some(end) = close_paren(bytes, args_start) {
                let args = src[args_start..end].trim();
                let target = if needle.ends_with("println(") { "__p" } else { "__pr" };
                // `println()` with no argument writes a bare newline; `__p()`
                // would not compile.
                let args = if args.is_empty() { "\"\"" } else { args };
                out.push_str(&format!("{target}({args})"));
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

fn unpairable(src: &str, rewritten: usize) -> Option<String> {
    if rewritten == 0 {
        return Some("no print call to collect".into());
    }
    // `printf` takes a format string; its output is not one append per call.
    if src.contains("System.out.printf") {
        return Some("printf — format output is not collectable".into());
    }
    // A program that hands `System.out` to something else writes past the
    // buffer, so the collected output would be short and the test would fail
    // for the wrong reason.
    if src.contains("System.out.flush") || src.contains("System.out)") {
        return Some("passes or flushes System.out directly".into());
    }
    None
}

/// Build the file. A source that already declares a class is a whole program
/// (`run_prints`); anything else is a `main` body that needs wrapping.
fn assemble(body: &str, prelude: Option<&str>, members: &str) -> String {
    let body = body.trim();
    if declares_class(body) {
        return splice_into_class(body, members);
    }
    let types = prelude.map(str::trim).unwrap_or("");
    let mut out = String::from("public class Main {\n");
    if !members.is_empty() {
        out.push_str(members);
        out.push_str("\n\n");
    }
    if !types.is_empty() {
        out.push_str(types);
        out.push('\n');
    }
    out.push_str("    public static void main(String[] args) {\n");
    out.push_str(body);
    out.push_str("\n    }\n}\n");
    out
}

/// Put the harness members just inside the first class body.
fn splice_into_class(src: &str, members: &str) -> String {
    if members.is_empty() {
        return src.to_string();
    }
    match src.find('{') {
        Some(at) => format!("{}{{\n{members}\n{}", &src[..at], &src[at + 1..]),
        None => format!("{src}\n{members}\n") }
}

/// `class Foo` / `interface Foo` at the start — the source is a full program.
fn declares_class(src: &str) -> bool {
    let head = src.trim_start();
    head.starts_with("public class")
        || head.starts_with("class ")
        || head.starts_with("public interface")
        || head.starts_with("abstract class")
        || head.starts_with("final class")
        || head.starts_with("public abstract class")
        || head.starts_with("public final class")
}

/// The harness file is a complete compilable class; the test needs its members.
fn harness_members(harness: &str) -> String {
    let Some(open) = harness.find('{') else {
        return harness.to_string();
    };
    let Some(close) = harness.rfind('}') else {
        return harness.to_string();
    };
    harness[open + 1..close].trim_end().to_string()
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

/// Java string and char literals, including text blocks (`"""…"""`).
fn skip_literal(bytes: &[u8], at: usize) -> Option<usize> {
    match bytes.get(at)? {
        b'"' => {
            let triple = bytes.get(at + 1) == Some(&b'"') && bytes.get(at + 2) == Some(&b'"');
            let mut i = at + if triple { 3 } else { 1 };
            while i < bytes.len() {
                if bytes[i] == b'\\' {
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
                    _ => i += 1 }
            }
            Some(bytes.len())
        }
        _ => None }
}

fn java_string(text: &str) -> String {
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
