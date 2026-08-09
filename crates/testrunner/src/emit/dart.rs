//! Dart emitter: one extracted case → a standalone `.dart` test.
//!
//! Sources are already complete programs with a `main`, so nothing is wrapped.
//! Every `print(x)` in the body becomes `__p(x)`, which appends to the
//! harness buffer, and the whole output is compared once — so a loop asserts
//! just like a straight-line program.
//!
//! `main` is renamed to `__vybeMain` and called from an `async` wrapper that
//! `await`s it. That covers `void main()`, `void main() async` and
//! `Future<void> main() async` with one shape, and avoids having to find the
//! closing brace of `main` to append the check.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
}

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("// vybe-test: {slug}\n// origin: {origin}\n");
    let body = case.source.trim();

    let Some(expected) = case.expected.as_ref() else {
        // `compile_ok` carried no expected output. Emit it as a RUN, not
        // compile-only: "runs to completion" subsumes "the frontend accepted
        // it". Those sources are bare declarations with no `main`, and while
        // `vybex -d` does not need one, `dart run` refuses outright —
        // "Invoked Dart programs must have a 'main' function defined" — so the
        // differential was comparing a compile against a run.
        let entry = if find_main_decl(body).is_some() {
            String::new()
        } else {
            "\n\nvoid main() {}\n".to_string()
        };
        return Emitted {
            text: format!("{header}\n{body}{entry}"),
            pairing: Pairing::Direct,
        };
    };

    let is_async = main_is_async(body);
    let Some(renamed) = rename_main(body) else {
        return Emitted {
            text: format!("{header}\n{body}\n"),
            pairing: Pairing::Unpairable("no `main` to wrap".into()),
        };
    };
    if let Some(reason) = unpairable(body) {
        return Emitted {
            text: format!("{header}\n{body}\n"),
            pairing: Pairing::Unpairable(reason),
        };
    }

    let collected = rewrite_prints(&renamed);
    let want = dart_string(&expected.join("\n"));
    // `await` only when `__vybeMain` actually returns a Future. Dart rejects
    // `await` on a `void` expression outright — "This expression has type
    // 'void' and can't be used" — so one shape does NOT cover both, even
    // though Vybe accepts it. Caught by the real SDK, not by us.
    let wrapper = if is_async {
        format!("Future<void> main() async {{\n  await __vybeMain();\n  __check({want});\n}}")
    } else {
        format!("void main() {{\n  __vybeMain();\n  __check({want});\n}}")
    };
    Emitted {
        text: format!("{header}\n{harness}\n\n{collected}\n\n{wrapper}\n"),
        pairing: Pairing::Direct,
    }
}

/// Whether `main` returns a Future — either declared `async` or with a
/// `Future` return type.
fn main_is_async(src: &str) -> bool {
    let Some(at) = find_main_decl(src) else {
        return false;
    };
    let line_start = src[..at].rfind('\n').map(|i| i + 1).unwrap_or(0);
    let head_end = src[at..].find('{').map(|o| at + o).unwrap_or(src.len());
    let head = &src[line_start..head_end];
    head.contains("async") || head.contains("Future")
}

fn find_main_decl(src: &str) -> Option<usize> {
    let bytes = src.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i) && src[i..].starts_with("main") {
            let before = if i == 0 { b' ' } else { bytes[i - 1] };
            let after = bytes.get(i + 4).copied().unwrap_or(b' ');
            if !is_ident(before) && (after == b'(' || after == b' ') {
                return Some(i);
            }
        }
        i += 1;
    }
    None
}

/// `void main()` → `void __vybeMain()`, preserving the return type and any
/// `async`. Only the DECLARATION is renamed; a recursive call to `main` inside
/// the body would be renamed too, which is what we want.
fn rename_main(src: &str) -> Option<String> {
    let bytes = src.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i) && src[i..].starts_with("main") {
            let before = if i == 0 { b' ' } else { bytes[i - 1] };
            let after = bytes.get(i + 4).copied().unwrap_or(b' ');
            // A declaration, not `domain` or `mainWindow`.
            if !is_ident(before) && (after == b'(' || after == b' ') {
                let mut out = String::with_capacity(src.len() + 8);
                out.push_str(&src[..i]);
                out.push_str("__vybeMain");
                out.push_str(&src[i + 4..]);
                return Some(out);
            }
        }
        i += 1;
    }
    None
}

/// `print(` → `__p(`, outside literals.
fn rewrite_prints(src: &str) -> String {
    let bytes = src.as_bytes();
    let mut out = String::with_capacity(src.len());
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            out.push_str(&src[i..next]);
            i = next;
            continue;
        }
        if src.is_char_boundary(i)
            && src[i..].starts_with("print(")
            && !is_ident(if i == 0 { b' ' } else { bytes[i - 1] })
        {
            out.push_str("__p(");
            i += "print(".len();
            continue;
        }
        let ch = src[i..].chars().next().unwrap_or(' ');
        out.push(ch);
        i += ch.len_utf8();
    }
    out
}

fn unpairable(src: &str) -> Option<String> {
    // `stdout.write` bypasses the buffer, so the collected output would be
    // short and the test would fail for the wrong reason.
    if src.contains("stdout.") || src.contains("stderr.") {
        return Some("writes to stdout/stderr directly — bypasses the buffer".into());
    }
    // A program that redefines `print` shadows the rewrite target.
    if src.contains("void print(") {
        return Some("defines its own `print`".into());
    }
    None
}

fn is_ident(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_' || byte == b'.' || byte == b'$'
}

/// Dart string and char literals, including raw (`r'…'`) and triple-quoted.
fn skip_literal(bytes: &[u8], at: usize) -> Option<usize> {
    let (quote, mut i) = match bytes.get(at)? {
        b'"' => (b'"', at),
        b'\'' => (b'\'', at),
        b'r' if matches!(bytes.get(at + 1), Some(b'"') | Some(b'\'')) => (bytes[at + 1], at + 1),
        _ => return None,
    };
    let triple = bytes.get(i + 1) == Some(&quote) && bytes.get(i + 2) == Some(&quote);
    i += if triple { 3 } else { 1 };
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
        }
        i += 1;
    }
    Some(bytes.len())
}

fn dart_string(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 2);
    out.push('\'');
    for ch in text.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '\'' => out.push_str("\\'"),
            // `$` starts an interpolation in a non-raw Dart string.
            '$' => out.push_str("\\$"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(ch),
        }
    }
    out.push('\'');
    out
}
