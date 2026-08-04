//! VB emitter: one extracted case → a standalone `.vb` test.
//!
//! **Output is COLLECTED, not paired.** Every `Console.WriteLine(x)` becomes
//! `__P(CStr(x))`, appending to a buffer the harness compares once at the end
//! of `Sub Main`. Pairing the i-th print with the i-th expected line cannot
//! assert anything about a loop, and that alone left 402 of 6,671 cases
//! without an assertion. `Console.Write` (no newline) becomes `__Pr`.
//!
//! VB is line-oriented and has no statement separator to unfold, so nothing is
//! reflowed — each call is replaced in place.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing }

pub fn emit(case: &Case, origin: &str, slug: &str, harness: &str) -> Emitted {
    let header = format!("' vybe-test: {slug}\n' origin: {origin}\n");

    let Some(raw) = case.expected.as_ref() else {
        return Emitted {
            text: format!("{header}' vybe-test-mode: compile\n\n{}\n", case.source.trim()),
            pairing: Pairing::Direct };
    };

    // The `vb_*_spec!`/`vb_case!` macros pass expectations through
    // `dotnet_expected_one`, which rewrites the Rust-style `true`/`false` to
    // the `True`/`False` .NET actually prints.
    let expected: Vec<String> = raw
        .iter()
        .map(|e| match e.as_str() {
            "true" => "True".to_string(),
            "false" => "False".to_string(),
            other => other.to_string() })
        .collect();

    let (body, rewritten) = rewrite_prints(&case.source);
    if let Some(reason) = unpairable(&case.source, rewritten) {
        return Emitted {
            text: format!("{header}\n{}\n", case.source.trim()),
            pairing: Pairing::Unpairable(reason) };
    }

    let want = vb_string(&expected.join("\n"));
    let body = close_main_with(body.trim(), &format!("__Check({want})"));

    Emitted {
        text: format!("{header}\n{harness}\n\n{}\n", body.trim_end()),
        pairing: Pairing::Direct }
}

/// Put `__Check(…)` at the END of `Sub Main`, where every print has already
/// run. Appending it to the file instead would land outside the module.
fn close_main_with(src: &str, call: &str) -> String {
    let lines: Vec<&str> = src.lines().collect();
    let Some(start) = lines.iter().position(|l| {
        let t = l.trim().to_ascii_lowercase();
        t.starts_with("sub main(") || t == "sub main"
    }) else {
        // No `Sub Main` to close — 199 of 6,671 cases. Leave the file as it is
        // rather than inventing a place for the check to live.
        return src.to_string();
    };

    let mut depth = 1usize;
    for (n, line) in lines.iter().enumerate().skip(start + 1) {
        let t = line.trim().to_ascii_lowercase();
        if t.starts_with("end sub") || t.starts_with("end function") {
            depth -= 1;
            if depth == 0 {
                let indent = " ".repeat(line.len() - line.trim_start().len() + 4);
                let mut out: Vec<String> = lines[..n].iter().map(|l| l.to_string()).collect();
                out.push(format!("{indent}{call}"));
                out.extend(lines[n..].iter().map(|l| l.to_string()));
                return out.join("\n");
            }
            continue;
        }
        // A one-line lambda opens no block. Only a bare declaration does.
        if t.starts_with("sub ") || t.starts_with("function ") {
            depth += 1;
        }
    }
    src.to_string()
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
                // Render HERE, where the expression still has its static type.
                // `WriteLine()` with no argument writes a bare newline.
                let rendered =
                    if args.is_empty() { "\"\"".to_string() } else { format!("CStr({args})") };
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
        _ => None }
}

/// VB keywords are case-insensitive.
/// VB has no backslash escapes; a quote is doubled.
fn vb_string(text: &str) -> String {
    format!("\"{}\"", text.replace('"', "\"\""))
}
