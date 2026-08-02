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

    // The i-th expected LINE belongs to the print that runs i-th, which is not
    // the i-th print in the source once `defer` is involved.
    let order = runtime_order(&case.source, &prints);

    // Rewrite back-to-front so the earlier spans keep their byte offsets.
    let mut pairing: Vec<Option<&String>> = vec![None; prints.len()];
    for (line, &print_index) in order.iter().enumerate() {
        pairing[print_index] = expected.get(line);
    }
    let mut body = case.source.clone();
    for (i, span) in prints.iter().enumerate().rev() {
        let Some(want) = pairing[i] else { continue };
        let args = &case.source[span.args_start..span.args_end];
        let replacement =
            format!("__check({}, {})", render_println_args(args), go_string(want));
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
    // A goroutine's output order is genuinely not knowable from the source.
    // `defer` is — but only in `main`; see below.
    if has_keyword(src, "go") {
        return Some("goroutine — output order is not source order".into());
    }
    // A `defer` runs when ITS function returns, not when the program ends. In
    // `main` those are the same moment, so the LIFO reordering in
    // `runtime_order` is exact. Anywhere else the deferred output lands
    // mid-stream at a point only real control-flow analysis would find.
    if !defer_spans(src).is_empty() {
        if !defers_are_all_top_level_in_main(src) {
            return Some("defer outside main — output lands mid-stream".into());
        }
        // `defer f()` where `f` is a named function that prints: the print sits
        // in `f`'s body, lexically outside any defer, so the reordering below
        // cannot see that it runs late.
        if let Some(body) = main_body_span(src) {
            if prints.iter().any(|p| p.start < body.0 || p.start >= body.1) {
                return Some("defer plus a printing helper — order needs call analysis".into());
            }
        }
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

/// Print indices in the order they actually produce output.
///
/// `defer` runs its statements LIFO when the function returns, so
/// `defer Println(3); defer Println(2); defer Println(1)` prints `1, 2, 3` —
/// the reverse of the source. Pairing the i-th print with the i-th expected
/// line without accounting for that mis-assigned 37 Go tests.
///
/// Deferred calls run in reverse registration order, and the prints *within*
/// one deferred closure still run in source order — so the result is: every
/// non-deferred print in source order, then each `defer` group in reverse
/// registration order.
fn runtime_order(src: &str, prints: &[Span]) -> Vec<usize> {
    let defers = defer_spans(src);
    let group_of = |span: &Span| -> Option<usize> {
        defers
            .iter()
            .position(|(start, end)| span.start >= *start && span.start < *end)
    };

    let mut immediate = Vec::new();
    let mut deferred: Vec<(usize, usize)> = Vec::new(); // (defer index, print index)
    for (i, span) in prints.iter().enumerate() {
        match group_of(span) {
            Some(group) => deferred.push((group, i)),
            None => immediate.push(i),
        }
    }
    // Later `defer`s run first; inside one, source order holds.
    deferred.sort_by(|a, b| b.0.cmp(&a.0).then(a.1.cmp(&b.1)));
    immediate.extend(deferred.into_iter().map(|(_, i)| i));
    immediate
}

/// Whether every `defer` sits directly in `main`'s body.
///
/// Directly: not inside a nested `func` literal that `main` *calls*, whose
/// defers would fire at that call rather than at exit. The deferred expression
/// itself may of course be a closure — that is `defer func() { … }()`.
fn defers_are_all_top_level_in_main(src: &str) -> bool {
    let bytes = src.as_bytes();
    let Some((body, _)) = main_body_span(src) else {
        return false;
    };

    let defers = defer_spans(src);
    let mut depth = 0i32;
    let mut i = body;
    let mut next_defer = 0usize;

    while i < bytes.len() && next_defer < defers.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        // A defer's own body may nest freely; step over it whole.
        if i == defers[next_defer].0 {
            if depth != 1 {
                return false;
            }
            i = defers[next_defer].1;
            next_defer += 1;
            continue;
        }
        match bytes[i] {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    // main closed with defers still unaccounted for.
                    return false;
                }
            }
            _ => {}
        }
        i += 1;
    }
    next_defer == defers.len()
}

/// Byte range of `main`'s body, from its opening brace to its closing one.
fn main_body_span(src: &str) -> Option<(usize, usize)> {
    let bytes = src.as_bytes();
    let main_at = src.find("func main(")?;
    let open = src[main_at..].find('{').map(|o| main_at + o)?;

    let mut depth = 0i32;
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
                    return Some((open, i + 1));
                }
            }
            _ => {}
        }
        i += 1;
    }
    None
}

/// Byte range of each `defer` statement, in registration order.
fn defer_spans(src: &str) -> Vec<(usize, usize)> {
    let bytes = src.as_bytes();
    let mut spans = Vec::new();
    let mut i = 0usize;

    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if !src.is_char_boundary(i) || !src[i..].starts_with("defer")
            || is_ident(if i == 0 { b' ' } else { bytes[i - 1] })
            || is_ident(bytes.get(i + 5).copied().unwrap_or(b' '))
        {
            i += 1;
            continue;
        }
        match defer_end(src, bytes, i + 5) {
            Some(end) => {
                spans.push((i, end));
                i = end;
            }
            None => i += 5,
        }
    }
    spans
}

/// End of the deferred expression starting at `from`.
///
/// `defer fmt.Println(x)` ends at the call's `)`. `defer func() { … }()` does
/// not: the first `)` closes the literal's parameter list, so a `{` or `(`
/// following it means the expression continues.
fn defer_end(src: &str, bytes: &[u8], from: usize) -> Option<usize> {
    let mut paren = 0usize;
    let mut brace = 0usize;
    let mut i = from;

    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        match bytes[i] {
            b'(' => paren += 1,
            b'{' => brace += 1,
            b'}' => brace = brace.saturating_sub(1),
            b')' => {
                paren -= 1;
                if paren == 0 && brace == 0 {
                    let rest = src[i + 1..].trim_start();
                    if rest.starts_with('{') || rest.starts_with('(') {
                        i += 1;
                        continue;
                    }
                    return Some(i + 1);
                }
            }
            _ => {}
        }
        i += 1;
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
        // Sources contain non-ASCII outside literals (an em dash in a comment,
        // for one), and slicing into the middle of a code point panics.
        if !src.is_char_boundary(i) {
            i += 1;
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

fn has_call(src: &str, name: &str) -> bool {
    let bytes = src.as_bytes();
    let needle = format!("{name}(");
    let mut i = 0usize;
    while i < bytes.len() {
        if let Some(next) = skip_literal(bytes, i) {
            i = next;
            continue;
        }
        if src.is_char_boundary(i) && src[i..].starts_with(&needle) {
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
        if !in_clause && paren == 0 && src.is_char_boundary(i) && starts_control_clause(src, bytes, i) {
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
    if !src.is_char_boundary(at) {
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

