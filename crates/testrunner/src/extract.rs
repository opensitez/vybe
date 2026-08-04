//! Pull test cases out of the native Rust test files.
//!
//! Every case is (name, source, expectation). The expectation is either a list
//! of stdout lines or `None`, meaning the case only ever asserted that the
//! frontend accepts the program — 3,571 of those exist and they carry no
//! output at all, so they become a compile-mode directive rather than an
//! assertion.

use crate::rustlit;

#[derive(Debug, Clone)]
pub struct Case {
    pub name: String,
    pub source: String,
    /// `None` for compile-only cases.
    pub expected: Option<Vec<String>>,
    /// True when the case asserted the frontend *rejects* the program.
    pub expect_failure: bool,
    /// The helper compared one value rather than a line vector, and what that
    /// value means (first line vs all lines joined) varies by module.
    pub single_line: bool,
    /// Run it and take the exit code; there is no output to pair. A `.wast`
    /// script carries its own `assert_return` directives, so executing it IS
    /// the assertion.
    pub run_only: bool,
    /// A SECOND source the helper wraps around the first. Java's
    /// `run_in_main(main_body, type_defs)` puts type declarations beside
    /// `main` inside `Main`; 835 cases use it, and 780 of those pass the
    /// declarations as a local `let types = r#"…"#` rather than inline, so it
    /// has to be resolved from the enclosing test fn or the emitted file is
    /// missing every type it references.
    pub prelude: Option<String> }

#[derive(Clone, Copy)]
enum Shape {
    /// `name => ("src", vec!["line"]),`
    Run,
    /// `name => "src",`
    Compile,
    /// `name => "src",` but the frontend must reject it.
    CompileFail,
    /// `name => { "src", ["line"] };` — the JS batch spelling.
    RunBraced,
    /// `name => { includes: [...], declarations: "…", body: "…", expect: [...] }`
    /// — C's spelling, where the program is assembled from three parts rather
    /// than given as one source.
    CFields,
    /// The same fields with no `expect:` — `c_compile_cases!`, whose whole
    /// assertion is `compile_ok`. Parsed identically, emitted as compile mode.
    CFieldsCompile }

fn shape_of(macro_name: &str) -> Option<Shape> {
    Some(match macro_name {
        "go_run_cases" | "run_cases" | "kotlin_run_cases" => Shape::Run,
        // `dart_cases!` / `fortran_cases!` are `name => { src, [expected] };`
        // — the same shape, `;`-separated, which the entry loop already takes.
        // Being absent here is why 91 dart modules (4,134 tests) and 43 fortran
        // ones (2,070) produced no files at all: an unknown macro name is
        // skipped silently, so the gap only shows against the cargo log.
        | "dart_cases" | "fortran_cases"
        | "js_cases" | "js_import_cases" | "php_cases" | "csharp_cases" | "wat_exec"
        // `lua_print! { name => { "src", "expected" } }` — same braced shape,
        // with a bare-string expectation rather than a list.
        | "lua_print" => Shape::RunBraced,
        "go_compile_cases" | "compile_cases" => Shape::Compile,
        "go_compile_fail_cases" => Shape::CompileFail,
        "c_cases" | "c_run_cases" => Shape::CFields,
        // Same braced fields, minus `expect:` — the case asserted only that the
        // frontend accepts the program. 647 of C's tests are this one macro,
        // and an unknown macro name is skipped in silence, so the gap showed
        // only against the cargo log: 7,512 there against 6,865 files.
        "c_compile_cases" => Shape::CFieldsCompile,
        _ => return None })
}

/// Every case in one `.rs` test module.
pub fn cases_in_file(text: &str) -> anyhow::Result<Vec<Case>> {
    let src = text.as_bytes();
    let mut cases = Vec::new();
    let mut at = 0usize;

    while at < src.len() {
        let Some((name, body_start)) = next_macro_header(src, at) else {
            break;
        };
        let Some(shape) = shape_of(&name) else {
            at = body_start;
            continue;
        };
        let body_end = matching_brace(src, body_start - 1)?;
        cases.extend(parse_entries(src, body_start, body_end, shape)?);
        at = body_end;
    }
    Ok(cases)
}

/// Find the next `<ident>! {` and return the macro's bare name plus the index
/// just past its opening brace. `crate::php_cases! {` yields `php_cases`.
fn next_macro_header(src: &[u8], from: usize) -> Option<(String, usize)> {
    let mut i = from;
    while i < src.len() {
        if src[i] != b'!' {
            i += 1;
            continue;
        }
        let after = rustlit::skip_trivia(src, i + 1);
        if src.get(after) != Some(&b'{') {
            i += 1;
            continue;
        }
        let mut start = i;
        while start > 0 && (src[start - 1].is_ascii_alphanumeric() || src[start - 1] == b'_') {
            start -= 1;
        }
        if start == i {
            i += 1;
            continue;
        }
        let name = String::from_utf8_lossy(&src[start..i]).into_owned();
        return Some((name, after + 1));
    }
    None
}

/// Index just past the `}` matching the `{` at `open`. String literals are
/// skipped whole so a brace inside a test body cannot close the macro.
fn matching_brace(src: &[u8], open: usize) -> anyhow::Result<usize> {
    let mut depth = 0usize;
    let mut i = open;
    while i < src.len() {
        // Comments can contain braces — one `// … extra `}` made it …` closed a
        // whole `js_cases!` block early and silently dropped the 21 entries
        // after it. Skip them before counting.
        if src[i] == b'/' {
            match src.get(i + 1) {
                Some(b'/') => {
                    while i < src.len() && src[i] != b'\n' {
                        i += 1;
                    }
                    continue;
                }
                Some(b'*') => {
                    i += 2;
                    while i + 1 < src.len() && &src[i..i + 2] != b"*/" {
                        i += 1;
                    }
                    i = (i + 2).min(src.len());
                    continue;
                }
                _ => {}
            }
        }
        match src[i] {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    return Ok(i + 1);
                }
            }
            b'"' => {
                let (_, next) = rustlit::scan(src, i)?;
                i = next;
                continue;
            }
            b'r' if matches!(src.get(i + 1), Some(b'"' | b'#')) => {
                let (_, next) = rustlit::scan(src, i)?;
                i = next;
                continue;
            }
            b'\'' => {
                // A char literal, or a lifetime. Only the former can hide a brace.
                if src.get(i + 2) == Some(&b'\'') || src.get(i + 3) == Some(&b'\'') {
                    i += 2;
                }
            }
            _ => {}
        }
        i += 1;
    }
    anyhow::bail!("unbalanced braces starting at byte {open}")
}

fn parse_entries(
    src: &[u8],
    mut at: usize,
    end: usize,
    shape: Shape,
) -> anyhow::Result<Vec<Case>> {
    let mut cases = Vec::new();
    loop {
        at = rustlit::skip_trivia(src, at);
        if at >= end || src[at] == b'}' {
            return Ok(cases);
        }

        let name_start = at;
        while at < end && (src[at].is_ascii_alphanumeric() || src[at] == b'_') {
            at += 1;
        }
        if at == name_start {
            anyhow::bail!("expected a case name at byte {at}");
        }
        let name = String::from_utf8_lossy(&src[name_start..at]).into_owned();

        at = rustlit::skip_trivia(src, at);
        if !src[at..].starts_with(b"=>") {
            anyhow::bail!("expected `=>` after case `{name}`");
        }
        at = rustlit::skip_trivia(src, at + 2);

        let case = match shape {
            // `{ includes: [...], decls|declarations: "…", body: "…",
            //    expect: [...] }` — order is not fixed and `includes` is
            //   optional, so read by KEY rather than by position.
            Shape::CFields | Shape::CFieldsCompile => {
                if src[at] != b'{' {
                    anyhow::bail!("expected `{{` in C case `{name}`");
                }
                let close = matching_brace(src, at)?;
                let mut includes: Vec<String> = Vec::new();
                let mut decls = String::new();
                let mut body = String::new();
                let mut expected: Vec<String> = Vec::new();
                let mut i = rustlit::skip_trivia(src, at + 1);
                while i < close {
                    let key_start = i;
                    while i < close && (src[i].is_ascii_alphanumeric() || src[i] == b'_') {
                        i += 1;
                    }
                    if i == key_start {
                        i = rustlit::skip_trivia(src, i + 1);
                        continue;
                    }
                    let key = String::from_utf8_lossy(&src[key_start..i]).into_owned();
                    i = rustlit::skip_trivia(src, i);
                    if src.get(i) != Some(&b':') {
                        continue;
                    }
                    i = rustlit::skip_trivia(src, i + 1);
                    match key.as_str() {
                        "includes" => {
                            let (list, next) = scan_expected(src, i)?;
                            includes = list;
                            i = next;
                        }
                        "expect" => {
                            let (list, next) = scan_expected(src, i)?;
                            expected = list;
                            i = next;
                        }
                        "decls" | "declarations" => {
                            let (text, next) = rustlit::scan(src, i)?;
                            decls = text;
                            i = next;
                        }
                        "body" => {
                            let (text, next) = rustlit::scan(src, i)?;
                            body = text;
                            i = next;
                        }
                        _ => {
                            i = rustlit::skip_trivia(src, i + 1);
                            continue;
                        }
                    }
                    i = rustlit::skip_trivia(src, i);
                    if src.get(i) == Some(&b',') {
                        i = rustlit::skip_trivia(src, i + 1);
                    }
                }
                at = close + 1;
                // The includes belong above the declarations, and both go
                // above `main` — so they travel together in `prelude`.
                let mut head = String::new();
                for inc in &includes {
                    head.push_str("#include ");
                    head.push_str(inc);
                    head.push('\n');
                }
                head.push_str(&decls);
                Case {
                    name,
                    source: body,
                    // `None` is what routes the case to compile mode. It has to
                    // come from the MACRO, not from an empty `expect:` list — a
                    // run case that legitimately prints nothing is a different
                    // thing from one that was never meant to run.
                    expected: match shape {
                        Shape::CFieldsCompile => None,
                        _ => Some(expected) },
                    expect_failure: false,
                    single_line: false,
                    run_only: false,
                    prelude: Some(head) }
            }
            Shape::Run => {
                if src[at] != b'(' {
                    anyhow::bail!("expected `(` in run case `{name}`");
                }
                at = rustlit::skip_trivia(src, at + 1);
                let (source, next) = rustlit::scan(src, at)?;
                at = rustlit::skip_trivia(src, next);
                if src[at] != b',' {
                    anyhow::bail!("expected `,` after the source of `{name}`");
                }
                at = rustlit::skip_trivia(src, at + 1);
                let (expected, next) = scan_expected(src, at)?;
                at = rustlit::skip_trivia(src, next);
                if src[at] != b')' {
                    anyhow::bail!("expected `)` closing run case `{name}`");
                }
                at += 1;
                Case { name, source, expected: Some(expected), expect_failure: false, single_line: false , run_only: false, prelude: None }
            }
            Shape::RunBraced => {
                if src[at] != b'{' {
                    anyhow::bail!("expected `{{` in run case `{name}`");
                }
                at = rustlit::skip_trivia(src, at + 1);
                let (source, next) = rustlit::scan(src, at)?;
                at = rustlit::skip_trivia(src, next);
                if src[at] != b',' {
                    anyhow::bail!("expected `,` after the source of `{name}`");
                }
                at = rustlit::skip_trivia(src, at + 1);
                let (expected, next) = scan_expected(src, at)?;
                at = rustlit::skip_trivia(src, next);
                if src[at] != b'}' {
                    anyhow::bail!("expected `}}` closing run case `{name}`");
                }
                at += 1;
                Case { name, source, expected: Some(expected), expect_failure: false, single_line: false , run_only: false, prelude: None }
            }
            Shape::Compile | Shape::CompileFail => {
                let (source, next) = rustlit::scan(src, at)?;
                at = next;
                Case {
                    name,
                    source,
                    expected: None,
                    expect_failure: matches!(shape, Shape::CompileFail),
                    single_line: false,
        run_only: false,
        prelude: None }
            }
        };
        cases.push(case);

        at = rustlit::skip_trivia(src, at);
        if at < end && matches!(src[at], b',' | b';') {
            at += 1;
        }
    }
}

/// `vec!["a", "b"]` or `["a", "b"]` — the two spellings the corpus uses.
fn scan_expected(src: &[u8], mut at: usize) -> anyhow::Result<(Vec<String>, usize)> {
    // `&["a"]` — a slice reference is as common as `vec![…]` and `[…]`.
    if src.get(at) == Some(&b'&') {
        at = rustlit::skip_trivia(src, at + 1);
    }
    // A BARE literal is a one-line expectation: `wat_exec!` writes
    // `name => { src, "21" }` rather than a list, and refusing it made every
    // such module extract zero cases.
    if starts_string_literal(src, at) {
        let (line, next) = rustlit::scan(src, at)?;
        return Ok((vec![line], next));
    }
    if src[at..].starts_with(b"vec!") {
        at = rustlit::skip_trivia(src, at + 4);
    }
    if src[at] != b'[' {
        anyhow::bail!("expected an expectation list at byte {at}");
    }
    at = rustlit::skip_trivia(src, at + 1);

    let mut lines = Vec::new();
    while src[at] != b']' {
        // Elements may be wrapped: `vec![String::from("15"), …]`. Step over an
        // identifier path plus `(`, and close it after the literal.
        let mut wrapped = false;
        if !matches!(src.get(at), Some(b'"') | Some(b'r')) {
            let mut probe = at;
            while probe < src.len()
                && (src[probe].is_ascii_alphanumeric() || src[probe] == b'_' || src[probe] == b':')
            {
                probe += 1;
            }
            let open = rustlit::skip_trivia(src, probe);
            if probe > at && src.get(open) == Some(&b'(') {
                at = rustlit::skip_trivia(src, open + 1);
                wrapped = true;
            }
        }
        let (line, next) = rustlit::scan(src, at)?;
        lines.push(line);
        at = rustlit::skip_trivia(src, next);
        if wrapped && src.get(at) == Some(&b')') {
            at = rustlit::skip_trivia(src, at + 1);
        }
        // `vec!["a".to_string(), "b".into()]` — the conversion suffix is not
        // part of the value, and choking on it dropped 66 tests in one module.
        for suffix in [".to_string()", ".into()", ".to_owned()"] {
            if src[at..].starts_with(suffix.as_bytes()) {
                at = rustlit::skip_trivia(src, at + suffix.len());
                break;
            }
        }
        if src[at] == b',' {
            at = rustlit::skip_trivia(src, at + 1);
        }
    }
    Ok((lines, at + 1))
}

// ── `#[test] fn` shape ──────────────────────────────────────────────────────
//
// The macro batches cover go/c/lua/dart/wast. Everything else — js, php,
// python, vb, pascal, java, cobol, ruby, ~45,000 cases — writes one `#[test]`
// function per case:
//
//     #[test]
//     fn test_class_method() {
//         let code = r#" … "#;
//         assert_eq!(run_js_one(code), "7");
//     }

/// `helper("source", ["expected", …])` — one call carrying both halves.
fn two_arg_call(name: &str, body: &str) -> Option<Case> {
    let bytes = body.as_bytes();
    let mut i = 0usize;

    while i < bytes.len() {
        if bytes[i] != b'(' {
            i += 1;
            continue;
        }
        // Must be preceded by an identifier (optionally a `!` for a macro).
        let mut start = i;
        if start > 0 && bytes[start - 1] == b'!' {
            start -= 1;
        }
        let end = start;
        while start > 0 && (bytes[start - 1].is_ascii_alphanumeric() || bytes[start - 1] == b'_') {
            start -= 1;
        }
        if start == end {
            i += 1;
            continue;
        }

        let arg = rustlit::skip_trivia(bytes, i + 1);
        if !starts_string_literal(bytes, arg) {
            i += 1;
            continue;
        }
        let Ok((source, next)) = rustlit::scan(bytes, arg) else {
            i += 1;
            continue;
        };
        let after = rustlit::skip_trivia(bytes, next);
        if bytes.get(after) != Some(&b',') {
            i += 1;
            continue;
        }
        let mut list = rustlit::skip_trivia(bytes, after + 1);
        if bytes.get(list) == Some(&b'&') {
            list = rustlit::skip_trivia(bytes, list + 1);
        }
        let Ok((expected, _)) = scan_expected(bytes, list) else {
            i += 1;
            continue;
        };
        return Some(Case {
            name: name.to_string(),
            source,
            expected: Some(expected),
            expect_failure: false,
            single_line: false,
            run_only: false,
            prelude: None });
    }
    None
}

/// `v128_eq(body, "…")`, `i32_eq(body, 42)`, `f32x4_eq(body, [l0,l1,l2,l3])`,
/// `f64x2_eq(body, [l0,l1])` — the SIMD suite's own wrappers, rebuilt so each
/// emitted file is a complete script rather than a loose function body.
fn simd_wrapper_call(name: &str, body: &str) -> Option<Case> {
    let bytes = body.as_bytes();
    for (helper, result_ty, lanes) in [
        ("v128_eq", "v128", 0usize),
        ("i32_eq", "i32", 0),
        ("f32x4_eq", "f32", 4),
        ("f64x2_eq", "f64", 2),
    ] {
        let Some(at) = body.find(&format!("{helper}(")) else {
            continue;
        };
        let arg = rustlit::skip_trivia(bytes, at + helper.len() + 1);
        if !starts_string_literal(bytes, arg) {
            continue;
        }
        let Ok((func_body, next)) = rustlit::scan(bytes, arg) else {
            continue;
        };
        let after = rustlit::skip_trivia(bytes, next);
        if bytes.get(after) != Some(&b',') {
            continue;
        }
        let value_at = rustlit::skip_trivia(bytes, after + 1);

        // Per-lane form: one exported function and one assertion per lane.
        if lanes > 0 {
            let (expected, _) = scan_expected(bytes, value_at).ok()?;
            if expected.len() != lanes {
                continue;
            }
            let mut funcs = String::new();
            let mut asserts = String::new();
            for (i, lane) in expected.iter().enumerate() {
                funcs.push_str(&format!(
                    "  (func (export \"f{i}\") (result {result_ty})\n{func_body}\n  {result_ty}x{lanes}.extract_lane {i})\n"
                ));
                asserts.push_str(&format!("(assert_return (invoke \"f{i}\") ({lane}))\n"));
            }
            return Some(Case {
                name: name.to_string(),
                source: format!("(module\n{funcs})\n{asserts}"),
                expected: None,
                expect_failure: false,
                single_line: false,
                run_only: true,
                prelude: None });
        }

        // Single-result form: the expectation is a literal or a bare number.
        let expected = if starts_string_literal(bytes, value_at) {
            let (text, _) = rustlit::scan(bytes, value_at).ok()?;
            text
        } else {
            let mut end = value_at;
            if bytes.get(end) == Some(&b'-') {
                end += 1;
            }
            while end < bytes.len() && bytes[end].is_ascii_digit() {
                end += 1;
            }
            if end == value_at {
                continue;
            }
            format!("{result_ty}.const {}", &body[value_at..end])
        };
        return Some(Case {
            name: name.to_string(),
            source: format!(
                "(module (func (export \"f\") (result {result_ty})\n{func_body}))\n\
                 (assert_return (invoke \"f\") ({expected}))\n"
            ),
            expected: None,
            expect_failure: false,
            single_line: false,
            run_only: true,
            prelude: None });
    }
    None
}

/// `assert_eq!(lines[0], "a")` repeated once per expected line.
fn indexed_asserts(
    name: &str,
    body: &str,
    sources: &[(String, String)],
    results: &[(String, String)],
) -> Option<Case> {
    let bytes = body.as_bytes();
    let mut found: Vec<(usize, String)> = Vec::new();
    let mut call: Option<String> = None;
    let mut at = 0usize;

    while let Some(o) = body[at..].find("assert_eq!(") {
        let args_start = at + o + "assert_eq!(".len();
        at = args_start;
        let Some(args_end) = close_paren(bytes, args_start) else { break };
        let args = &body[args_start..args_end];
        let Some(comma) = top_level_comma(args) else { continue };
        let (lhs, rhs) = (args[..comma].trim(), args[comma + 1..].trim());

        let Some(open) = lhs.find('[') else { return None };
        let ident = lhs[..open].trim();
        // The closing bracket must come AFTER the opening one. Taking the
        // first `]` in the whole expression could land before `[` and panic
        // with "byte range starts at 44 but ends at 35".
        let close = open + 1 + lhs[open + 1..].find(']')?;
        let index: usize = lhs[open + 1..close].trim().parse().ok()?;
        let run = results.iter().find(|(n, _)| n == ident)?;
        call = Some(run.1.clone());

        found.push((index, parse_expected(rhs)?.into_iter().next()?));
    }

    // Only a run of lines starting at 0 can be paired positionally.
    found.sort_by_key(|(i, _)| *i);
    if found.is_empty() || found.iter().enumerate().any(|(n, (i, _))| n != *i) {
        return None;
    }

    let call = call?;
    let open = call.find('(')?;
    let inner = call[open + 1..call.rfind(')')?].trim();
    let (source, after_first) = if starts_string_literal(inner.as_bytes(), 0) {
        let (text, end) = rustlit::scan(inner.as_bytes(), 0).ok()?;
        (text, Some(end))
    } else {
        let ident = inner.trim_end_matches(".as_str()").trim_start_matches('&').trim();
        (sources.iter().find(|(n, _)| n == ident)?.1.clone(), None)
    };

    // A second argument, if the helper takes one. Either inline or an ident
    // bound earlier in the same fn.
    let prelude = after_first
        .and_then(|end| inner.get(end..))
        .map(str::trim)
        .and_then(|rest| rest.strip_prefix(','))
        .map(str::trim)
        .and_then(|arg| {
            let arg = arg.trim_end_matches(',').trim();
            if starts_string_literal(arg.as_bytes(), 0) {
                rustlit::scan(arg.as_bytes(), 0).ok().map(|(t, _)| t)
            } else {
                let ident = arg.trim_end_matches(".as_str()").trim_start_matches('&').trim();
                sources.iter().find(|(n, _)| n == ident).map(|(_, t)| t.clone())
            }
        });

    Some(Case {
        name: name.to_string(),
        source,
        expected: Some(found.into_iter().map(|(_, e)| e).collect()),
        expect_failure: false,
        single_line: false,
        run_only: false,
        prelude })
}

/// Whether a Rust string literal starts at `at`.
///
/// The `r` of a raw literal must be followed by `#` or `"` — otherwise it is
/// just an identifier that happens to start with r, and `run_js(...)` is
/// exactly that. Treating it as a literal made every `let out = run_js(…)`
/// module extract zero cases.
fn starts_string_literal(bytes: &[u8], at: usize) -> bool {
    match bytes.get(at) {
        Some(b'"') => true,
        Some(b'r') => matches!(bytes.get(at + 1), Some(b'"') | Some(b'#')),
        _ => false }
}

/// Every `#[test] fn` in a module that asserts on a runner helper's output.
pub fn test_fns_in_file(text: &str) -> Vec<Case> {
    let src = text.as_bytes();
    let mut cases = Vec::new();
    let mut at = 0usize;
    let consts = const_fns(text);

    while let Some(found) = text[at..].find("#[test]") {
        let start = at + found;
        at = start + "#[test]".len();

        // A COMMENTED-OUT test is not a test. `test_dart_apis.rs` parks one on
        // a `// #[test] fn pattern_matching() { … }` line; extracting it minted
        // a file for a case cargo does not run, which shows up as a suite that
        // has MORE tests than the corpus.
        let line_start = text[..start].rfind('\n').map(|o| o + 1).unwrap_or(0);
        if text[line_start..start].contains("//") {
            continue;
        }

        let Some(name_at) = text[at..].find("fn ") else { break };
        let name_start = at + name_at + 3;
        let mut name_end = name_start;
        while name_end < src.len() && (src[name_end].is_ascii_alphanumeric() || src[name_end] == b'_')
        {
            name_end += 1;
        }
        let name = text[name_start..name_end].to_string();

        let Some(brace) = text[name_end..].find('{').map(|o| name_end + o) else { break };
        let Ok(body_end) = matching_brace(src, brace) else { break };
        let body = &text[brace..body_end];
        at = body_end;

        if let Some(case) = case_from_body(name, body, &consts) {
            cases.push(case);
        }
    }
    cases
}

/// `&p(data, body)` — a local two-argument wrapper that builds the program
/// from two sources. Returns `(first, second)`.
///
/// COBOL's `p(working_storage, procedure_division)` is the case that forced
/// this: without it `two_arg_call` reads `p("01 S PIC X(20).", "ACCEPT …")` as
/// `helper(source, expected)` and takes the DATA DIVISION as the program, so
/// every emitted file put its declarations inside the PROCEDURE DIVISION and
/// no longer compiled.
fn wrapper_two_sources(text: &str, consts: &[(String, String)]) -> Option<(String, String)> {
    let t = text.trim().trim_start_matches('&').trim();
    let open = t.find('(')?;
    let name = t[..open].trim();
    if name.is_empty()
        || is_run_helper(name)
        || !name.chars().all(|c| c.is_alphanumeric() || c == '_')
    {
        return None;
    }
    let close = t.rfind(')')?;
    if close <= open {
        return None;
    }
    let args = t[open + 1..close].trim_start();
    let (first, end) = scan_source_arg(args, consts)?;
    let rest = args[end..].trim_start().trim_start_matches(',').trim_start();
    let (second, _) = scan_source_arg(rest, consts)?;
    Some((first, second))
}

/// One argument of a source-building wrapper: a literal, or a call to a local
/// constant function that returns one.
///
/// `compile_ok(&p(d(), "    COMPUTE R = FUNCTION SQRT(16)."))` is COBOL's
/// spelling for 94 of its tests — the DATA DIVISION is shared by a whole module
/// so it lives in `fn d() -> &'static str`. Requiring a literal in that slot
/// left `cics_full`, `intrinsics`, `enterprise` and `embedded_sql` extracting
/// almost nothing.
fn scan_source_arg(arg: &str, consts: &[(String, String)]) -> Option<(String, usize)> {
    if starts_string_literal(arg.as_bytes(), 0) {
        return rustlit::scan(arg.as_bytes(), 0).ok();
    }
    let open = arg.find('(')?;
    let name = arg[..open].trim();
    if name.is_empty() || !name.chars().all(|c| c.is_alphanumeric() || c == '_') {
        return None;
    }
    // Zero arguments only: anything else is a call this cannot evaluate.
    if !arg[open + 1..].trim_start().starts_with(')') {
        return None;
    }
    let end = open + 1 + arg[open + 1..].find(')')? + 1;
    let text = consts.iter().find(|(n, _)| n == name)?.1.clone();
    Some((text, end))
}

/// Module-level `fn <name>() -> &'static str { "…" }` — a shared source
/// fragment, referenced by call rather than by name.
fn const_fns(text: &str) -> Vec<(String, String)> {
    let bytes = text.as_bytes();
    let mut out = Vec::new();
    let mut at = 0usize;
    while let Some(found) = text[at..].find("fn ") {
        let start = at + found;
        at = start + 3;
        let mut end = at;
        while end < bytes.len() && (bytes[end].is_ascii_alphanumeric() || bytes[end] == b'_') {
            end += 1;
        }
        let name = text[at..end].to_string();
        // `fn d()` and nothing else — a parameter means the body is not a
        // constant.
        let Some(rest) = text[end..].strip_prefix("()") else { continue };
        let Some(brace) = rest.find('{') else { continue };
        // A return type may sit between; `-> String` and `-> &'static str`
        // both appear, and neither changes what the body is.
        if rest[..brace].contains(';') {
            continue;
        }
        let body_at = end + "()".len() + brace + 1;
        let value_at = rustlit::skip_trivia(bytes, body_at);
        if !starts_string_literal(bytes, value_at) {
            continue;
        }
        let Ok((value, after)) = rustlit::scan(bytes, value_at) else { continue };
        // The literal must BE the body: `{ "…" }`, not the first of several
        // statements.
        if bytes.get(rustlit::skip_trivia(bytes, after)) == Some(&b'}') {
            out.push((name, value));
        }
    }
    out
}

/// Is this a run helper, however it was imported?
///
/// COBOL writes `helpers::run_prints(src)` rather than importing the name, and
/// a bare `starts_with("run_")` says no to that — which is why 31 of its
/// modules (247 tests) extracted nothing at all. The call is the same call; the
/// path in front of it is a Rust import detail.
fn is_run_helper(name: &str) -> bool {
    name.rsplit("::").next().is_some_and(|n| n.trim().starts_with("run_"))
}

fn case_from_body(name: String, body: &str, consts: &[(String, String)]) -> Option<Case> {
    let bytes = body.as_bytes();

    // Locals, in two flavours the corpus mixes freely:
    //   let code = r#"…"#;          the program, bound then passed
    //   let out  = run_js(code);    the RESULT, bound then asserted on
    let mut sources: Vec<(String, String)> = Vec::new();
    let mut results: Vec<(String, String)> = Vec::new(); // ident -> "helper(arg)"
    let mut i = 0usize;
    while let Some(found) = body[i..].find("let ") {
        let after = i + found + 4;
        let mut end = after;
        while end < bytes.len() && (bytes[end].is_ascii_alphanumeric() || bytes[end] == b'_') {
            end += 1;
        }
        let ident = body[after..end].to_string();
        // Skip a type annotation: `let out: Vec<String> = …`.
        let mut eq = rustlit::skip_trivia(bytes, end);
        if bytes.get(eq) == Some(&b':') {
            match body[eq..].find('=') {
                Some(o) => eq += o,
                None => {
                    i = end;
                    continue;
                }
            }
        }
        if bytes.get(eq) == Some(&b'=') {
            let value_at = rustlit::skip_trivia(bytes, eq + 1);
            if starts_string_literal(bytes, value_at) {
                if let Ok((text, _)) = rustlit::scan(bytes, value_at) {
                    sources.push((ident, text));
                }
            } else if let Some(open) = body[value_at..].find('(') {
                let helper = body[value_at..value_at + open].trim();
                if is_run_helper(helper) {
                    if let Some(close) = close_paren(bytes, value_at + open + 1) {
                        results.push((ident, body[value_at..close + 1].to_string()));
                    }
                }
            }
        }
        i = end;
    }

    // Module-local SIMD helpers wrap a FUNCTION BODY into a self-verifying
    // script. Rebuild each wrapper or the emitted file is a fragment: a bare
    // `v128.const …` as a module field, which is 51 parse errors.
    if !body.contains("assert_eq!(")
        && let Some(case) = simd_wrapper_call(&name, body)
    {
        return Some(case);
    }

    // Any `helper("src", [expected])` call — `case!` in two JS modules,
    // `assert_php_output` / `assert_js` defined locally in others. Matching the
    // SHAPE rather than a list of names means a module that rolls its own
    // wrapper still extracts.
    // BEFORE `two_arg_call`: a compile-mode helper wrapping a two-source
    // builder — `compile_ok(&p(data, body))` — looks exactly like
    // `helper(source, expected)` to the generic two-argument matcher.
    if !body.contains("assert_eq!(") {
        for helper in ["compile_ok_check", "compile_err", "compile_ok", "parse_ok"] {
            let Some(at) = body.find(&format!("{helper}(")) else {
                continue;
            };
            let open = at + helper.len() + 1;
            let close = close_paren(bytes, open)?;
            let Some((first, second)) = wrapper_two_sources(&body[open..close], consts) else {
                break;
            };
            return Some(Case {
                name,
                source: second,
                expected: None,
                expect_failure: helper.ends_with("_err"),
                single_line: false,
                run_only: false,
                prelude: Some(first) });
        }
    }

    if !body.contains("assert_eq!(") {
        if let Some(case) = two_arg_call(&name, body) {
            return Some(case);
        }
    }

    // `#[test] fn x() { compile_ok("src"); }` — the frontend must accept it,
    // nothing more. Same assertion the `*_compile_cases!` batches make.
    if !body.contains("assert_eq!(") {
        // Longest first: `compile_ok(` also matches inside `compile_ok_check(`,
        // and the shorter match lands mid-identifier.
        // `parse_err`/`compile_err` assert the front-end REJECTS the source —
        // that is compile-fail mode, which already exists. Longest names first:
        // `parse_ok(` also matches inside `parse_ok_check(`.
        for helper in [
            "compile_ok_check",
            "parse_ok_check",
            "compile_err",
            "parse_err",
            "compile_ok",
            "parse_ok",
            // Whole self-verifying scripts: RUN them and take the exit code.
            "run_wast_asserts",
            "must_fail",
            "ok",
        ] {
            if let Some(at) = body.find(&format!("{helper}(")) {
                let open = at + helper.len() + 1;
                let arg = rustlit::skip_trivia(bytes, open);
                if starts_string_literal(bytes, arg) {
                    let (source, _) = rustlit::scan(bytes, arg).ok()?;
                    return Some(Case {
                        name,
                        source,
                        expected: None,
                        expect_failure: helper.ends_with("_err"),
                        single_line: false,
        run_only: false,
        prelude: None });
                }
            }
        }
        return None;
    }

    // `let lines = run_js(code); assert_eq!(lines[0], "a"); assert_eq!(lines[1], "b");`
    // — one assertion per output line, indexed. Collect them all.
    if let Some(case) = indexed_asserts(&name, body, &sources, &results) {
        return Some(case);
    }

    let assert_at = body.find("assert_eq!(")?;
    let args_start = assert_at + "assert_eq!(".len();
    let args_end = close_paren(bytes, args_start)?;
    let args = &body[args_start..args_end];

    let comma = top_level_comma(args)?;
    let (lhs, expected_expr) = (args[..comma].trim(), args[comma + 1..].trim());

    // The asserted value is either the run call itself or a local holding it.
    let call = if is_run_helper(lhs) {
        lhs.to_string()
    } else {
        let ident = lhs.trim_end_matches(".clone()").trim();
        results.iter().find(|(n, _)| n == ident)?.1.clone()
    };

    let open = call.find('(')?;
    let helper = call[..open].trim();
    if !is_run_helper(helper) {
        return None;
    }
    let inner = call[open + 1..call.rfind(')')?].trim();

    // `run_prints(&p(data, body))` — COBOL builds its program from a local
    // two-argument wrapper, so the run helper's argument is another CALL, not
    // a literal. Unwrap one level: first arg is the prelude (WORKING-STORAGE),
    // second is the source (PROCEDURE DIVISION). Without this only the handful
    // of cases that pass a bare literal extract at all.
    let unwrapped = inner.trim_start_matches('&').trim();
    if let Some(open) = unwrapped.find('(')
        && unwrapped[..open].trim().chars().all(|c| c.is_alphanumeric() || c == '_')
        && !unwrapped[..open].trim().is_empty()
        && !unwrapped[..open].trim().starts_with("run_")
        && let Some(close) = unwrapped.rfind(')')
        && close > open
    {
        // The wrapper call is usually written across several lines, so the
        // first argument does not start at index 0.
        let args = unwrapped[open + 1..close].trim_start();
        if starts_string_literal(args.as_bytes(), 0)
            && let Ok((first, end)) = rustlit::scan(args.as_bytes(), 0)
        {
            let rest = args[end..].trim_start().trim_start_matches(',').trim_start();
            if starts_string_literal(rest.as_bytes(), 0)
                && let Ok((second, _)) = rustlit::scan(rest.as_bytes(), 0)
            {
                let expected = parse_expected(expected_expr)?;
                return Some(Case {
                    name,
                    source: second,
                    expected: Some(expected),
                    expect_failure: false,
                    single_line: helper.ends_with("_one"),
                    run_only: false,
                    prelude: Some(first) });
            }
        }
    }

    let (source, after_first) = if starts_string_literal(inner.as_bytes(), 0) {
        let (text, end) = rustlit::scan(inner.as_bytes(), 0).ok()?;
        (text, Some(end))
    } else {
        let ident = inner.trim_end_matches(".as_str()").trim_start_matches('&').trim();
        (sources.iter().find(|(n, _)| n == ident)?.1.clone(), None)
    };

    // A second argument, if the helper takes one. Either inline or an ident
    // bound earlier in the same fn.
    let prelude = after_first
        .and_then(|end| inner.get(end..))
        .map(str::trim)
        .and_then(|rest| rest.strip_prefix(','))
        .map(str::trim)
        .and_then(|arg| {
            let arg = arg.trim_end_matches(',').trim();
            if starts_string_literal(arg.as_bytes(), 0) {
                rustlit::scan(arg.as_bytes(), 0).ok().map(|(t, _)| t)
            } else {
                let ident = arg.trim_end_matches(".as_str()").trim_start_matches('&').trim();
                sources.iter().find(|(n, _)| n == ident).map(|(_, t)| t.clone())
            }
        });

    let expected = parse_expected(expected_expr)?;
    Some(Case {
        name,
        source,
        expected: Some(expected),
        expect_failure: false,
        // A `_one` helper compares against a single value whose meaning varies
        // by module — first line in some, all lines joined in others. The
        // emitter refuses to pair it unless the program prints exactly once.
        single_line: helper.ends_with("_one"),
        run_only: false,
        prelude })
}

/// `vec!["a", "b"]`, `["a"]`, or a bare `"a"`.
fn parse_expected(expr: &str) -> Option<Vec<String>> {
    let expr = expr.trim().trim_end_matches(';').trim();
    // A COMMENT can sit between the source and the expectation, and the corpus
    // uses that spot to explain a surprising value:
    //
    //     assert_eq!(
    //         run_prints(r#"…"#),
    //         // Dart's FileSystemEntity.parent on root returns root itself.
    //         vec!["/"]
    //     );
    //
    // Trimming whitespace alone leaves `//`, which matches no expectation form,
    // so the case was returned as `None` and dropped in silence.
    let expr = &expr[rustlit::skip_trivia(expr.as_bytes(), 0)..];
    if starts_string_literal(expr.as_bytes(), 0) {
        return Some(vec![crate::rustlit::scan(expr.as_bytes(), 0).ok()?.0]);
    }
    // `vec![…]`, `[…]`, and `&[…]` all appear; the reference form is what
    // `assert_eq!(run_js(…), &["1","2"])` uses.
    let body = expr.strip_prefix('&').unwrap_or(expr).trim();
    let body = body.strip_prefix("vec!").unwrap_or(body).trim();
    // `Vec::<String>::new()` — an expectation of NO output, which `vec![]`
    // spells as an empty list but the turbofish form does not. A test whose
    // whole point is that the program prints nothing (`exit(EXIT_FAILURE)`,
    // a watcher that never fires) is exactly the one this dropped.
    // A trailing comment is common on this form precisely because the empty
    // expectation needs explaining: `Vec::<String>::new() // fine as long as
    // nothing crashes`. No string literal can reach here, so splitting on `//`
    // is safe.
    if body.starts_with("Vec::")
        && body.split("//").next().unwrap_or(body).trim_end().ends_with("new()")
    {
        return Some(Vec::new());
    }
    if !body.starts_with('[') {
        return None;
    }
    let (lines, _) = scan_expected(body.as_bytes(), 0).ok()?;
    Some(lines)
}

fn close_paren(bytes: &[u8], from: usize) -> Option<usize> {
    let mut depth = 1usize;
    let mut i = from;
    while i < bytes.len() {
        match bytes[i] {
            b'"' => {
                i = crate::rustlit::scan(bytes, i).ok()?.1;
                continue;
            }
            b'r' if matches!(bytes.get(i + 1), Some(b'"') | Some(b'#')) => {
                i = crate::rustlit::scan(bytes, i).ok()?.1;
                continue;
            }
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

/// The comma separating `assert_eq!`'s two arguments.
fn top_level_comma(args: &str) -> Option<usize> {
    let bytes = args.as_bytes();
    let mut depth = 0i32;
    let mut i = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b'"' => {
                i = crate::rustlit::scan(bytes, i).ok()?.1;
                continue;
            }
            b'r' if matches!(bytes.get(i + 1), Some(b'"') | Some(b'#')) => {
                i = crate::rustlit::scan(bytes, i).ok()?.1;
                continue;
            }
            b'(' | b'[' | b'{' => depth += 1,
            b')' | b']' | b'}' => depth -= 1,
            b',' if depth == 0 => return Some(i),
            _ => {}
        }
        i += 1;
    }
    None
}

// ── paren-form macros ───────────────────────────────────────────────────────
//
// Python's corpus invokes one case per call rather than batching them:
//
//     crate::runtime_case!(for_range_basic, "for i in range(3):\n print(i)\n", "0");
//     crate::compile_case!(typing_list_int, "from typing import List\n");
//
// 3,469 of Python's 10,374 tests are written this way.

/// Every `name!(case_name, "src"[, expected])` invocation in a module.
pub fn paren_macros_in_file(text: &str) -> Vec<Case> {
    let bytes = text.as_bytes();
    let mut cases = Vec::new();
    let mut i = 0usize;

    while i < bytes.len() {
        if bytes[i] != b'!' || bytes.get(i + 1) != Some(&b'(') {
            i += 1;
            continue;
        }
        // Require a macro name before the `!`.
        let mut start = i;
        while start > 0 && (bytes[start - 1].is_ascii_alphanumeric() || bytes[start - 1] == b'_') {
            start -= 1;
        }
        if start == i {
            i += 1;
            continue;
        }

        // Std macros have the same SHAPE as a case macro —
        // `assert_eq!(out, "81")` is (identifier, string literal) — and matching
        // one manufactured a case named `out` whose source was `81`, giving 53
        // bogus parse failures.
        const NOT_CASES: [&str; 12] = [
            "assert_eq", "assert_ne", "assert", "debug_assert", "debug_assert_eq",
            "debug_assert_ne", "format", "panic", "print", "println", "write", "writeln",
        ];
        if NOT_CASES.contains(&&text[start..i]) {
            i += 1;
            continue;
        }

        let mut at = rustlit::skip_trivia(bytes, i + 2);
        // arg 1: the test name
        let name_start = at;
        while at < bytes.len() && (bytes[at].is_ascii_alphanumeric() || bytes[at] == b'_') {
            at += 1;
        }
        if at == name_start {
            i += 1;
            continue;
        }
        let name = text[name_start..at].to_string();

        at = rustlit::skip_trivia(bytes, at);
        if bytes.get(at) != Some(&b',') {
            i += 1;
            continue;
        }
        at = rustlit::skip_trivia(bytes, at + 1);
        if !starts_string_literal(bytes, at) {
            i += 1;
            continue;
        }
        let Ok((mut source, next)) = rustlit::scan(bytes, at) else {
            i += 1;
            continue;
        };
        // `vb_expr_spec!` passes a bare EXPRESSION, not a program: the macro
        // wraps it in a module before running. Rebuild that wrapper here or the
        // emitted file is not a valid program.
        if text[start..i].ends_with("vb_expr_spec") {
            source = format!(
                "Module M\n    Sub Main()\n        Console.WriteLine({source})\n    End Sub\nEnd Module\n"
            );
        }

        at = rustlit::skip_trivia(bytes, next);
        let case = match bytes.get(at) {
            // No third argument: compile-only.
            Some(b')') => Case {
                name,
                source,
                expected: None,
                expect_failure: false,
                single_line: false,
        run_only: false,
        prelude: None },
            Some(b',') => {
                let expr_at = rustlit::skip_trivia(bytes, at + 1);
                if bytes.get(expr_at) == Some(&b')') {
                    // Trailing comma, still compile-only.
                    Case { name, source, expected: None, expect_failure: false, single_line: false, run_only: false, prelude: None }
                } else {
                    let Some(close) = close_paren(bytes, i + 2) else {
                        i += 1;
                        continue;
                    };
                    let Some(expected) = parse_expected(&text[expr_at..close]) else {
                        i += 1;
                        continue;
                    };
                    // `runtime_case!` compares against `run_python_one`, which
                    // JOINS every line with "\n" — one value, many lines.
                    let single = expected.len() == 1;
                    Case { name, source, expected: Some(expected), expect_failure: false, single_line: single , run_only: false, prelude: None }
                }
            }
            _ => {
                i += 1;
                continue;
            }
        };
        cases.push(case);
        i = at;
    }
    cases
}
