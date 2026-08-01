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
}

#[derive(Clone, Copy)]
enum Shape {
    /// `name => ("src", vec!["line"]),`
    Run,
    /// `name => "src",`
    Compile,
    /// `name => "src",` but the frontend must reject it.
    CompileFail,
}

fn shape_of(macro_name: &str) -> Option<Shape> {
    Some(match macro_name {
        "go_run_cases" | "run_cases" => Shape::Run,
        "go_compile_cases" | "compile_cases" => Shape::Compile,
        "go_compile_fail_cases" => Shape::CompileFail,
        _ => return None,
    })
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
                Case { name, source, expected: Some(expected), expect_failure: false }
            }
            Shape::Compile | Shape::CompileFail => {
                let (source, next) = rustlit::scan(src, at)?;
                at = next;
                Case {
                    name,
                    source,
                    expected: None,
                    expect_failure: matches!(shape, Shape::CompileFail),
                }
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
    if src[at..].starts_with(b"vec!") {
        at = rustlit::skip_trivia(src, at + 4);
    }
    if src[at] != b'[' {
        anyhow::bail!("expected an expectation list at byte {at}");
    }
    at = rustlit::skip_trivia(src, at + 1);

    let mut lines = Vec::new();
    while src[at] != b']' {
        let (line, next) = rustlit::scan(src, at)?;
        lines.push(line);
        at = rustlit::skip_trivia(src, next);
        if src[at] == b',' {
            at = rustlit::skip_trivia(src, at + 1);
        }
    }
    Ok((lines, at + 1))
}
