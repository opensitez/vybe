//! Fortran emitter: one extracted case → a standalone `.f90` test.
//!
//! Sources are already complete programs (`program t … end program t`), so
//! nothing is wrapped — each `print *, <expr>` is replaced by a comparison
//! against the expected line.
//!
//! **Values, not printed text.** The corpus records what Vybe's logging host
//! produced — bare `["8", "9", "false", "true"]` — while gfortran's
//! list-directed `print *` pads (`"           8"`) and writes logicals as
//! `T`/`F`. Comparing text would fail under gfortran on formatting alone and
//! the differential would tell us nothing; comparing the value is semantics,
//! which both runtimes agree on. See `harness/fortran/check.f90`.
//!
//! Failure uses `stop <code>`, which carries a real status only since the
//! walker stopped lowering `stop` to a bare `Return` (it mapped to `noop` in
//! the profile, so every `stop 1` exited 0 where gfortran gave 1).

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing }

pub fn emit(case: &Case, origin: &str, slug: &str, _harness: &str) -> Emitted {
    let header = format!("! vybe-test: {slug}\n! origin: {origin}\n");
    let body = case.source.trim_end();

    let Some(expected) = case.expected.as_ref() else {
        return Emitted {
            text: format!("{header}{body}\n"),
            pairing: Pairing::Direct };
    };

    let prints = find_prints(body);
    if let Some(reason) = unpairable(body, &prints, expected.len()) {
        return Emitted {
            text: format!("{header}{body}\n"),
            pairing: Pairing::Unpairable(reason) };
    }

    let mut out = body.to_string();
    for (i, (start, end, expr)) in prints.iter().enumerate().rev() {
        out.replace_range(*start..*end, &check_for(expr, &expected[i]));
    }

    Emitted {
        text: format!("{header}{out}\n"),
        pairing: Pairing::Direct }
}

/// The comparison implied by the shape of the expected text.
fn check_for(expr: &str, want: &str) -> String {
    let w = want.trim();
    let condition = if is_integer(w) {
        format!("({expr}) /= {w}")
    } else if is_real(w) {
        // Never `/=` on a real: the same value can print identically and
        // differ in the last bit.
        format!("abs(({expr}) - {w}) > 1.0e-6")
    } else if w.eq_ignore_ascii_case("true") || w.eq_ignore_ascii_case("false") {
        format!("({expr}) .neqv. .{}.", w.to_ascii_lowercase())
    } else {
        format!("trim({expr}) /= \"{}\"", w.replace('"', "'"))
    };
    format!(
        "if ({condition}) then\n    print *, \"FAIL: want [{}] got [\", {expr}, \"]\"\n    stop 1\nend if",
        w.replace('"', "'")
    )
}

fn is_integer(text: &str) -> bool {
    let t = text.strip_prefix('-').unwrap_or(text);
    !t.is_empty() && t.chars().all(|c| c.is_ascii_digit())
}

fn is_real(text: &str) -> bool {
    let t = text.strip_prefix('-').unwrap_or(text);
    let mut parts = t.splitn(2, '.');
    let (Some(a), Some(b)) = (parts.next(), parts.next()) else {
        return false;
    };
    !a.is_empty()
        && !b.is_empty()
        && a.chars().all(|c| c.is_ascii_digit())
        && b.chars().all(|c| c.is_ascii_digit())
}

/// (start, end, printed expression) for each `print *, <expr>` on its own line.
type Print = (usize, usize, String);

fn find_prints(src: &str) -> Vec<Print> {
    let mut out = Vec::new();
    let mut offset = 0usize;
    for line in src.split_inclusive('\n') {
        let trimmed = line.trim_end_matches(['\n', '\r']);
        let lead = trimmed.len() - trimmed.trim_start().len();
        let body = trimmed.trim_start();
        if let Some(rest) = body.strip_prefix("print *,") {
            let expr = rest.trim();
            if !expr.is_empty() {
                out.push((offset + lead, offset + trimmed.len(), expr.to_string()));
            }
        }
        offset += line.len();
    }
    out
}

fn unpairable(src: &str, prints: &[Print], expected: usize) -> Option<String> {
    if prints.is_empty() {
        return Some("no `print *,` to pair".into());
    }
    // A print inside a loop runs an unknown number of times, so the i-th print
    // is not the i-th line.
    if src.contains("do ") || src.contains("do concurrent") {
        return Some("loop — print count is not static".into());
    }
    // `write` writes on its own terms; its output is not one line per `print`.
    if src.contains("write(") || src.contains("write (") {
        return Some("uses `write` — output is not one line per print".into());
    }
    if prints.len() != expected {
        return Some(format!(
            "{} print(s) but {expected} expected line(s)",
            prints.len()
        ));
    }
    // A print of several comma-separated values produces ONE line holding all
    // of them, so there is no single value to compare.
    if let Some((_, _, expr)) = prints.iter().find(|(_, _, e)| has_top_level_comma(e)) {
        return Some(format!("multi-value print (`{expr}`) — one line, several values"));
    }
    None
}

fn has_top_level_comma(expr: &str) -> bool {
    let mut depth = 0i32;
    let mut quote: Option<char> = None;
    for ch in expr.chars() {
        match quote {
            Some(q) if ch == q => quote = None,
            Some(_) => {}
            None => match ch {
                '"' | '\'' => quote = Some(ch),
                '(' | '[' => depth += 1,
                ')' | ']' => depth -= 1,
                ',' if depth == 0 => return true,
                _ => {}
            } }
    }
    false
}
