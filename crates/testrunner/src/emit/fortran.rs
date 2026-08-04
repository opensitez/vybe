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
    // STATIC pairing first, unchanged: where the i-th print really is the i-th
    // line, the emitted file stays byte-identical to what was verified before.
    if unpairable(body, &prints, expected.len()).is_none() {
        let mut out = body.to_string();
        for (i, (start, end, expr)) in prints.iter().enumerate().rev() {
            out.replace_range(*start..*end, &check_for(expr, &expected[i]));
        }
        return Emitted {
            text: format!("{header}{out}\n"),
            pairing: Pairing::Direct };
    }

    // RUNTIME pairing: an expected table plus a counter, checked where each
    // print stands. The counter advances as the program runs, so a print in a
    // LOOP pairs correctly and an `if`/`else` pair contributes exactly the one
    // line that actually executed — 503 and 365 cases respectively, which
    // static pairing can only refuse.
    //
    // Still VALUES, not text. Fortran's list-directed field widths are
    // processor-dependent (`8` here, `           8` under gfortran), so text
    // comparison is not portable at any pairing strategy.
    match runtime_paired(body, &prints, expected) {
        Ok(out) => Emitted {
            text: format!("{header}{out}\n"),
            pairing: Pairing::Direct },
        Err(reason) => Emitted {
            text: format!("{header}{body}\n"),
            pairing: Pairing::Unpairable(reason) } }
}

// NOTE: the names must begin with a LETTER. `__vybe_i` is not a Fortran
// identifier at all — gfortran rejects it outright ("Invalid character in name").
//
/// The Fortran type one table can hold. Arrays are homogeneous, so a case whose
/// expectations mix kinds cannot be tabulated at all.
#[derive(PartialEq, Clone, Copy)]
enum TableKind {
    Integer,
    Real,
    Logical,
    Character }

fn table_kind(expected: &[String]) -> Option<TableKind> {
    let mut kind: Option<TableKind> = None;
    for w in expected {
        let w = w.trim();
        let this = if is_integer(w) {
            TableKind::Integer
        } else if is_real(w) {
            TableKind::Real
        } else if as_logical(w).is_some() {
            TableKind::Logical
        } else {
            TableKind::Character
        };
        kind = Some(match kind {
            None => this,
            // An integer among reals is still a real table; anything else must
            // match exactly.
            Some(prev) if prev == this => prev,
            Some(TableKind::Real) if this == TableKind::Integer => TableKind::Real,
            Some(TableKind::Integer) if this == TableKind::Real => TableKind::Real,
            Some(_) => return None });
    }
    kind
}

fn runtime_paired(src: &str, prints: &[Print], expected: &[String]) -> Result<String, String> {
    if prints.is_empty() {
        return Err("no `print *,` to pair".into());
    }
    if src.contains("write(") || src.contains("write (") {
        return Err("uses `write` — output does not pass through `print`".into());
    }
    if let Some((_, _, expr)) = prints.iter().find(|(_, _, e)| has_top_level_comma(e)) {
        return Err(format!("multi-value print (`{expr}`) — one line, several values"));
    }
    let Some(kind) = table_kind(expected) else {
        return Err("expectations mix types — a Fortran array is homogeneous".into());
    };
    if expected.is_empty() {
        return Err("no expected lines to tabulate".into());
    }

    let n = expected.len();
    let decls = table_declaration(kind, expected);
    let at = declarations_insert_at(src);

    let mut out = src.to_string();
    // Back-to-front so the earlier spans keep their offsets, and the
    // declaration insert (which is before all of them) goes last.
    for (start, end, expr) in prints.iter().rev() {
        let indent = " ".repeat(*start - line_start(src, *start));
        out.replace_range(*start..*end, &runtime_check(kind, expr, n, &indent));
    }
    out = format!("{}{decls}{}", &out[..at], &out[at..]);

    // Too FEW lines is a failure too, and only the end can see it.
    let epilogue = format!(
        "if (vybe_check_i /= {n}) then\n    print *, \"FAIL: \", vybe_check_i, \" line(s), wanted {n}\"\n    stop 1\nend if\n"
    );
    match find_end_program(&out) {
        Some(end_at) => out.insert_str(end_at, &epilogue),
        None => {
            // Without the separator this runs straight into the last statement
            // — `end ifif (vybe_check_i /= 1) …`, which gfortran reports as
            // "Expected terminating name".
            if !out.ends_with('\n') {
                out.push('\n');
            }
            out.push_str(&epilogue);
        } }
    Ok(out)
}

fn table_declaration(kind: TableKind, expected: &[String]) -> String {
    let n = expected.len();
    let (ty, items) = match kind {
        TableKind::Integer => (
            "integer".to_string(),
            expected.iter().map(|w| w.trim().to_string()).collect::<Vec<_>>(),
        ),
        TableKind::Real => (
            "real".to_string(),
            expected.iter().map(|w| w.trim().to_string()).collect(),
        ),
        TableKind::Logical => (
            "logical".to_string(),
            expected
                .iter()
                .map(|w| format!(".{}.", as_logical(w).unwrap_or("false")))
                .collect(),
        ),
        TableKind::Character => {
            // Every element of a character array has the SAME length, so the
            // table is declared at the longest and the comparison trims.
            let width = expected.iter().map(|w| w.trim().len()).max().unwrap_or(1).max(1);
            (
                format!("character(len={width})"),
                expected
                    .iter()
                    .map(|w| format!("\"{}\"", w.trim().replace('"', "'")))
                    .collect(),
            )
        }
    };
    format!(
        "integer :: vybe_check_i = 0\n{ty} :: vybe_check_w({n}) = [ {} ]\n",
        items.join(", ")
    )
}

fn runtime_check(kind: TableKind, expr: &str, n: usize, indent: &str) -> String {
    // `.or.` does NOT short-circuit in Fortran, so the bounds test is its own
    // `if` — folding it into the comparison would index `vybe_check_w` past its end.
    let differs = match kind {
        TableKind::Integer => format!("({expr}) /= vybe_check_w(vybe_check_i)"),
        // Never `/=` on a real: the same value can print identically and differ
        // in the last bit.
        TableKind::Real => format!("abs(({expr}) - vybe_check_w(vybe_check_i)) > 1.0e-6"),
        TableKind::Logical => format!("({expr}) .neqv. vybe_check_w(vybe_check_i)"),
        TableKind::Character => format!("trim({expr}) /= trim(vybe_check_w(vybe_check_i))") };
    format!(
        "{indent}vybe_check_i = vybe_check_i + 1\n\
         {indent}if (vybe_check_i > {n}) then\n\
         {indent}    print *, \"FAIL: more than {n} line(s)\"\n\
         {indent}    stop 1\n\
         {indent}end if\n\
         {indent}if ({differs}) then\n\
         {indent}    print *, \"FAIL at \", vybe_check_i, \" got [\", {expr}, \"]\"\n\
         {indent}    stop 1\n\
         {indent}end if"
    )
}

/// Where a declaration may be inserted: after `implicit none` if present,
/// otherwise straight after the `program` statement. A declaration cannot
/// precede `implicit none`, and none may follow an executable statement.
fn declarations_insert_at(src: &str) -> usize {
    let mut at = 0usize;
    let mut best = 0usize;
    for line in src.split_inclusive('\n') {
        let t = line.trim().to_ascii_lowercase();
        at += line.len();
        if t.starts_with("program ") || t == "program" || t.starts_with("implicit ") {
            best = at;
        } else if !t.is_empty() && !t.starts_with('!') {
            break;
        }
    }
    best
}

/// Where the program's terminator begins. Both spellings occur in the corpus:
/// `end program t` and a bare `end`.
fn find_end_program(src: &str) -> Option<usize> {
    let mut at = 0usize;
    for line in src.split_inclusive('\n') {
        let t = line.trim().to_ascii_lowercase();
        if t.starts_with("end program") || t == "end" {
            return Some(at);
        }
        at += line.len();
    }
    None
}

fn line_start(src: &str, at: usize) -> usize {
    src[..at].rfind('\n').map(|o| o + 1).unwrap_or(0)
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
    } else if let Some(b) = as_logical(w) {
        format!("({expr}) .neqv. .{b}.")
    } else {
        format!("trim({expr}) /= \"{}\"", w.replace('"', "'"))
    };
    format!(
        "if ({condition}) then\n    print *, \"FAIL: want [{}] got [\", {expr}, \"]\"\n    stop 1\nend if",
        w.replace('"', "'")
    )
}

/// A logical expectation, in either spelling the corpus carries.
///
/// `T`/`F` is what the standard specifies and what gfortran writes; `true`/
/// `false` is what Vybe used to write, so older recorded expectations use it.
/// Both map to the same `.true.`/`.false.` comparison — the check is on the
/// VALUE, so the rendering the text came from does not matter.
fn as_logical(text: &str) -> Option<&'static str> {
    match text.trim() {
        "T" | "t" => Some("true"),
        "F" | "f" => Some("false"),
        w if w.eq_ignore_ascii_case("true") => Some("true"),
        w if w.eq_ignore_ascii_case("false") => Some("false"),
        _ => None }
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
