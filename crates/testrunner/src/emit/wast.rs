//! WAST emitter: one extracted case → a standalone `.wat` file.
//!
//! WAST is the one corpus that is mostly NOT about output: 428 `parse_ok` and
//! 90 `compile_ok` cases assert only that the front-end accepts the module.
//! Those need no harness and no rewriting — they become compile-mode files.
//!
//! A source carrying `(assert_return …)` is a `.wast` SCRIPT and is emitted
//! verbatim with a `.wast` extension: that is the portable form every spec
//! interpreter, wabt and `wasm-tools` already run. Do NOT rewrite it into WAT
//! with `unreachable` — that would make the file less portable, not more.
//! `vybex` cannot run these today because it never registers
//! `vybe:wast.assert_return`; that host lives only in
//! `languages/wast/tests/wast/helpers.rs`. It is a vybex gap, not a test defect.
//!
//! The `wat_exec!` cases DO assert on output, but WAT has no string handling,
//! so a `__check` cannot be written the way Go's or JS's is. Their values are
//! numeric — the wrapper's imports are `log(i32)`, `log_i64`, `log_f32`,
//! `log_f64` — so the eventual form is a check function comparing numerically
//! and `unreachable`-ing on mismatch: a trap is a non-zero exit, the same
//! verdict mechanism every other language uses. Until that exists they are
//! emitted without assertions and reported as unpairable.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing,
    /// `.wast` for a script carrying assertions, `.wat` for a bare module —
    /// the extension is what tells a spec interpreter to run the directives.
    pub extension: &'static str,
}

/// A source with script directives is a `.wast`, not a plain module.
fn extension_for(src: &str) -> &'static str {
    if src.contains("(assert_") || src.contains("(invoke") || src.contains("(register") {
        "wast"
    } else {
        "wat"
    }
}

/// `wat_exec!` wraps a bare function body in a module carrying the four logging
/// imports — unless the source already is a module.
const MODULE_WRAPPER: &str = r#"(module
  (import "web:console" "log" (func $log (param i32)))
  (import "web:console" "log" (func $log_i64 (param i64)))
  (import "web:console" "log" (func $log_f32 (param f32)))
  (import "web:console" "log" (func $log_f64 (param f64)))
"#;

pub fn emit(case: &Case, origin: &str, slug: &str, _harness: &str) -> Emitted {
    let header = format!(";; vybe-test: {slug}\n;; origin: {origin}\n");
    let body = wrap_module(case.source.trim());

    let Some(expected) = case.expected.as_ref() else {
        let mode = if case.run_only {
            // A `.wast` script asserts through its own directives, so running it
            // IS the check — `must_fail` inverts the verdict.
            if case.expect_failure { "run-fail" } else { "run" }
        } else if case.expect_failure {
            "compile-fail"
        } else {
            "compile"
        };
        return Emitted {
            text: format!("{header};; vybe-test-mode: {mode}\n\n{body}\n"),
            pairing: Pairing::Direct,
            extension: extension_for(&case.source),
        };
    };

    // `wat_exec!` treats the literal expectation "trap" as "the program must
    // fail at run time" — distinct from compile-fail, which asserts the
    // FRONT-END rejects it. The runner has no run-fail mode yet.
    if expected.first().map(String::as_str) == Some("trap") && expected.len() == 1 {
        return Emitted {
            text: format!("{header};; vybe-test-mode: run-fail\n\n{body}\n"),
            pairing: Pairing::Direct,
            extension: extension_for(&case.source),
        };
    }
    match numeric_checked(&body, expected) {
        Ok(out) => Emitted {
            text: format!("{header}\n{out}\n"),
            pairing: Pairing::Direct,
            extension: extension_for(&case.source),
        },
        Err(reason) => Emitted {
            text: format!("{header}\n{body}\n"),
            pairing: Pairing::Unpairable(reason),
            extension: extension_for(&case.source),
        } }
}

/// The four logging imports the wrapper declares, and the wasm type each takes.
const LOG_CALLS: [(&str, &str); 4] = [
    ("call $log_i64", "i64"),
    ("call $log_f32", "f32"),
    ("call $log_f64", "f64"),
    ("call $log", "i32"),
];

/// Replace each `call $log*` with a comparison against the expected value.
///
/// WAT has no string handling, so there is no `__check(got, want)` of the shape
/// every other language uses — but it does not need one. The logged value is
/// ALREADY on the stack at the call, so pushing the expected constant after it
/// gives a two-parameter check function its arguments in order. A mismatch
/// executes `unreachable`, and a trap is a non-zero exit: the same verdict
/// mechanism as everywhere else.
fn numeric_checked(src: &str, expected: &[String]) -> Result<String, String> {
    let calls = find_log_calls(src);
    if calls.is_empty() {
        return Err("no `call $log` to check".into());
    }
    // A logged value inside a loop runs an unknown number of times, so the i-th
    // call is not the i-th line.
    if src.contains("loop ") || src.contains("(loop") {
        return Err("loop — logged value count is not static".into());
    }
    if calls.len() != expected.len() {
        return Err(format!(
            "{} logging call(s) but {} expected line(s)",
            calls.len(),
            expected.len()
        ));
    }

    let mut used: Vec<&str> = Vec::new();
    let mut nan_used: Vec<&str> = Vec::new();
    let mut out = src.to_string();
    // Back-to-front so the earlier spans keep their offsets.
    for (i, (start, end, ty)) in calls.iter().enumerate().rev() {
        let want = expected[i].trim();
        // NaN is not equal to ANYTHING, itself included, so `f64.ne` against a
        // NaN constant traps on a correct value. The assertion for NaN is
        // therefore "is not a number" — `got == got` is false exactly when it
        // holds — and it needs its own one-parameter check function.
        if is_nan_text(want) {
            if !matches!(*ty, "f32" | "f64") {
                return Err(format!("expected `{want}` is not a {ty} value"));
            }
            let name = format!("vybe_check_nan_{ty}");
            if !nan_used.contains(ty) {
                nan_used.push(ty);
            }
            out.replace_range(*start..*end, &format!("call ${name}"));
            continue;
        }
        let Some(literal) = wasm_const(ty, want) else {
            return Err(format!("expected `{want}` is not a {ty} value"));
        };
        if !used.contains(ty) {
            used.push(ty);
        }
        out.replace_range(*start..*end, &format!("{ty}.const {literal} call $vybe_check_{ty}"));
    }

    // The check functions go after the LAST import: an import may not follow a
    // function definition.
    let at = after_last_import(&out);
    let mut defs = String::new();
    for ty in &used {
        // STACK form, NOT the folded `(if (cond) (then …))`. Vybe's WAT
        // frontend does not execute a folded if at all — measured: a module
        // whose `_start` is `(if (i32.ne (i32.const 5) (i32.const 9)) (then
        // unreachable))` exits 0, while `i32.const 1 if unreachable end` exits
        // 1. Written folded, this check was VACUOUS: a deliberately corrupted
        // expected value still passed.
        defs.push_str(&format!(
            "\n  (func $vybe_check_{ty} (param {ty}) (param {ty})\
             \n    local.get 0\n    local.get 1\n    {ty}.ne\n    if\n      unreachable\n    end)"
        ));
    }
    for ty in &nan_used {
        // `got == got` is TRUE only for a non-NaN, so that is the trap
        // condition. Stack form, for the same reason as above.
        defs.push_str(&format!(
            "\n  (func $vybe_check_nan_{ty} (param {ty})\
             \n    local.get 0\n    local.get 0\n    {ty}.eq\n    if\n      unreachable\n    end)"
        ));
    }
    out.insert_str(at, &defs);
    Ok(out)
}

/// Every `call $log*`, in source order, as (start, end, wasm type).
fn find_log_calls(src: &str) -> Vec<(usize, usize, &'static str)> {
    let mut out: Vec<(usize, usize, &'static str)> = Vec::new();
    let bytes = src.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        // Longest first: `call $log` is a prefix of `call $log_i64`.
        let hit = LOG_CALLS
            .iter()
            .find(|(needle, _)| src.is_char_boundary(i) && src[i..].starts_with(needle));
        if let Some((needle, ty)) = hit {
            let end = i + needle.len();
            // `call $log` must not be the prefix of a longer name.
            let next = bytes.get(end).copied().unwrap_or(b' ');
            if !next.is_ascii_alphanumeric() && next != b'_' {
                out.push((i, end, ty));
                i = end;
                continue;
            }
        }
        i += 1;
    }
    out
}

/// Is this expectation a NaN in any of the spellings the corpus uses?
fn is_nan_text(want: &str) -> bool {
    let t = want.trim().trim_start_matches(['-', '+']).to_ascii_lowercase();
    t == "nan" || t.starts_with("nan:")
}

/// The expected text as a constant of that wasm type, or None if it is not one.
fn wasm_const(ty: &str, want: &str) -> Option<String> {
    match ty {
        "i32" | "i64" => {
            let t = want.strip_prefix('-').unwrap_or(want);
            (!t.is_empty() && t.chars().all(|c| c.is_ascii_digit())).then(|| want.to_string())
        }
        _ => {
            want.parse::<f64>().ok()?;
            Some(want.to_string())
        } }
}

fn after_last_import(src: &str) -> usize {
    let mut at = 0usize;
    let mut best: Option<usize> = None;
    for line in src.split_inclusive('\n') {
        at += line.len();
        if line.contains("(import ") {
            best = Some(at - 1);
        }
    }
    // No imports at all — straight after `(module`.
    best.unwrap_or_else(|| src.find("(module").map(|o| o + "(module".len()).unwrap_or(0))
}

fn wrap_module(src: &str) -> String {
    // Anywhere, not just at the start: `(; block comment ;) (module)` is a
    // complete module behind a preamble, and wrapping it produced 54 parse
    // errors.
    if src.contains("(module") {
        return src.to_string();
    }
    // A script directive is TOP-LEVEL, never a module field. Wrapping one put
    // `(assert_malformed …)` inside the import module and wasmtime rejected it.
    if extension_for(src) == "wast" {
        return src.to_string();
    }
    format!("{MODULE_WRAPPER}  {src}\n)")
}
