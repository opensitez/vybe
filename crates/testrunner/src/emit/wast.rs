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
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
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
    let reason = {
        "WAT has no string compare — needs the numeric __check harness".to_string()
    };
    Emitted {
        text: format!("{header}\n{body}\n"),
        pairing: Pairing::Unpairable(reason),
        extension: extension_for(&case.source),
    }
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
