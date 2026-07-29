use std::sync::{Arc, Mutex};
use vybe_runtime::{HostContext, VM, Value};

/// Render a value the way a WASM host logs it: an i64 is a plain integer, not a
/// JS BigInt, so drop the trailing `n` that `Value::BigInt`'s Display appends
/// (`9223372036854775807n` → `9223372036854775807`). The `n` is not part of the
/// WASM value — the `assert_return` path already strips it in `norm_str`; this
/// keeps the `call $log`/`wat_exec` path consistent. Only a digit-terminated
/// `n` is stripped, so `nan`/`inf` are untouched.
fn wasm_render(v: &Value) -> String {
    let s = format!("{v}");
    if let Some(head) = s.strip_suffix('n') {
        if head.chars().last().is_some_and(|c| c.is_ascii_digit()) {
            return head.to_string();
        }
    }
    s
}

/// Parse WAST/WAT source and return the Module (parse-only check).
pub fn parse_ok(src: &str) {
    vybe_language_wast::parse(src)
        .unwrap_or_else(|e| panic!("WAST parse failed:\n{}\nSource:\n{}", e, src));
}

/// Parse and expect a parse error (negative test).
#[allow(dead_code)]
pub fn parse_err(src: &str) {
    assert!(
        vybe_language_wast::parse(src).is_err(),
        "Expected parse error but succeeded for:\n{}",
        src
    );
}

/// Parse + compile through the full pipeline (no VM execution).
pub fn compile_ok(src: &str) {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_wast::register);
    }
    let module =
        vybe_language_wast::parse(src).unwrap_or_else(|e| panic!("WAST parse failed:\n{}", e));
    let profile = vybe_compiler::profile::parse_profile(vybe_language_wast::profile_source())
        .expect("Failed to parse WAST profile");
    vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap_or_else(|e| panic!("WAST compile failed:\n{}", e));
}

/// Run WAST source through the full pipeline and return console output lines.
pub fn run_wast(src: &str) -> Vec<String> {
    run_wast_result(src).unwrap_or_else(|e| panic!("WAST run failed:\n{}", e))
}

/// Like [`run_wast`] but surfaces a VM runtime error (a WASM trap) as `Err`
/// instead of panicking, so trap-expecting spec tests can assert on it. Parse
/// and compile failures still panic — those are test-setup errors, not traps.
pub fn run_wast_result(src: &str) -> Result<Vec<String>, String> {
    let module =
        vybe_language_wast::parse(src).unwrap_or_else(|e| panic!("WAST parse failed:\n{}", e));
    let profile = vybe_compiler::profile::parse_profile(vybe_language_wast::profile_source())
        .expect("Failed to parse WAST profile");
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap_or_else(|e| panic!("WAST compile failed:\n{}", e));

    for (i, c) in chunks.iter().enumerate() {
        println!("Chunk {i}:\n{}", vybe_runtime::debug::disassemble(c));
    }

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    let out_cloned = out.clone();
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(wasm_render).collect();
            out_cloned.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    let out_cloned2 = out.clone();
    vm.register_host_fn(
        "wasi:cli",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(wasm_render).collect();
            out_cloned2.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    // Real WAST assertion validation: `(assert_return (invoke …) (expected))`
    // lowers to `__wast_assert_return(actual, expected…)`. Compare and record a
    // failure so a wrong result (e.g. an integer that didn't wrap) is caught,
    // instead of the old compile-only checks that validated nothing.
    let asserts = out.clone();
    vm.register_host_fn(
        "vybe:wast",
        "assert_return",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            // Compare values, not display quirks: an i64 carried as a BigInt
            // prints with a trailing `n`, which is not part of the WASM value.
            fn norm_str(s: String) -> String {
                let low = s.to_ascii_lowercase();
                // Every NaN form compares equal — the WASM `assert_return` NaN
                // patterns (`nan:canonical`/`nan:arithmetic`) don't pin a payload,
                // and F32 prints `nan` while F64 prints `NaN`.
                if low == "nan" || low == "-nan" || low.starts_with("nan:") {
                    return "nan".to_string();
                }
                // An i64 carried as a BigInt prints a trailing `n` (`42n`) that is
                // not part of the WASM value — strip it, but only when it follows
                // a digit (so it never eats the `n` of `nan`/`inf`).
                if let Some(head) = s.strip_suffix('n') {
                    if head.chars().last().is_some_and(|c| c.is_ascii_digit()) {
                        return head.to_string();
                    }
                }
                // A whole-number float prints as `2.0` when carried as an F32 but
                // `2` when the expected `f32.const 2.0` is folded to an integral
                // number — the same WASM value. Drop a trailing `.0` so the two
                // spellings compare equal (only the exact `.0` fraction; `2.5`
                // and `-0.0` keep their form).
                if let Some(head) = s.strip_suffix(".0") {
                    if head.chars().all(|c| c.is_ascii_digit() || c == '-') && !head.is_empty() {
                        return head.to_string();
                    }
                }
                s
            }
            fn norm(v: &Value) -> String {
                norm_str(format!("{v}"))
            }
            if args.len() >= 2 {
                let expected: Vec<String> = args[1..].iter().map(norm).collect();
                // A multi-value `invoke` result arrives packed as a single array
                // that Displays comma-joined (`1,2,3`); split it so each result is
                // compared position-wise to its expected value. A single result
                // keeps the original any-match behavior.
                let raw = format!("{}", args[0]);
                let actuals: Vec<String> = if expected.len() > 1 {
                    raw.split(',')
                        .map(|s| norm_str(s.trim().to_string()))
                        .collect()
                } else {
                    vec![norm(&args[0])]
                };
                // The spec's abstract `(ref.func)` pattern matches ANY non-null
                // funcref (Displayed as `[function …]`); `__wast_any_funcref` is
                // the walker's sentinel for it.
                fn matches(actual: &str, expected: &str) -> bool {
                    if expected == "__wast_any_funcref" {
                        return actual.starts_with("[function") || actual.starts_with("function ");
                    }
                    actual == expected
                }
                let ok = if actuals.len() == expected.len() {
                    actuals.iter().zip(&expected).all(|(a, e)| matches(a, e))
                } else {
                    // Shape mismatch: fall back to the single-value any-match.
                    let actual = norm(&args[0]);
                    expected.iter().any(|e| matches(&actual, e))
                };
                if !ok {
                    asserts.lock().unwrap().push(format!(
                        "ASSERT_FAIL: got {} expected {}",
                        actuals.join(","),
                        expected.join(" ")
                    ));
                }
            }
            Value::Null
        }),
    );
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).map_err(|e| e.to_string())?;
    let result = output.lock().unwrap().clone();
    // Surface assertion failures as an error so result-checking tests fail loudly.
    if let Some(fail) = result.iter().find(|l| l.starts_with("ASSERT_FAIL:")) {
        return Err(fail.clone());
    }
    Ok(result)
}

/// Run a WAST script that contains `assert_return` commands; returns Ok(()) only
/// if every assertion's actual result matched its expected value.
pub fn run_wast_asserts(src: &str) -> Result<(), String> {
    run_wast_result(src).map(|_| ())
}

pub fn run_wast_one(src: &str) -> String {
    run_wast(src).into_iter().next().unwrap_or_default()
}

#[macro_export]
macro_rules! wat_exec {
    ($($name:ident => { $src:expr, $expect:expr }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                let wrapped_src = format!(
                    r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  {}
)
"#,
                    $src
                );

                // If it already is a module, don't wrap it. Just assume if it starts with (module it's complete.
                let final_src = if $src.trim().starts_with("(module") {
                    $src.to_string()
                } else {
                    wrapped_src
                };

                // A `"trap"` expectation means the program must raise a WASM
                // trap (VM runtime error); anything else asserts on stdout.
                if $expect == "trap" {
                    assert!(
                        $crate::helpers::run_wast_result(&final_src).is_err(),
                        "expected a trap, but the program completed normally"
                    );
                } else {
                    let out = $crate::helpers::run_wast_one(&final_src);
                    assert_eq!(out, $expect);
                }
            }
        )*
    };
}
