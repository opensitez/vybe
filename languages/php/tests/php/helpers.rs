//! PHP test helpers — compile through the vybex pipeline.
//!
//! The original `vybe_compiler_php` test files use four entry points:
//! `parse`, `compile`, `compile_ok`, `run`, `run_prints`. Every test
//! file currently in this suite uses only `compile_ok` (parse + compile
//! must succeed; runtime is not asserted), so we provide that shim
//! plus the others as drop-in replacements.
//!
//! `compile_ok(src)` runs `vybe_language_php::parse` → walker →
//! `vybe_compiler::primitives::Compiler::with_profile` and asserts the result is
//! Ok. Any parse error, walker error, or compile error fails the test
//! with the underlying message.

#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_runtime::{HostContext, VM, Value};

#[macro_export]
macro_rules! php_cases {
    ($($name:ident => { $src:expr, [$($expected:expr),* $(,)?] };)+) => {
        $(
            #[test]
            fn $name() {
                $crate::helpers::assert_php_prints($src, &[$($expected),*]);
            }
        )+
    };
}

// Output capture mirrors the two real output surfaces:
//
// - `wasi:cli/stdout.write-via-stream(data: stream<u8>)` (echo/printf):
//   the WASI 0.3 stdout surface — raw PHP-stdout bytes, NO implicit
//   newline, so consecutive writes concatenate onto the current line
//   (`echo 'yes'; echo 'no';` → "yesno"; `printf('hi'); echo ' 2';`
//   → "hi 2").
// - `wasi:logging/logging.log` (console-style logging): one line-oriented
//   record per call — the real host fn println!s the message, so the
//   capture appends the same newline.
//
// Fragments are buffered and split into lines in `finish_output`.
fn capture_log_lines(output: &Arc<Mutex<Vec<String>>>, args: &[Value]) {
    let mut joined = args
        .iter()
        .map(|arg| format!("{}", arg))
        .collect::<Vec<_>>()
        .join(" ");
    joined.push('\n');
    output.lock().unwrap().push(joined);
}

fn register_output_capture(vm: &mut VM, output: &Arc<Mutex<Vec<String>>>) {
    let out = output.clone();
    vm.register_host_fn(
        "web:console",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            capture_log_lines(&out, args);
            Value::Null
        }),
    );
    let out = output.clone();
    vm.register_host_fn(
        "wasi:cli/stdout",
        "write-via-stream",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let stream_val = args.first().cloned().unwrap_or(Value::Null);
            let bytes = ctx.stream_drain(&stream_val);
            if !bytes.is_empty() {
                let text = String::from_utf8_lossy(&bytes).into_owned();
                let mut vec = out.lock().unwrap();
                // Concatenate with last stream fragment (PHP echo is
                // unbuffered — no newline between calls).
                let append = vec.last().map_or(false, |l| !l.ends_with('\n'));
                if append {
                    vec.last_mut().unwrap().push_str(&text);
                } else {
                    vec.push(text);
                }
            }
            Value::Null
        }),
    );
}

fn finish_output(output: &Arc<Mutex<Vec<String>>>) -> Vec<String> {
    let fragments = output.lock().unwrap().clone();
    let mut result: Vec<String> = Vec::new();
    for fragment in fragments {
        let mut parts = fragment.split('\n').peekable();
        while let Some(part) = parts.next() {
            let part = part.trim_end_matches('\r');
            let has_more = parts.peek().is_some();
            if part.is_empty() {
                continue;
            }
            if part.starts_with(char::is_whitespace) {
                if let Some(prev) = result.last_mut() {
                    prev.push_str(part);
                } else {
                    result.push(part.to_string());
                }
            } else if has_more || !part.is_empty() {
                result.push(part.to_string());
            }
        }
    }
    while result
        .last()
        .map(|s: &String| s.is_empty())
        .unwrap_or(false)
    {
        result.pop();
    }
    result
}

fn compile_chunks(src: &str) -> Result<Vec<vybe_runtime::Chunk>, String> {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_php::register);
    }
    let module = vybe_language_php::parse(src)?;
    let profile = php_profile();
    vybe_compiler::primitives::Compiler::with_profile(profile).compile(&module)
}

fn php_profile() -> vybe_compiler::profile::LanguageProfile {
    static PROFILE: std::sync::OnceLock<vybe_compiler::profile::LanguageProfile> =
        std::sync::OnceLock::new();
    PROFILE
        .get_or_init(|| {
            vybe_compiler::profile::parse_profile(vybe_language_php::profile_source())
                .expect("php profile parse failed")
        })
        .clone()
}

/// Asserts that `src` parses + compiles cleanly. Used by ~all PHP test
/// files; the original tests don't run the bytecode, just verify the
/// pipeline accepts the source.
pub fn compile_ok(src: &str) {
    match compile_chunks(src) {
        Ok(chunks) => {
            assert!(!chunks.is_empty(), "compile produced no chunks");
        }
        Err(e) => panic!("compile failed: {}", e),
    }
}

/// Returns the compiled chunks, or panics with the error.
pub fn compile(src: &str) -> Vec<vybe_runtime::Chunk> {
    match compile_chunks(src) {
        Ok(c) => c,
        Err(e) => panic!("compile failed: {}", e),
    }
}

/// Compile + run, return the final value popped from the stack.
pub fn run(src: &str) -> Value {
    let chunks = compile(src);
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    vm.run(chunks).expect("run failed")
}

/// Compile + run, capture stdout (anything routed through `wasi:cli::log`),
/// return the captured lines.
pub fn run_prints(src: &str) -> Vec<String> {
    // Run through the RuntimeCompilerService like the real exe (cli.rs) and
    // the JS harness (`run_with_js_dynamic_runtime`) do — a bare `vm.run`
    // never activates the compiler-as-a-service, so `eval`/`include` and
    // other dynamic-compile paths silently break. See `run_prints_dynamic`.
    run_prints_dynamic(src, "test.php")
}

pub fn assert_php_prints(src: &str, expected: &[&str]) {
    let actual = run_prints(src);
    let expected_vec: Vec<String> = expected.iter().map(|line| (*line).to_string()).collect();
    assert_eq!(actual, expected_vec);
}

pub fn run_prints_dynamic(src: &str, virtual_path: &str) -> Vec<String> {
    // Register the PHP language plugin so `find_by_name` can resolve it —
    // languages live in their own crates (plugin registry) post-migration.
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_php::register);
    }
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    register_output_capture(&mut vm, &output);
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);

    let language = vybe_compiler::languages::find_by_name("php").expect("php language not found");
    let mut runtime = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    runtime
        .compile_and_run_source(src, language, virtual_path)
        .expect("run failed");

    finish_output(&output)
}

/// Compile + parse only — the original `parse` helper from the old
/// PHP tests. We don't expose the raw `Module` AST here; the caller
/// just needs "did this parse?". Returns unit on success.
pub fn parse(src: &str) {
    match vybe_language_php::parse(src) {
        Ok(_) => {}
        Err(e) => panic!("parse failed: {}", e),
    }
}

/// Returns true if the source parses successfully. Replaces the
/// `parse_ok(...)` helper used by the old `vybe_parser_php`-based tests.
pub fn parse_ok(src: &str) -> bool {
    vybe_language_php::parse(src).is_ok()
}

/// Returns true if the source parses + compiles successfully. Replaces the
/// `compile_ok(...) -> bool` helper from the old vybe_compiler_php tests
/// (named `compile_ok_check` here so it doesn't shadow the panicking
/// `compile_ok`).
pub fn compile_ok_check(src: &str) -> bool {
    compile_chunks(src).is_ok()
}
