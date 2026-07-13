use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

#[macro_export]
macro_rules! runtime_case {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!($crate::helpers::run_python_one($src), $expected);
        }
    };
}

#[macro_export]
macro_rules! compile_case {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            $crate::helpers::compile_ok($src);
        }
    };
}

// Output capture mirrors the two real output surfaces (same model as the PHP
// harness):
//
// - `wasi:cli/stdout.write-via-stream(data: stream<u8>)` — Python `print`
//   writes raw stdout bytes here with NO implicit newline, so consecutive
//   writes concatenate onto the current line (`print('x', end='')` then
//   `print('y')` → "xy\n"). The newline is part of `print`'s `end` argument.
// - `wasi:logging/logging.log` — line-oriented logging; one record per call.
//
// Fragments are buffered in call order and split into lines in `finish_output`.
fn register_output_capture(vm: &mut VM, output: &Arc<Mutex<Vec<String>>>) {
    let out = output.clone();
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let mut joined = args
                .iter()
                .map(|v| format!("{v}"))
                .collect::<Vec<_>>()
                .join(" ");
            joined.push('\n');
            out.lock().unwrap().push(joined);
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
                out.lock()
                    .unwrap()
                    .push(String::from_utf8_lossy(&bytes).into_owned());
            }
            Value::Null
        }),
    );
}

/// Concatenate the buffered fragments (in call order) and split into lines.
/// A trailing newline (the default `print` end) produces a final empty
/// element, which is dropped.
fn finish_output(output: &Arc<Mutex<Vec<String>>>) -> Vec<String> {
    let joined: String = output.lock().unwrap().concat();
    let mut lines: Vec<String> = joined
        .split('\n')
        .map(|l| l.trim_end_matches('\r').to_string())
        .collect();
    while lines.last().map_or(false, |l| l.is_empty()) {
        lines.pop();
    }
    lines
}

/// Run Python source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_python(src: &str) -> Vec<String> {
    // Run through the RuntimeCompilerService like the real exe (and the PHP
    // harness) — a bare `vm.run` never activates the compiler-as-a-service, so
    // `eval`/`exec`/`compile` (which route to the `vybe:eval` host) break.
    { static R: std::sync::Once = std::sync::Once::new(); R.call_once(vybe_language_python::register); }
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    vybe_host::register_all(&mut vm);
    register_output_capture(&mut vm, &output);
    vybe_host::setup_namespaces(&mut vm);

    let language =
        vybe_compiler::languages::find_by_name("python").expect("python language not found");
    let mut runtime = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    runtime
        .compile_and_run_source(src, language, "test.py")
        .expect("Python run failed");
    finish_output(&output)
}

/// The program's full stdout, lines joined by "\n" (Python's line separator).
pub fn run_python_one(src: &str) -> String {
    run_python(src).join("\n")
}

/// Parse-only: verify the grammar accepts the source without errors
pub fn parse_ok(src: &str) {
    vybe_language_python::parse(src).expect("Python parse failed");
}

/// Parse + compile: verify the full pipeline up to bytecode emission
pub fn compile_ok(src: &str) {
    { static R: std::sync::Once = std::sync::Once::new(); R.call_once(vybe_language_python::register); }
    let module = vybe_language_python::parse(src).expect("Python parse failed");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_python::profile_source())
            .expect("Failed to parse Python profile");
    let _chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Python compile failed");
}

/// Shorthand: `print(expr)` then return the single logged line.
pub fn run_print(expr: &str) -> String {
    run_python_one(&format!("print({expr})\n"))
}
