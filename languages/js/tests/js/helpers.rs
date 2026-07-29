use std::sync::{Arc, Mutex};
use vybe_runtime::{HostContext, VM, Value};

/// Register the test output-capture bindings. Mirrors the real host's two
/// output surfaces so assertions see what a user would:
/// - `wasi:cli/stdout.write-via-stream` — where `console.log`/print now
///   write (concatenated args + "\n" per call); drained and split to lines.
/// - `wasi:logging/logging.log` — leveled logging; one line per call.
fn register_output_capture(vm: &mut VM, output: &Arc<Mutex<Vec<String>>>) {
    let out = output.clone();
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out.lock().unwrap().push(parts.join(" "));
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
                let trimmed = text.strip_suffix('\n').unwrap_or(&text);
                let mut vec = out.lock().unwrap();
                for line in trimmed.split('\n') {
                    vec.push(line.to_string());
                }
            }
            Value::Null
        }),
    );
}

#[macro_export]
macro_rules! js_cases {
    ($($name:ident => { $src:expr, [$($expected:expr),* $(,)?] };)+) => {
        $(
            #[test]
            fn $name() {
                $crate::helpers::assert_js($src, &[$($expected),*]);
            }
        )+
    };
}

#[macro_export]
macro_rules! js_import_cases {
    ($($name:ident => { $src:expr, [$($expected:expr),* $(,)?] };)+) => {
        $(
            #[test]
            fn $name() {
                $crate::helpers::assert_js_with_imports($src, &[$($expected),*]);
            }
        )+
    };
}

/// Run JS source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_js(src: &str) -> Vec<String> {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_js::register);
    }
    let module = vybe_language_js::parse(src).expect("JS parse failed");

    let profile = vybe_compiler::profile::parse_profile(vybe_language_js::profile_source())
        .expect("Failed to parse JS profile");

    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("JS compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    register_output_capture(&mut vm, &output);
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vybe_compiler::dynamic::run_with_js_dynamic_runtime(
        &mut vm,
        vybe_runtime::capabilities::Capabilities::all(),
        chunks,
    )
    .expect("JS run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn assert_js(src: &str, expected: &[&str]) {
    let actual = run_js(src);
    let expected_vec: Vec<String> = expected.iter().map(|line| (*line).to_string()).collect();
    assert_eq!(actual, expected_vec);
}

/// Run JS source through the full pipeline including ESM host-import
/// installation. Needed for tests that use `import { X } from "wasi:*"`
/// and then read X as a value or use `import * as ns`.
pub fn run_js_with_imports(src: &str) -> Vec<String> {
    let module = vybe_language_js::parse(src).expect("JS parse failed");
    let profile = vybe_compiler::profile::parse_profile(vybe_language_js::profile_source())
        .expect("Failed to parse JS profile");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    register_output_capture(&mut vm, &output);
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vybe_compiler::adapters::register_all(&mut vm).expect("adapter registration failed");

    let module_exports = vybe_compiler::bundle::flatten_module_exports(&vm.modules);
    let value_exports = vybe_compiler::bundle::flatten_module_value_exports(&vm.modules);
    let result = vybe_compiler::primitives::Compiler::with_profile(profile)
        .with_module_exports(module_exports)
        .with_module_value_exports(value_exports)
        .compile_with_imports(&module)
        .expect("JS compile failed");

    vybe_compiler::host_imports::install(&mut vm, &result.host_imports);
    vybe_compiler::dynamic::run_with_js_dynamic_runtime(
        &mut vm,
        vybe_runtime::capabilities::Capabilities::all(),
        result.chunks,
    )
    .expect("JS run failed");
    let lines = output.lock().unwrap().clone();
    lines
}

pub fn assert_js_with_imports(src: &str, expected: &[&str]) {
    let actual = run_js_with_imports(src);
    let expected_vec: Vec<String> = expected.iter().map(|line| (*line).to_string()).collect();
    assert_eq!(actual, expected_vec);
}

/// Run JS source, return (VM, output) for post-run inspection.
pub fn run_js_vm(src: &str) -> (VM, Arc<Mutex<Vec<String>>>) {
    let module = vybe_language_js::parse(src).expect("JS parse failed");
    let profile = vybe_compiler::profile::parse_profile(vybe_language_js::profile_source())
        .expect("Failed to parse JS profile");
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("JS compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    register_output_capture(&mut vm, &output);
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vybe_compiler::dynamic::run_with_js_dynamic_runtime(
        &mut vm,
        vybe_runtime::capabilities::Capabilities::all(),
        chunks,
    )
    .expect("JS run failed");
    (vm, output)
}
