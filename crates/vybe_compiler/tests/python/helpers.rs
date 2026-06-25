use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

/// Run Python source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_python(src: &str) -> Vec<String> {
    let module = vybe_compiler::languages::python::parse(src).expect("Python parse failed");

    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::python::profile_source())
            .expect("Failed to parse Python profile");

    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Python compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("Python run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn run_python_one(src: &str) -> String {
    run_python(src).into_iter().next().unwrap_or_default()
}

/// Parse-only: verify the grammar accepts the source without errors
pub fn parse_ok(src: &str) {
    vybe_compiler::languages::python::parse(src).expect("Python parse failed");
}

/// Parse + compile: verify the full pipeline up to bytecode emission
pub fn compile_ok(src: &str) {
    let module = vybe_compiler::languages::python::parse(src).expect("Python parse failed");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::python::profile_source())
            .expect("Failed to parse Python profile");
    let _chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Python compile failed");
}

/// Shorthand: `print(expr)` then return the single logged line.
pub fn run_print(expr: &str) -> String {
    run_python_one(&format!("print({expr})\n"))
}
