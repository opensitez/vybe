use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

/// Parse WAST/WAT source and return the Module (parse-only check).
pub fn parse_ok(src: &str) {
    vybe_compiler::languages::wast::parse(src)
        .unwrap_or_else(|e| panic!("WAST parse failed:\n{}\nSource:\n{}", e, src));
}

/// Parse and expect a parse error (negative test).
pub fn parse_err(src: &str) {
    assert!(
        vybe_compiler::languages::wast::parse(src).is_err(),
        "Expected parse error but succeeded for:\n{}",
        src
    );
}

/// Parse + compile through the full pipeline (no VM execution).
pub fn compile_ok(src: &str) {
    let module = vybe_compiler::languages::wast::parse(src)
        .unwrap_or_else(|e| panic!("WAST parse failed:\n{}", e));
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::wast::profile_source())
            .expect("Failed to parse WAST profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap_or_else(|e| panic!("WAST compile failed:\n{}", e));
}

/// Run WAST source through the full pipeline and return console output lines.
pub fn run_wast(src: &str) -> Vec<String> {
    let module = vybe_compiler::languages::wast::parse(src)
        .unwrap_or_else(|e| panic!("WAST parse failed:\n{}", e));
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::wast::profile_source())
            .expect("Failed to parse WAST profile");
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap_or_else(|e| panic!("WAST compile failed:\n{}", e));

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn(
        "wasi:cli",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks)
        .unwrap_or_else(|e| panic!("WAST run failed:\n{}", e));
    output.lock().unwrap().clone()
}

pub fn run_wast_one(src: &str) -> String {
    run_wast(src).into_iter().next().unwrap_or_default()
}
