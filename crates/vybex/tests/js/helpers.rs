use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

/// Run JS source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_js(src: &str) -> Vec<String> {
    let module = vybex::languages::js::parse(src).expect("JS parse failed");

    let profile = vybex::profile::parse_profile(vybex::languages::js::profile_source())
        .expect("Failed to parse JS profile");

    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("JS compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("JS run failed");
    let result = output.lock().unwrap().clone();
    result
}

/// Run JS source through the full pipeline including ESM host-import
/// installation. Needed for tests that use `import { X } from "wasi:*"`
/// and then read X as a value or use `import * as ns`.
pub fn run_js_with_imports(src: &str) -> Vec<String> {
    let module = vybex::languages::js::parse(src).expect("JS parse failed");
    let profile = vybex::profile::parse_profile(vybex::languages::js::profile_source())
        .expect("Failed to parse JS profile");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vybex::adapters::register_all(&mut vm).expect("adapter registration failed");

    let module_exports = vybex::bundle::flatten_module_exports(&vm.modules);
    let result = vybex::compiler::Compiler::with_profile(profile)
        .with_module_exports(module_exports)
        .compile_with_imports(&module).expect("JS compile failed");

    vybex::host_imports::install(&mut vm, &result.host_imports);
    vm.run(result.chunks).expect("JS run failed");
    let lines = output.lock().unwrap().clone();
    lines
}

/// Run JS source, return (VM, output) for post-run inspection.
pub fn run_js_vm(src: &str) -> (VM, Arc<Mutex<Vec<String>>>) {
    let module = vybex::languages::js::parse(src).expect("JS parse failed");
    let profile = vybex::profile::parse_profile(vybex::languages::js::profile_source())
        .expect("Failed to parse JS profile");
    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("JS compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("JS run failed");
    (vm, output)
}
