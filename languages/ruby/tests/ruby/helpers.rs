use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

/// Run Ruby source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_ruby(src: &str) -> Vec<String> {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_ruby::register);
    }

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_compiler::compiler::platforms::init_platforms(&mut vm);
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    vybe_compiler::compiler::platforms::finalize_platforms(&mut vm);
    let language = vybe_compiler::languages::find_by_name("ruby").expect("ruby language not found");
    let mut runtime = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    runtime
        .compile_and_run_source(src, language, "test.rb")
        .expect("Ruby run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn run_ruby_one(src: &str) -> String {
    run_ruby(src).join("\n")
}

/// Parse-only: verify the grammar accepts the source without errors
#[allow(dead_code)]
pub fn parse_ok(src: &str) {
    vybe_language_ruby::parse(src).expect("Ruby parse failed");
}

/// Parse + compile: verify the full pipeline up to bytecode emission
pub fn compile_ok(src: &str) {
    let module = vybe_language_ruby::parse(src).expect("Ruby parse failed");
    let profile = vybe_compiler::profile::parse_profile(vybe_language_ruby::profile_source())
        .expect("Failed to parse Ruby profile");
    let _chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Ruby compile failed");
}
