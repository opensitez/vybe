use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

/// Run Pascal source through vybex pipeline: pest grammar -> walker -> common AST -> compiler -> VM
pub fn run_pascal(src: &str) -> Vec<String> {
    let module = vybex::languages::pascal::parse(src).expect("Pascal parse failed");

    let profile = load_pascal_profile();

    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("Pascal compile failed");

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
    vm.run(chunks).expect("Pascal run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn load_pascal_profile() -> vybex::profile::LanguageProfile {
    vybex::profile::parse_profile(vybex::languages::pascal::profile_source())
        .expect("Failed to parse Pascal profile")
}
