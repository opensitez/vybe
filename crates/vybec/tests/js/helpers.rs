use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

pub fn run_js(code: &str) -> Vec<String> {
    let program = vybec::parser_js::parse(code).expect("parse failed");
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vybec::compiler_js::register_js_coercion(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    let chunks = vybec::compiler_js::Compiler::new().compile(&program).expect("compile failed");
    vm.run(chunks).expect("runtime error");
    output.lock().unwrap().clone()
}
