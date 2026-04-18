use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

pub fn run_cs(source: &str) -> Vec<String> {
    let unit = vybec::parser_csharp::parse(source).unwrap_or_else(|e| panic!("Parse error: {e}"));
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vybe_host::setup_namespaces(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    let chunks = vybec::compiler_csharp::Compiler::new().compile(&unit).unwrap_or_else(|e| panic!("Compile error: {e}"));
    vm.run(chunks).unwrap_or_else(|e| panic!("Runtime error: {e}"));
    output.lock().unwrap().clone()
}
