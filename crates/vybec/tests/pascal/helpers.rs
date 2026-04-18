use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

pub fn run(src: &str) -> Vec<String> {
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
    let program = vybec::parser_pascal::parse(src).expect("parse");
    let chunks = vybec::compiler_pascal::Compiler::new().compile(&program).expect("compile");
    vm.run(chunks).expect("run");
    output.lock().unwrap().clone()
}
