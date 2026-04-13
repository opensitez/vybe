#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

fn compile_chunks(src: &str) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    let module = vybex::languages::cobol::parse(src)?;
    let profile = vybex::profile::parse_profile(vybex::languages::cobol::profile_source())
        .map_err(|e| format!("profile parse failed: {}", e))?;
    vybex::compiler::Compiler::with_profile(profile).compile(&module)
}

pub fn compile_ok(src: &str) {
    match compile_chunks(src) {
        Ok(chunks) => {
            assert!(!chunks.is_empty(), "compile produced no chunks");
        }
        Err(e) => panic!("compile failed: {}", e),
    }
}

pub fn compile(src: &str) -> Vec<vybe_bytecode::Chunk> {
    match compile_chunks(src) {
        Ok(c) => c,
        Err(e) => panic!("compile failed: {}", e),
    }
}

pub fn run(src: &str) -> Value {
    let chunks = compile(src);
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    vm.run(chunks).expect("run failed")
}

pub fn run_prints(src: &str) -> Vec<String> {
    let chunks = compile(src);
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
        out.lock().unwrap().push(s.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn parse_ok(src: &str) -> bool {
    vybex::languages::cobol::parse(src).is_ok()
}

pub fn compile_ok_check(src: &str) -> bool {
    compile_chunks(src).is_ok()
}
