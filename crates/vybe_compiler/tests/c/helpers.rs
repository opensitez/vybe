#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

fn compile_chunks(src: &str) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    let module = vybe_compiler::languages::c::parse(src)?;
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::c::profile_source())
            .map_err(|e| format!("profile parse failed: {}", e))?;
    vybe_compiler::compiler::Compiler::with_profile(profile).compile(&module)
}

pub fn compile_ok(src: &str) {
    match compile_chunks(src) {
        Ok(chunks) => assert!(!chunks.is_empty(), "compile produced no chunks"),
        Err(e) => panic!("compile failed: {}", e),
    }
}

pub fn run_prints(src: &str) -> Vec<String> {
    let chunks = match compile_chunks(src) {
        Ok(c) => c,
        Err(e) => panic!("compile failed: {}", e),
    };
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn(
        "wasi:cli",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
            let joined = s.join(" ");
            // C printf embeds \n in the format string; split so each line
            // becomes a separate captured entry matching test expectations.
            let mut guard = out.lock().unwrap();
            for line in joined.split('\n') {
                if !line.is_empty() {
                    guard.push(line.to_string());
                }
            }
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("run failed");
    output.lock().unwrap().clone()
}

pub fn assert_outputs(src: &str, expected: &[&str]) {
    let out = run_prints(src);
    let expected = expected.iter().map(|value| value.to_string()).collect::<Vec<_>>();
    assert_eq!(out, expected);
}

pub fn assert_program(includes: &[&str], declarations: &str, body: &str, expected: &[&str]) {
    let mut src = String::new();
    for include in includes {
        src.push_str("#include ");
        src.push_str(include);
        src.push('\n');
    }
    if !declarations.is_empty() {
        src.push_str(declarations);
        if !declarations.ends_with('\n') {
            src.push('\n');
        }
    }
    src.push_str("int main() {\n");
    src.push_str(body);
    if !body.ends_with('\n') {
        src.push('\n');
    }
    src.push_str("}\n");
    assert_outputs(&src, expected);
}

pub fn parse_ok(src: &str) -> bool {
    vybe_compiler::languages::c::parse(src).is_ok()
}
