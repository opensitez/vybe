#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, Value, VM};

fn compile_chunks(src: &str) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    let module = vybe_compiler::languages::java::parse(src)?;
    let profile =
        vybe_compiler::profile::parse_profile(vybe_compiler::languages::java::profile_source())
            .map_err(|e| format!("profile parse failed: {}", e))?;
    vybe_compiler::compiler::Compiler::with_profile(profile).compile(&module)
}

pub fn compile_ok(src: &str) {
    match compile_chunks(src) {
        Ok(chunks) => assert!(!chunks.is_empty(), "compile produced no chunks"),
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
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();

    let (tx, rx) = std::sync::mpsc::channel();
    let handle = std::thread::spawn(move || {
        let mut vm = VM::new();
        vybe_host::register_all(&mut vm);
        vm.register_host_fn(
            "wasi:logging/logging",
            "log",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
                out.lock().unwrap().push(s.join(" "));
                Value::Null
            }),
        );
        vybe_host::setup_namespaces(&mut vm);
        let result = vm.run(chunks);
        let _ = tx.send(result);
    });

    match rx.recv_timeout(std::time::Duration::from_secs(5)) {
        Ok(result) => {
            result.expect("run failed");
            handle.join().unwrap();
        }
        Err(_) => panic!("run failed: timed out after 5s (probable infinite loop)"),
    }
    output.lock().unwrap().clone()
}

pub fn parse_ok(src: &str) -> bool {
    vybe_compiler::languages::java::parse(src).is_ok()
}

pub fn compile_ok_check(src: &str) -> bool {
    compile_chunks(src).is_ok()
}

/// Wrap `body` in a standard `public class Main` with `main` and run it.
pub fn run_main(body: &str) -> Vec<String> {
    run_prints(&format!(
        "public class Main {{ public static void main(String[] args) {{ {body} }} }}"
    ))
}

/// Run `main_body` inside `Main` with extra nested type definitions.
pub fn run_in_main(main_body: &str, type_defs: &str) -> Vec<String> {
    run_prints(&format!(
        "public class Main {{ {type_defs} public static void main(String[] args) {{ {main_body} }} }}"
    ))
}
