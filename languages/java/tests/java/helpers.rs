#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_runtime::{HostContext, VM, Value};

fn compile_chunks(src: &str) -> Result<Vec<vybe_runtime::Chunk>, String> {
    static REG: std::sync::Once = std::sync::Once::new();
    REG.call_once(vybe_runtime::init_registered);
    let module = vybe_language_java::parse(src)?;
    let profile = vybe_compiler::profile::parse_profile(vybe_language_java::profile_source())
        .map_err(|e| format!("profile parse failed: {}", e))?;
    vybe_compiler::primitives::Compiler::with_profile(profile).compile(&module)
}

pub fn compile_ok(src: &str) {
    match compile_chunks(src) {
        Ok(chunks) => assert!(!chunks.is_empty(), "compile produced no chunks"),
        Err(e) => panic!("compile failed: {}", e),
    }
}

pub fn compile(src: &str) -> Vec<vybe_runtime::Chunk> {
    match compile_chunks(src) {
        Ok(c) => c,
        Err(e) => panic!("compile failed: {}", e),
    }
}

pub fn run(src: &str) -> Value {
    let chunks = compile(src);
    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    vm.run(chunks).expect("run failed")
}

pub fn run_prints(src: &str) -> Vec<String> {
    let chunks = compile(src);
    // Raw stdout fragments in call order (the C harness model). Java
    // output reaches the harness on two surfaces:
    // - `wasi:io/streams.[method]output-stream.blocking-write-and-flush` —
    //   the byte-faithful stdout path (`__j_write` / `intrinsic:
    //   write_stdout`); text is the second arg, newlines are the
    //   program's own.
    // - `wasi:logging/logging.log` — line-oriented; one record per call,
    //   newline implied (bare `println` builtin paths).
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let log_out = output.clone();
    let stream_out = output.clone();

    let (tx, rx) = std::sync::mpsc::channel();
    let handle = std::thread::spawn(move || {
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::init_platforms(&mut vm);
        vm.register_host_fn(
            "wasi:logging/logging",
            "log",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
                let mut joined = s.join(" ");
                joined.push('\n');
                log_out.lock().unwrap().push(joined);
                Value::Null
            }),
        );
        vm.register_host_fn(
            "wasi:io/streams",
            "[method]output-stream.blocking-write-and-flush",
            Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                if let Some(text) = args.get(1) {
                    let s = format!("{}", text);
                    if !s.is_empty() {
                        stream_out.lock().unwrap().push(s);
                    }
                }
                Value::Null
            }),
        );
        vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
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
    // Concatenate fragments and split into lines so each completed output
    // line is one captured entry — identical expectations to before (one
    // entry per line), independent of which sink carried the bytes.
    let joined: String = output.lock().unwrap().concat();
    if joined.is_empty() {
        return Vec::new();
    }
    let mut lines: Vec<String> = joined
        .split('\n')
        .map(|l| l.trim_end_matches('\r').to_string())
        .collect();
    if joined.ends_with('\n') {
        lines.pop();
    }
    lines
}

pub fn parse_ok(src: &str) -> bool {
    vybe_language_java::parse(src).is_ok()
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
