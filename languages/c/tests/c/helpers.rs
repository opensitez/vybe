#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

/// Build a complete C program from includes, declarations, and main body.
pub fn program_src(includes: &[&str], declarations: &str, body: &str) -> String {
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
    src
}

#[macro_export]
macro_rules! c_run_cases {
    ($($name:ident => { includes: [$($inc:literal),* $(,)?], decls: $decls:expr, body: $body:expr, expect: [$($exp:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                $crate::helpers::assert_program(&[$($inc),*], $decls, $body, &[$($exp),*]);
            }
        )*
    };
}

#[macro_export]
macro_rules! c_compile_cases {
    ($($name:ident => { includes: [$($inc:literal),* $(,)?], decls: $decls:expr, body: $body:expr }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                let src = $crate::helpers::program_src(&[$($inc),*], $decls, $body);
                $crate::helpers::compile_ok(&src);
            }
        )*
    };
}

fn compile_chunks(src: &str) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_c::register);
    }
    let module = vybe_language_c::parse(src)?;
    let profile = vybe_compiler::profile::parse_profile(vybe_language_c::profile_source())
        .map_err(|e| format!("profile parse failed: {}", e))?;
    vybe_compiler::primitives::Compiler::with_profile(profile).compile(&module)
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
    // Raw stdout fragments in call order. libc output reaches the harness on
    // two surfaces:
    // - `wasi:io/streams.[method]output-stream.blocking-write-and-flush` —
    //   the byte-faithful libc stdout path (`intrinsic:write_stdout`); text
    //   is the second arg, newlines are the program's own.
    // - `wasi:logging/logging.log` — line-oriented; one record per call,
    //   newline implied.
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    let out = output.clone();
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
            let mut joined = s.join(" ");
            joined.push('\n');
            out.lock().unwrap().push(joined);
            Value::Null
        }),
    );
    let out = output.clone();
    vm.register_host_fn(
        "wasi:io/streams",
        "[method]output-stream.blocking-write-and-flush",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            if let Some(text) = args.get(1) {
                let s = format!("{}", text);
                if !s.is_empty() {
                    out.lock().unwrap().push(s);
                }
            }
            Value::Null
        }),
    );
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("run failed");
    // Concatenate fragments and split into lines so each printf line becomes
    // one captured entry. Strip only the final empty artifact of a trailing
    // newline — interior empties are real content (`puts("")` → "").
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

pub fn assert_outputs(src: &str, expected: &[&str]) {
    let out = run_prints(src);
    let expected = expected
        .iter()
        .map(|value| value.to_string())
        .collect::<Vec<_>>();
    assert_eq!(out, expected);
}

pub fn assert_program(includes: &[&str], declarations: &str, body: &str, expected: &[&str]) {
    assert_outputs(&program_src(includes, declarations, body), expected);
}

pub fn parse_ok(src: &str) -> bool {
    vybe_language_c::parse(src).is_ok()
}
