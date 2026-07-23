#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

#[macro_export]
macro_rules! fortran_cases {
    ($($name:ident => { $src:expr, [$($expected:expr),* $(,)?] };)+) => {
        $(
            #[test]
            fn $name() {
                $crate::helpers::assert_fortran($src, &[$($expected),*]);
            }
        )+
    };
}

fn compile_chunks(src: &str) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_fortran::register);
    }
    let module = vybe_language_fortran::parse(src)?;
    let profile = vybe_compiler::profile::parse_profile(vybe_language_fortran::profile_source())
        .map_err(|e| format!("profile parse failed: {}", e))?;
    vybe_compiler::compiler::Compiler::with_profile(profile).compile(&module)
}

pub fn compile_ok(src: &str) {
    match compile_chunks(src) {
        Ok(chunks) => {
            assert!(!chunks.is_empty(), "compile produced no chunks");
        }
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
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
            out.lock().unwrap().push(s.join(" "));
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn parse_ok(src: &str) -> bool {
    vybe_language_fortran::parse(src).is_ok()
}

pub fn assert_fortran(src: &str, expected: &[&str]) {
    let actual = run_prints(src);
    let expected_vec: Vec<String> = expected.iter().map(|line| (*line).to_string()).collect();
    assert_eq!(actual, expected_vec);
}
