#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_runtime::{HostContext, VM, Value};

#[macro_export]
macro_rules! kotlin_run_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = $crate::helpers::run_prints($src);
            assert_eq!(out, $expected);
        }
    };
}

#[macro_export]
macro_rules! kotlin_compile_test {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            $crate::helpers::compile_ok($src);
        }
    };
}

#[macro_export]
macro_rules! kotlin_run_cases {
    ($($name:ident => ($src:expr, $expected:expr $(,)?),)+) => {
        $(kotlin_run_test!($name, $src, $expected);)+
    };
}

#[macro_export]
macro_rules! kotlin_compile_cases {
    ($($name:ident => $src:expr,)+) => {
        $(kotlin_compile_test!($name, $src);)+
    };
}

fn compile_chunks(src: &str) -> Result<Vec<vybe_runtime::Chunk>, String> {
    static REG: std::sync::Once = std::sync::Once::new();
    REG.call_once(vybe_language_kotlin::register);
    let module = vybe_language_kotlin::parse(src)?;
    let profile = vybe_compiler::profile::parse_profile(vybe_language_kotlin::profile_source())
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
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let s: Vec<String> = args.iter().map(|a| format!("{}", a)).collect();
            out.lock().unwrap().push(s.join(" "));
            Value::Null
        }),
    );
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("run failed");
    let result = output.lock().unwrap().clone();
    result
}

pub fn parse_ok(src: &str) -> bool {
    vybe_language_kotlin::parse(src).is_ok()
}

pub fn compile_ok_check(src: &str) -> bool {
    compile_chunks(src).is_ok()
}
