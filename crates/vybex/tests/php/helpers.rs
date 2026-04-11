//! PHP test helpers — compile through the vybex pipeline.
//!
//! The original `vybe_compiler_php` test files use four entry points:
//! `parse`, `compile`, `compile_ok`, `run`, `run_prints`. Every test
//! file currently in this suite uses only `compile_ok` (parse + compile
//! must succeed; runtime is not asserted), so we provide that shim
//! plus the others as drop-in replacements.
//!
//! `compile_ok(src)` runs `vybex::languages::php::parse` → walker →
//! `vybex::compiler::Compiler::with_profile` and asserts the result is
//! Ok. Any parse error, walker error, or compile error fails the test
//! with the underlying message.

#![allow(dead_code)]

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

fn compile_chunks(src: &str) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    let module = vybex::languages::php::parse(src)?;
    let profile = vybex::profile::parse_profile(vybex::languages::php::profile_source())
        .map_err(|e| format!("profile parse failed: {}", e))?;
    vybex::compiler::Compiler::with_profile(profile).compile(&module)
}

/// Asserts that `src` parses + compiles cleanly. Used by ~all PHP test
/// files; the original tests don't run the bytecode, just verify the
/// pipeline accepts the source.
pub fn compile_ok(src: &str) {
    match compile_chunks(src) {
        Ok(chunks) => {
            assert!(!chunks.is_empty(), "compile produced no chunks");
        }
        Err(e) => panic!("compile failed: {}", e),
    }
}

/// Returns the compiled chunks, or panics with the error.
pub fn compile(src: &str) -> Vec<vybe_bytecode::Chunk> {
    match compile_chunks(src) {
        Ok(c) => c,
        Err(e) => panic!("compile failed: {}", e),
    }
}

/// Compile + run, return the final value popped from the stack.
pub fn run(src: &str) -> Value {
    let chunks = compile(src);
    let mut vm = VM::new();
    vybe_host::register_all(&mut vm);
    vm.run(chunks).expect("run failed")
}

/// Compile + run, capture stdout (anything routed through `wasi:cli::log`),
/// return the captured lines.
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

/// Compile + parse only — the original `parse` helper from the old
/// PHP tests. We don't expose the raw `Module` AST here; the caller
/// just needs "did this parse?". Returns unit on success.
pub fn parse(src: &str) {
    match vybex::languages::php::parse(src) {
        Ok(_) => {}
        Err(e) => panic!("parse failed: {}", e),
    }
}

/// Returns true if the source parses successfully. Replaces the
/// `parse_ok(...)` helper used by the old `vybe_parser_php`-based tests.
pub fn parse_ok(src: &str) -> bool {
    vybex::languages::php::parse(src).is_ok()
}

/// Returns true if the source parses + compiles successfully. Replaces the
/// `compile_ok(...) -> bool` helper from the old vybe_compiler_php tests
/// (named `compile_ok_check` here so it doesn't shadow the panicking
/// `compile_ok`).
pub fn compile_ok_check(src: &str) -> bool {
    compile_chunks(src).is_ok()
}
