//! Drive the VM's type recorder over a suite of real-language programs
//! and dump per-slot histograms. The output tells us, empirically, how
//! much of the anyref/ABI migration's Phase 3 (typed locals) can ever
//! win: the fraction of slots that are monomorphic at runtime is a
//! hard upper bound on the opcodes we can skip the box/unbox dance on.
//!
//! Run with:
//!   `cargo test -p vybex --test abi_instrumentation -- --nocapture`
//!
//! These are not pass/fail tests — they always pass; they print.

use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

fn compile_python(src: &str) -> Vec<vybe_bytecode::Chunk> {
    let module = vybe_language_python::parse(src).expect("parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_python::profile_source())
            .expect("profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("compile")
}

fn compile_js(src: &str) -> Vec<vybe_bytecode::Chunk> {
    let module = vybe_language_js::parse(src).expect("parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_js::profile_source())
            .expect("profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("compile")
}

fn compile_vb(src: &str) -> Vec<vybe_bytecode::Chunk> {
    let module = vybe_language_vb::parse(src).expect("parse");
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_vb::profile_source())
            .expect("profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("compile")
}

fn run_with_recorder(chunks: Vec<vybe_bytecode::Chunk>) -> (Vec<String>, String) {
    let chunk_names: Vec<String> = chunks.iter().map(|c| c.name.clone()).collect();
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    vm.record_types(true);
    let _ = vm.run(chunks);
    let rec = vm.take_type_record().unwrap();
    let report = rec.format_report(&chunk_names);
    let logs = output.lock().unwrap().clone();
    (logs, report)
}

#[test]
fn python_fib_monomorphism_report() {
    let chunks = compile_python(
        r#"
def fib(n):
    if n <= 1:
        return n
    return fib(n - 1) + fib(n - 2)

for i in range(8):
    print(fib(i))
"#,
    );
    let (_out, report) = run_with_recorder(chunks);
    println!("\n=== Python fib recorder report ===\n{report}");
}

#[test]
fn js_tight_loop_monomorphism_report() {
    let chunks = compile_js(
        r#"
let total = 0;
for (let i = 0; i < 100; i++) {
    total = total + i * 2;
}
console.log(total);
"#,
    );
    let (_out, report) = run_with_recorder(chunks);
    println!("\n=== JS tight loop recorder report ===\n{report}");
}

#[test]
fn vb_mixed_types_recorder_report() {
    let chunks = compile_vb(
        r#"
Module Program
    Sub Main()
        Dim n As Integer = 10
        Dim s As String = "hello"
        Dim total As Integer = 0
        For i As Integer = 1 To n
            total = total + i
        Next
        Console.WriteLine(total)
        Console.WriteLine(s)
    End Sub
End Module
"#,
    );
    let (_out, report) = run_with_recorder(chunks);
    println!("\n=== VB mixed-types recorder report ===\n{report}");
}
