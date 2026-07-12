use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

/// Parse WAST/WAT source and return the Module (parse-only check).
pub fn parse_ok(src: &str) {
    vybe_language_wast::parse(src)
        .unwrap_or_else(|e| panic!("WAST parse failed:\n{}\nSource:\n{}", e, src));
}

/// Parse and expect a parse error (negative test).
#[allow(dead_code)]
pub fn parse_err(src: &str) {
    assert!(
        vybe_language_wast::parse(src).is_err(),
        "Expected parse error but succeeded for:\n{}",
        src
    );
}

/// Parse + compile through the full pipeline (no VM execution).
pub fn compile_ok(src: &str) {
    { static R: std::sync::Once = std::sync::Once::new(); R.call_once(vybe_language_wast::register); }
    let module = vybe_language_wast::parse(src)
        .unwrap_or_else(|e| panic!("WAST parse failed:\n{}", e));
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_wast::profile_source())
            .expect("Failed to parse WAST profile");
    vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap_or_else(|e| panic!("WAST compile failed:\n{}", e));
}

/// Run WAST source through the full pipeline and return console output lines.
pub fn run_wast(src: &str) -> Vec<String> {
    let module = vybe_language_wast::parse(src)
        .unwrap_or_else(|e| panic!("WAST parse failed:\n{}", e));
    let profile =
        vybe_compiler::profile::parse_profile(vybe_language_wast::profile_source())
            .expect("Failed to parse WAST profile");
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .unwrap_or_else(|e| panic!("WAST compile failed:\n{}", e));

    for (i, c) in chunks.iter().enumerate() {
        println!("Chunk {i}:\n{}", vybe_bytecode::debug::disassemble(c));
    }

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    let out_cloned = out.clone();
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out_cloned.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    let out_cloned2 = out.clone();
    vm.register_host_fn(
        "wasi:cli",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out_cloned2.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks)
        .unwrap_or_else(|e| panic!("WAST run failed:\n{}", e));
    output.lock().unwrap().clone()
}

pub fn run_wast_one(src: &str) -> String {
    run_wast(src).into_iter().next().unwrap_or_default()
}

#[macro_export]
macro_rules! wat_exec {
    ($($name:ident => { $src:expr, $expect:expr }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                let wrapped_src = format!(
                    r#"
(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  {}
)
"#,
                    $src
                );

                // If it already is a module, don't wrap it. Just assume if it starts with (module it's complete.
                let final_src = if $src.trim().starts_with("(module") {
                    $src.to_string()
                } else {
                    wrapped_src
                };

                let out = $crate::helpers::run_wast_one(&final_src);
                assert_eq!(out, $expect);
            }
        )*
    };
}
