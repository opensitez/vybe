use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

fn compile_chunks(src: &str) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_lua::register);
    }
    let module = vybe_language_lua::parse(src)?;
    let profile = vybe_compiler::profile::parse_profile(vybe_language_lua::profile_source())
        .map_err(|e| format!("profile parse failed: {e}"))?;
    vybe_compiler::compiler::Compiler::with_profile(profile).compile(&module)
}

pub fn parse_ok(src: &str) {
    vybe_language_lua::parse(src).expect("Lua parse failed");
}

pub fn compile_ok(src: &str) {
    match compile_chunks(src) {
        Ok(chunks) => assert!(!chunks.is_empty(), "compile produced no chunks"),
        Err(e) => panic!("compile failed: {e}"),
    }
}

pub fn run_lua(src: &str) -> Vec<String> {
    let chunks = compile_chunks(src).expect("Lua compile failed");
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_emitter::platforms::init_platforms(&mut vm);
    // `emit = "print"` → emitter/io.rs → wasi:logging/logging.log
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    vybe_emitter::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("Lua run failed");
    output.lock().unwrap().clone()
}

pub fn run_lua_one(src: &str) -> String {
    run_lua(src).into_iter().next().unwrap_or_default()
}

/// Assert `print(...)` produces one line. Written from Lua semantics, not compiler internals.
macro_rules! lua_print {
    ($($name:ident => { $src:expr, $expect:expr }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_eq!($crate::helpers::run_lua_one($src), $expect);
            }
        )*
    };
}
