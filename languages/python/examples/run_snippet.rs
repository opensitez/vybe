// Ad-hoc Python snippet runner mirroring tests/python/helpers.rs::run_python,
// used to verify language work when the shared test binary is blocked by
// unrelated in-flight files. Usage: cargo run -p vybe_language_python \
//   --example run_snippet -- path/to/snippet.py
use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

fn main() {
    let path = std::env::args().nth(1).expect("usage: run_snippet <file.py>");
    let src = std::fs::read_to_string(&path).expect("read source");

    vybe_language_python::register();
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    vybe_emitter::platforms::init_platforms(&mut vm);

    let out = output.clone();
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let mut joined = args
                .iter()
                .map(|v| format!("{v}"))
                .collect::<Vec<_>>()
                .join(" ");
            joined.push('\n');
            out.lock().unwrap().push(joined);
            Value::Null
        }),
    );
    let out = output.clone();
    vm.register_host_fn(
        "wasi:cli/stdout",
        "write-via-stream",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let stream_val = args.first().cloned().unwrap_or(Value::Null);
            let bytes = ctx.stream_drain(&stream_val);
            if !bytes.is_empty() {
                out.lock()
                    .unwrap()
                    .push(String::from_utf8_lossy(&bytes).into_owned());
            }
            Value::Null
        }),
    );
    vybe_emitter::platforms::finalize_platforms(&mut vm);

    let language = vybe_compiler::languages::find_by_name("python").expect("python language");
    let mut runtime = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    match runtime.compile_and_run_source(&src, language, "snippet.py") {
        Ok(_) => {
            let joined: String = output.lock().unwrap().concat();
            print!("{joined}");
        }
        Err(e) => {
            eprintln!("RUN ERROR: {e:?}");
            std::process::exit(1);
        }
    }
}
