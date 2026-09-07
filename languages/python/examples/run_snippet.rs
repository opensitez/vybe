// Ad-hoc Python snippet runner mirroring tests/python/helpers.rs::run_python,
// used to verify language work when the shared test binary is blocked by
// unrelated in-flight files. Usage: cargo run -p vybe_language_python \
//   --example run_snippet -- path/to/snippet.py
use std::sync::{Arc, Mutex};
use vybe_runtime::{HostContext, VM, Value};

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let dump_ast = args.first().is_some_and(|arg| arg == "--dump-ast");
    let path = args
        .get(if dump_ast { 1 } else { 0 })
        .expect("usage: run_snippet <file.py>");
    let src = std::fs::read_to_string(&path).expect("read source");

    if dump_ast {
        match vybe_language_python::parse(&src) {
            Ok(module) => {
                println!("{module:#?}");
                return;
            }
            Err(e) => {
                eprintln!("Parse error: {e}");
                std::process::exit(1);
            }
        }
    }

    vybe_language_python::register();
    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);

    let out = output.clone();
    vm.register_host_fn(
        "web:console",
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
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);

    let language = vybe_compiler::languages::find_by_name("python").expect("python language");
    let mut runtime = vybe_compiler::dynamic::RuntimeCompilerService::new(&mut vm);
    match runtime.compile_and_run_source(&src, language, "snippet.py") {
        Ok(_) => {
            let joined: String = output.lock().unwrap().concat();
            print!("{joined}");
        }
        Err(e) => {
            let joined: String = output.lock().unwrap().concat();
            print!("{joined}");
            eprintln!("RUN ERROR: {e:?}");
            std::process::exit(1);
        }
    }
}
