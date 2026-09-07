//! Run one washm/bash file through the current language crate.
//!
//! This is deliberately scoped to `vybe_language_washm` so bash debugging does
//! not depend on rebuilding the repo-wide `vybex` binary.

use std::path::PathBuf;
use std::sync::{Arc, Mutex};

use vybe_compiler::bundle::{Bundle, EntryPoint, SourceFile};
use vybe_runtime::{HostContext, VM, Value};

fn main() {
    let Some(path) = std::env::args().nth(1).map(PathBuf::from) else {
        eprintln!("usage: run_washm <file.sh>");
        std::process::exit(2);
    };
    let code = match std::fs::read_to_string(&path) {
        Ok(code) => code,
        Err(e) => {
            eprintln!("read {}: {e}", path.display());
            std::process::exit(1);
        }
    };

    vybe_language_washm::register();
    let language = vybe_runtime::registry::find("bash").expect("bash language registered");
    let bundle = Bundle {
        name: path.display().to_string(),
        language,
        sources: vec![SourceFile {
            path: path.clone(),
            code,
        }],
        wasm_files: Vec::new(),
        entry_point: EntryPoint::Auto,
    };

    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
    vybe_compiler::dynamic::register_dynamic_runtime_imports(&mut vm);

    let output = Arc::new(Mutex::new(String::new()));
    let stdout = Arc::clone(&output);
    vm.register_host_fn(
        "wasi:cli/stdout",
        "write-via-stream",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let stream_val = args.first().cloned().unwrap_or(Value::Null);
            let bytes = ctx.stream_drain(&stream_val);
            if !bytes.is_empty() {
                stdout
                    .lock()
                    .unwrap()
                    .push_str(&String::from_utf8_lossy(&bytes));
            }
            let (fut, fut_id) = ctx.create_future();
            ctx.resolve_future(fut_id, Value::Null);
            fut
        }),
    );

    let compiled = match bundle.compile_full() {
        Ok(compiled) => compiled,
        Err(e) => {
            eprintln!("compile {}: {e}", path.display());
            std::process::exit(1);
        }
    };
    if let Err(e) = vm.run(compiled.chunks) {
        eprintln!("runtime {}: {e}", path.display());
        eprintln!("runtime debug: {e:?}");
        std::process::exit(1);
    }
    print!("{}", output.lock().unwrap());
    if vm.pending_exit_code != 0 {
        std::process::exit(vm.pending_exit_code);
    }
}
