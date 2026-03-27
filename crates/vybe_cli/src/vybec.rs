//! vybec — Vybe bytecode compiler and runner.
//!
//! Compiles .vb and .js files to the same bytecode VM and runs them.
//! Both languages share the same host functions and namespace objects.
//! If the program creates a GUI (RunApplication), a Dioxus window is launched.
//!
//! Usage:
//!   vybec hello.vb          # compile + run VB
//!   vybec calculator.js     # compile + run JS
//!   vybec --dump hello.vb   # compile + dump bytecode (no run)

use std::env;
use std::fs;
use std::path::Path;
use std::rc::Rc;
use std::cell::RefCell;
use vybe_bytecode::VM;

fn main() {
    let args: Vec<String> = env::args().collect();
    let mut dump = false;
    let mut file_arg = None;

    for arg in &args[1..] {
        if arg == "--dump" || arg == "-d" {
            dump = true;
        } else if file_arg.is_none() {
            file_arg = Some(arg.clone());
        }
    }

    let file_path = match file_arg {
        Some(f) => f,
        None => {
            eprintln!("Usage: vybec [--dump] <file.vb|file.js>");
            std::process::exit(1);
        }
    };

    let path = Path::new(&file_path);
    if !path.exists() {
        eprintln!("Error: file not found: {}", path.display());
        std::process::exit(1);
    }

    let ext = path.extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();

    match ext.as_str() {
        "vb" => run_vb(path, dump),
        "js" => run_js(path, dump),
        _ => {
            eprintln!("Error: unsupported file type '.{}'. Expected .vb or .js", ext);
            std::process::exit(1);
        }
    }
}

fn run_vb(path: &Path, dump: bool) {
    let source = read_file(path);
    let program = match vybe_parser_basic::parse_program(&source) {
        Ok(p) => p,
        Err(e) => { eprintln!("Parse error: {e:?}"); std::process::exit(1); }
    };

    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    vybe_host::register_all_with_gui(&mut vm, queue.clone());
    vybe_host::setup_namespaces(&mut vm);

    let chunks = match vybe_compiler_vb::Compiler::new().compile(&program) {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {e}"); std::process::exit(1); }
    };

    if dump { dump_chunks(&chunks); return; }

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    // Launch GUI if RunApplication was called, otherwise just print console output
    vybe_ui::launch_vm_form(vm, queue);
}

fn run_js(path: &Path, dump: bool) {
    let source = read_file(path);
    let program = match vybe_parser_js::parse(&source) {
        Ok(p) => p,
        Err(e) => { eprintln!("Parse error: {e}"); std::process::exit(1); }
    };

    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    vybe_host::register_all_with_gui(&mut vm, queue.clone());
    vybe_compiler_js::register_js_coercion(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    let chunks = match vybe_compiler_js::Compiler::new().compile(&program) {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {e}"); std::process::exit(1); }
    };

    if dump { dump_chunks(&chunks); return; }

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    vybe_ui::launch_vm_form(vm, queue);
}

fn read_file(path: &Path) -> String {
    match fs::read_to_string(path) {
        Ok(s) => s,
        Err(e) => { eprintln!("Error reading {}: {e}", path.display()); std::process::exit(1); }
    }
}

fn dump_chunks(chunks: &[vybe_bytecode::Chunk]) {
    for (i, chunk) in chunks.iter().enumerate() {
        println!("=== Chunk {} ({}) ===", i, chunk.name);
        println!("  arity: {}, locals: {}", chunk.arity, chunk.local_count);
        if !chunk.imports.is_empty() {
            println!("  imports:");
            for (j, imp) in chunk.imports.iter().enumerate() {
                println!("    [{}] {}:{}", j, imp.module, imp.name);
            }
        }
        println!("  bytecode: {} bytes", chunk.code.len());
        println!();
    }
}
