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
    let mut emit_wasm = false;
    let mut sandbox = false;
    let mut file_arg = None;

    for arg in &args[1..] {
        match arg.as_str() {
            "--dump" | "-d" => dump = true,
            "--emit-wasm" | "-w" => emit_wasm = true,
            "--sandbox" | "-s" => sandbox = true,
            _ if file_arg.is_none() => file_arg = Some(arg.clone()),
            _ => {}
        }
    }

    let file_path = match file_arg {
        Some(f) => f,
        None => {
            eprintln!("Usage: vybec [--dump] [--sandbox] <file.vb|file.js|file.dart|file.vbp|file.vbproj>");
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
        "vb" => run_vb(path, dump, emit_wasm, sandbox),
        "js" => run_js(path, dump, emit_wasm, sandbox),
        "dart" => run_dart(path, dump, emit_wasm, sandbox),
        "wasm" => run_wasm(path),
        "vybe" => run_project(path, dump),
        "vbp" | "vbproj" => vybe_cli::runner::run(path, &[]),
        "cs" => vybe_cli::runner::run(path, &[]),
        _ => {
            // Check if directory has a .vybe project file
            let vybe_path = path.join("project.vybe");
            if vybe_path.exists() {
                run_project(&vybe_path, dump);
            } else {
                eprintln!("Error: unsupported file type '.{}'. Expected .vb, .js, .dart, .cs, .vbp, .vbproj, or .vybe", ext);
                std::process::exit(1);
            }
        }
    }
}

fn run_project(path: &Path, dump: bool) {
    let toml_content = read_file(path);
    let config = match vybe_bytecode::ProjectConfig::parse(&toml_content) {
        Ok(c) => c,
        Err(e) => { eprintln!("Project error: {e}"); std::process::exit(1); }
    };

    let project_dir = path.parent().unwrap_or(Path::new("."));

    // Set up VM + linker
    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    if config.host.gui {
        vybe_host::register_all_with_gui(&mut vm, queue.clone());
    } else {
        vybe_host::register_all(&mut vm);
    }
    vybe_host::setup_namespaces(&mut vm);

    // Create linker and register host exports
    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);

    // Compile and run each source file.
    // Library files run first (set up globals/exports), entry file runs last.
    let mut file_chunks: Vec<(String, Vec<vybe_bytecode::Chunk>)> = Vec::new();

    for file in &config.files {
        let file_path = project_dir.join(file);
        let ext = file_path.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();

        let chunks = match ext.as_str() {
            "wasm" => {
                let data = match std::fs::read(&file_path) {
                    Ok(d) => d,
                    Err(e) => { eprintln!("Error reading {}: {e}", file); std::process::exit(1); }
                };
                match vybe_bytecode::wasm::read_wasm(&data) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("WASM error in {}: {e}", file); std::process::exit(1); }
                }
            }
            "vb" => {
                let source = read_file(&file_path);
                let program = match vybe_parser_basic::parse_program(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e:?}", file); std::process::exit(1); }
                };
                match vybe_compiler_vb::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                }
            }
            "js" => {
                let source = read_file(&file_path);
                vybe_compiler_js::register_js_coercion(&mut vm);
                let program = match vybe_parser_js::parse(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e}", file); std::process::exit(1); }
                };
                match vybe_compiler_js::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                }
            }
            "dart" => {
                let source = read_file(&file_path);
                let program = match vybe_parser_dart::parse(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e}", file); std::process::exit(1); }
                };
                match vybe_compiler_dart::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                }
            }
            _ => { eprintln!("Unknown file type: {}", file); continue; }
        };
        file_chunks.push((file.clone(), chunks));
    }

    // Run library files first, entry file last
    for (file, chunks) in &file_chunks {
        if *file != config.entry {
            let ext = std::path::Path::new(file).extension()
                .and_then(|e| e.to_str()).unwrap_or("").to_lowercase();

            if ext == "wasm" {
                // WASM library: register exported functions as VM globals
                // so VB/JS can call them by name
                let chunk_offset = vm.chunks.len();
                vm.chunks.extend(chunks.clone());
                for chunk in chunks {
                    if chunk.name != "<script>" && !chunk.name.starts_with("func_") {
                        // Create a closure value for this function
                        let func_idx = chunk_offset + chunks.iter().position(|c| c.name == chunk.name).unwrap_or(0);
                        let func = vybe_bytecode::value::Function {
                            name: Some(chunk.name.clone()),
                            arity: chunk.arity,
                            chunk_index: func_idx,
                            upvalues: Vec::new(),
                        };
                        let obj = vybe_bytecode::value::Object {
                            properties: std::collections::HashMap::new(),
                            kind: vybe_bytecode::value::ObjectKind::Function(func),
                            type_id: 0, fields: Vec::new(),
                        };
                        let val = vybe_bytecode::Value::Object(
                            Rc::new(RefCell::new(obj))
                        );
                        vm.globals.insert(chunk.name.to_lowercase(), val);
                        eprintln!("  Registered WASM function: {}", chunk.name);
                    }
                }
            } else {
                if let Err(e) = vm.run(chunks.clone()) {
                    eprintln!("Runtime error in {}: {e}", file);
                }
            }
        }
    }

    // Collect entry file chunks + append any WASM library chunks
    // so chunk_index references remain valid
    let mut all_chunks: Vec<vybe_bytecode::Chunk> = file_chunks.into_iter()
        .find(|(f, _)| *f == config.entry)
        .map(|(_, c)| c)
        .unwrap_or_default();

    // Append WASM library chunks that were registered as globals
    // Their chunk_index was set relative to vm.chunks, so we need
    // to adjust or just ensure vm.chunks includes them after run()
    // Actually: append them to all_chunks and update the global Function's chunk_index
    let wasm_chunk_offset = all_chunks.len();
    let wasm_chunks: Vec<vybe_bytecode::Chunk> = vm.chunks.drain(..).collect();
    all_chunks.extend(wasm_chunks);

    // Update globals that reference WASM chunks
    let globals_to_fix: Vec<String> = vm.globals.keys().cloned().collect();
    for key in globals_to_fix {
        if let Some(vybe_bytecode::Value::Object(obj)) = vm.globals.get(&key) {
            let mut o = obj.borrow_mut();
            if let vybe_bytecode::value::ObjectKind::Function(ref mut func) = o.kind {
                if func.chunk_index > 0 && func.chunk_index < wasm_chunk_offset + 100 {
                    func.chunk_index += wasm_chunk_offset;
                }
            }
        }
    }

    if dump {
        for (i, chunk) in all_chunks.iter().enumerate() {
            println!("=== Chunk {} ({}) ===", i, chunk.name);
            println!("  arity: {}, locals: {}", chunk.arity, chunk.local_count);
            println!("  bytecode: {} bytes", chunk.code.len());
            println!();
        }
        return;
    }

    if all_chunks.is_empty() {
        eprintln!("No entry file compiled");
        std::process::exit(1);
    }

    // Run entry file
    match vm.run(all_chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    vybe_cli::runner::launch_vm_form(vm, queue, None);
}

fn run_vb(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool) {
    let source = read_file(path);
    let program = match vybe_parser_basic::parse_program(&source) {
        Ok(p) => p,
        Err(e) => { eprintln!("Parse error: {e:?}"); std::process::exit(1); }
    };

    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    if sandbox {
        eprintln!("[sandbox] Restricted mode: no filesystem, network, or database access");
        vybe_host::register_with_capabilities_and_gui(
            &mut vm, &vybe_host::Capabilities::safe(), queue.clone(),
        );
    } else {
        vybe_host::register_all_with_gui(&mut vm, queue.clone());
    }
    vybe_host::setup_namespaces(&mut vm);

    let chunks = match vybe_compiler_vb::Compiler::new().compile(&program) {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {e}"); std::process::exit(1); }
    };

    if dump { dump_chunks(&chunks); return; }
    if emit_wasm {
        let wasm_bytes = vybe_bytecode::wasm::write_wasm(&chunks);
        let out_path = path.with_extension("wasm");
        std::fs::write(&out_path, &wasm_bytes).unwrap();
        eprintln!("Wrote {} bytes to {}", wasm_bytes.len(), out_path.display());
        return;
    }

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    vybe_cli::runner::launch_vm_form(vm, queue, None);
}

fn run_js(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool) {
    let source = read_file(path);
    let program = match vybe_parser_js::parse(&source) {
        Ok(p) => p,
        Err(e) => { eprintln!("Parse error: {e}"); std::process::exit(1); }
    };

    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    if sandbox {
        eprintln!("[sandbox] Restricted mode: no filesystem, network, or database access");
        vybe_host::register_with_capabilities_and_gui(
            &mut vm, &vybe_host::Capabilities::safe(), queue.clone(),
        );
    } else {
        vybe_host::register_all_with_gui(&mut vm, queue.clone());
    }
    vybe_compiler_js::register_js_coercion(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    let chunks = match vybe_compiler_js::Compiler::new().compile(&program) {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {e}"); std::process::exit(1); }
    };

    if dump { dump_chunks(&chunks); return; }
    if emit_wasm {
        let wasm_bytes = vybe_bytecode::wasm::write_wasm(&chunks);
        let out_path = path.with_extension("wasm");
        std::fs::write(&out_path, &wasm_bytes).unwrap();
        eprintln!("Wrote {} bytes to {}", wasm_bytes.len(), out_path.display());
        return;
    }

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    vybe_cli::runner::launch_vm_form(vm, queue, None);
}

fn run_dart(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool) {
    let source = read_file(path);
    let program = match vybe_parser_dart::parse(&source) {
        Ok(p) => p,
        Err(e) => { eprintln!("Parse error: {e}"); std::process::exit(1); }
    };

    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    if sandbox {
        eprintln!("[sandbox] Restricted mode: no filesystem, network, or database access");
        vybe_host::register_with_capabilities_and_gui(
            &mut vm, &vybe_host::Capabilities::safe(), queue.clone(),
        );
    } else {
        vybe_host::register_all_with_gui(&mut vm, queue.clone());
    }
    vybe_host::setup_namespaces(&mut vm);

    let chunks = match vybe_compiler_dart::Compiler::new().compile(&program) {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {e}"); std::process::exit(1); }
    };

    if dump { dump_chunks(&chunks); return; }
    if emit_wasm {
        let wasm_bytes = vybe_bytecode::wasm::write_wasm(&chunks);
        let out_path = path.with_extension("wasm");
        std::fs::write(&out_path, &wasm_bytes).unwrap();
        eprintln!("Wrote {} bytes to {}", wasm_bytes.len(), out_path.display());
        return;
    }

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    vybe_cli::runner::launch_vm_form(vm, queue, None);
}

fn run_wasm(path: &Path) {
    let data = match std::fs::read(path) {
        Ok(d) => d,
        Err(e) => { eprintln!("Error reading {}: {e}", path.display()); std::process::exit(1); }
    };
    eprintln!("Loading WASM: {} ({} bytes)", path.display(), data.len());
    let chunks = match vybe_bytecode::wasm::read_wasm(&data) {
        Ok(c) => c,
        Err(e) => { eprintln!("WASM error: {e}"); std::process::exit(1); }
    };
    eprintln!("Loaded {} chunks:", chunks.len());
    for (i, chunk) in chunks.iter().enumerate() {
        eprintln!("  [{}] {} (arity={}, locals={}, code={}B)", i, chunk.name, chunk.arity, chunk.local_count, chunk.code.len());
    }

    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    vybe_host::register_all_with_gui(&mut vm, queue.clone());
    vybe_host::setup_namespaces(&mut vm);

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }
    vybe_cli::runner::launch_vm_form(vm, queue, None);
}

fn read_file(path: &Path) -> String {
    match fs::read_to_string(path) {
        Ok(s) => s,
        Err(e) => { eprintln!("Error reading {}: {e}", path.display()); std::process::exit(1); }
    }
}

fn dump_chunks(chunks: &[vybe_bytecode::Chunk]) {
    for (i, chunk) in chunks.iter().enumerate() {
        println!("{}", vybe_bytecode::debug::disassemble(chunk));
    }
}
