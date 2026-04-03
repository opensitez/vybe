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
    let mut portable = false;
    let mut file_arg = None;

    for arg in &args[1..] {
        match arg.as_str() {
            "--dump" | "-d" => dump = true,
            "--emit-wasm" | "-w" => emit_wasm = true,
            "--sandbox" | "-s" => sandbox = true,
            "--portable" | "-p" => portable = true,
            _ if file_arg.is_none() => file_arg = Some(arg.clone()),
            _ => {}
        }
    }

    let file_path = match file_arg {
        Some(f) => f,
        None => {
            eprintln!("Usage: vybec [--dump] [--sandbox] [--portable] <file.vb|file.js|file.dart|file.py|file.php|file.rb>");
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
        "vb" => run_vb(path, dump, emit_wasm, sandbox, portable),
        "js" => run_js(path, dump, emit_wasm, sandbox, portable),
        "dart" => run_dart(path, dump, emit_wasm, sandbox, portable),
        "py" | "py3" => run_python(path, dump, emit_wasm, sandbox, portable),
        "php" => run_php(path, dump, emit_wasm, sandbox, portable),
        "rb" => run_ruby(path, dump, emit_wasm, sandbox, portable),
        "cob" | "cbl" | "cobol" => run_cobol(path, dump, emit_wasm, sandbox, portable),
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
                eprintln!("Error: unsupported file type '.{}'. Expected .vb, .js, .dart, .py, .php, .rb, .cs, .vbp, .vbproj, or .vybe", ext);
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

    // ── Phase 1: Set up VM with host functions ─────────────────
    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    if config.host.gui {
        vybe_host::register_all_with_gui(&mut vm, queue.clone());
    } else {
        vybe_host::register_all(&mut vm);
    }
    vybe_host::setup_namespaces(&mut vm);

    // ── Phase 2: Compile ALL files into Components ─────────────
    // Each source file becomes a Component with exports/imports.
    // No vm.run() calls here — just compilation.
    let mut components: Vec<vybe_bytecode::Component> = Vec::new();
    let mut entry_idx: Option<usize> = None;

    for file in &config.files {
        let file_path = project_dir.join(file);
        let ext = file_path.extension().and_then(|e| e.to_str()).unwrap_or("").to_lowercase();
        let module_name = file_path.file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("module")
            .to_string();

        let (language, chunks) = match ext.as_str() {
            "wasm" => {
                let data = match std::fs::read(&file_path) {
                    Ok(d) => d,
                    Err(e) => { eprintln!("Error reading {}: {e}", file); std::process::exit(1); }
                };
                let c = match vybe_bytecode::wasm::read_wasm(&data) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("WASM error in {}: {e}", file); std::process::exit(1); }
                };
                (vybe_bytecode::Language::Wasm, c)
            }
            "vb" => {
                let source = read_file(&file_path);
                let program = match vybe_parser_basic::parse_program(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e:?}", file); std::process::exit(1); }
                };
                let c = match vybe_compiler_vb::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                };
                (vybe_bytecode::Language::VB, c)
            }
            "js" => {
                vybe_compiler_js::register_js_coercion(&mut vm);
                let source = read_file(&file_path);
                let program = match vybe_parser_js::parse(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e}", file); std::process::exit(1); }
                };
                let c = match vybe_compiler_js::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                };
                (vybe_bytecode::Language::JS, c)
            }
            "cs" => {
                let source = read_file(&file_path);
                let program = match vybe_parser_csharp::parse(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e}", file); std::process::exit(1); }
                };
                let c = match vybe_compiler_csharp::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                };
                (vybe_bytecode::Language::CSharp, c)
            }
            "dart" => {
                let source = read_file(&file_path);
                let program = match vybe_parser_dart::parse(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e}", file); std::process::exit(1); }
                };
                let c = match vybe_compiler_dart::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                };
                (vybe_bytecode::Language::Dart, c)
            }
            "py" | "py3" => {
                let source = read_file(&file_path);
                let module = match vybe_parser_python::parse(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e}", file); std::process::exit(1); }
                };
                let c = match vybe_compiler_python::Compiler::new().compile(&module) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                };
                (vybe_bytecode::Language::Python, c)
            }
            "php" => {
                let source = read_file(&file_path);
                let program = match vybe_parser_php::parse(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e}", file); std::process::exit(1); }
                };
                let c = match vybe_compiler_php::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                };
                (vybe_bytecode::Language::Php, c)
            }
            "rb" => {
                let source = read_file(&file_path);
                let program = match vybe_parser_ruby::parse(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e}", file); std::process::exit(1); }
                };
                let c = match vybe_compiler_ruby::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                };
                (vybe_bytecode::Language::Ruby, c)
            }
            "cob" | "cbl" | "cobol" => {
                let source = read_file(&file_path);
                let program = match vybe_parser_cobol::parse(&source) {
                    Ok(p) => p,
                    Err(e) => { eprintln!("Parse error in {}: {e}", file); std::process::exit(1); }
                };
                let c = match vybe_compiler_cobol::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => { eprintln!("Compile error in {}: {e}", file); std::process::exit(1); }
                };
                (vybe_bytecode::Language::Cobol, c)
            }
            _ => { eprintln!("Unknown file type: {}", file); continue; }
        };

        // Build Component with exports
        let component = vybe_compiler_common::components::build_component(
            &module_name, language, chunks,
        );

        if *file == config.entry {
            entry_idx = Some(components.len());
        }
        components.push(component);
    }

    // ── Phase 3: Link all components via Linker ────────────────
    // The Linker merges chunks, adjusts ref_func indices,
    // resolves imports/exports, and applies CLS case resolution.
    let mut linker = vybe_bytecode::Linker::new();
    linker.register_host_from_vm(&vm);
    for comp in &components {
        linker.add_component(comp.clone());
    }

    let link_result = match linker.link() {
        Ok(r) => r,
        Err(e) => { eprintln!("Link error: {e}"); std::process::exit(1); }
    };

    if dump {
        for (i, chunk) in link_result.chunks.iter().enumerate() {
            println!("=== Chunk {} ({}) ===", i, chunk.name);
            println!("  arity: {}, locals: {}", chunk.arity, chunk.local_count);
            println!("  bytecode: {} bytes", chunk.code.len());
            println!();
        }
        return;
    }

    // ── Phase 4: Build bootstrap + run ───────────────────────────
    // The Linker merged all chunks. Each component's script chunk (chunk 0)
    // is at its component_offset. We create a bootstrap chunk that calls
    // each script chunk in order: libraries first, entry last.
    //
    // Bootstrap bytecode:
    //   ref_func comp1_script_idx; call_ref 0; drop;
    //   ref_func comp2_script_idx; call_ref 0; drop;
    //   ...
    //   ref_func entry_script_idx; call_ref 0; drop;
    //   halt;
    let mut bootstrap = vybe_bytecode::Chunk::new("<bootstrap>");
    let line = 0u32;

    // Call library script chunks first (non-entry)
    for (i, _comp) in components.iter().enumerate() {
        if Some(i) == entry_idx { continue; }
        let script_idx = link_result.component_offsets[i];
        // ref_func + call_ref 0 + drop
        bootstrap.emit_op_u16(vybe_bytecode::Op::ref_func, script_idx as u16, line);
        bootstrap.emit(0, line); // 0 upvalues
        bootstrap.emit_op_u8(vybe_bytecode::Op::call_ref, 0, line);
        bootstrap.emit_op(vybe_bytecode::Op::drop, line);
    }

    // Call entry script chunk last
    if let Some(ei) = entry_idx {
        let script_idx = link_result.component_offsets[ei];
        bootstrap.emit_op_u16(vybe_bytecode::Op::ref_func, script_idx as u16, line);
        bootstrap.emit(0, line);
        bootstrap.emit_op_u8(vybe_bytecode::Op::call_ref, 0, line);
        bootstrap.emit_op(vybe_bytecode::Op::drop, line);
    }

    bootstrap.emit_op(vybe_bytecode::Op::halt, line);
    bootstrap.local_count = 16;

    // Prepend bootstrap chunk — all other chunk indices shift by 1
    // (The Linker already adjusted ref_func indices relative to its own offsets,
    //  but now we're prepending a chunk, so we need to shift everything by 1)
    let mut all_chunks = vec![bootstrap];
    for mut chunk in link_result.chunks {
        // Adjust ref_func indices in each chunk: +1 for the prepended bootstrap
        let code = &mut chunk.code;
        let mut ip = 0;
        while ip < code.len() {
            let op_byte = code[ip];
            if let Some(op) = vybe_bytecode::Op::from_byte(op_byte) {
                match op {
                    vybe_bytecode::Op::ref_func => {
                        if ip + 2 < code.len() {
                            let old_idx = ((code[ip + 1] as u16) << 8) | (code[ip + 2] as u16);
                            let new_idx = old_idx + 1; // shift by 1 for bootstrap
                            code[ip + 1] = (new_idx >> 8) as u8;
                            code[ip + 2] = (new_idx & 0xff) as u8;
                        }
                        ip += 3 + 1;
                        if ip - 1 < code.len() {
                            let uv_count = code[ip - 1] as usize;
                            ip += uv_count * 2;
                        }
                        continue;
                    }
                    vybe_bytecode::Op::call_import => { ip += 4; continue; }
                    vybe_bytecode::Op::call | vybe_bytecode::Op::call_ref => { ip += 2; continue; }
                    vybe_bytecode::Op::r#const | vybe_bytecode::Op::local_get | vybe_bytecode::Op::local_set
                    | vybe_bytecode::Op::global_get | vybe_bytecode::Op::global_set
                    | vybe_bytecode::Op::struct_get | vybe_bytecode::Op::struct_set
                    | vybe_bytecode::Op::array_new
                    | vybe_bytecode::Op::br | vybe_bytecode::Op::br_if_true | vybe_bytecode::Op::br_if_false
                    | vybe_bytecode::Op::r#loop => { ip += 3; continue; }
                    _ => { ip += 1; continue; }
                }
            } else {
                ip += 1;
            }
        }
        all_chunks.push(chunk);
    }

    // Also adjust the bootstrap's ref_func indices by +1 (they pointed to Linker offsets)
    {
        let code = &mut all_chunks[0].code;
        let mut ip = 0;
        while ip < code.len() {
            let op_byte = code[ip];
            if let Some(op) = vybe_bytecode::Op::from_byte(op_byte) {
                match op {
                    vybe_bytecode::Op::ref_func => {
                        if ip + 2 < code.len() {
                            let old_idx = ((code[ip + 1] as u16) << 8) | (code[ip + 2] as u16);
                            let new_idx = old_idx + 1;
                            code[ip + 1] = (new_idx >> 8) as u8;
                            code[ip + 2] = (new_idx & 0xff) as u8;
                        }
                        ip += 3 + 1;
                        if ip - 1 < code.len() {
                            let uv_count = code[ip - 1] as usize;
                            ip += uv_count * 2;
                        }
                        continue;
                    }
                    vybe_bytecode::Op::call_ref => { ip += 2; continue; }
                    _ => { ip += 1; continue; }
                }
            } else { ip += 1; }
        }
    }

    match vm.run(all_chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    vybe_cli::runner::launch_vm_form(vm, queue, None);
}

fn run_vb(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool, _portable: bool) {
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

fn run_js(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool, _portable: bool) {
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

fn run_dart(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool, _portable: bool) {
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

fn run_python(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool, portable: bool) {
    let source = read_file(path);
    let module = match vybe_parser_python::parse(&source) {
        Ok(p) => p,
        Err(e) => { eprintln!("Parse error: {e}"); std::process::exit(1); }
    };

    let mut vm = VM::new();
    // Stdlib is bundled in the compiled chunks (via global_inits + RefFunc).
    // On Vybe, register_all overwrites __vybe_* globals with fast host fns.
    // On --portable, only minimal WASI imports are registered.

    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    if !portable {
        if sandbox {
            eprintln!("[sandbox] Restricted mode: no filesystem, network, or database access");
            vybe_host::register_with_capabilities_and_gui(
                &mut vm, &vybe_host::Capabilities::safe(), queue.clone(),
            );
        } else {
            vybe_host::register_all_with_gui(&mut vm, queue.clone());
        }
        vybe_host::setup_namespaces(&mut vm);
    } else {
        eprintln!("[portable] Running with WASM stdlib only — no Vybe host optimizations");
        // Register minimal WASI imports for I/O
        vm.register_host_fn("wasi:cli", "log", Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[vybe_bytecode::Value]| {
            for a in args { print!("{}", a); }
            println!();
            vybe_bytecode::Value::Null
        }));
        vm.register_host_fn("wasi:cli", "readLine", Box::new(|_ctx: &mut vybe_bytecode::HostContext, _| {
            let mut line = String::new();
            std::io::stdin().read_line(&mut line).ok();
            vybe_bytecode::Value::String(std::rc::Rc::from(line.trim()))
        }));
    }

    let chunks = match vybe_compiler_python::Compiler::new().compile(&module) {
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

fn run_php(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool, _portable: bool) {
    let source = read_file(path);
    let program = match vybe_parser_php::parse(&source) {
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

    let chunks = match vybe_compiler_php::Compiler::new().compile(&program) {
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

fn run_cobol(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool, _portable: bool) {
    let source = read_file(path);
    let program = match vybe_parser_cobol::parse(&source) {
        Ok(p) => p,
        Err(e) => { eprintln!("Parse error: {e}"); std::process::exit(1); }
    };

    let mut vm = VM::new();
    let queue = Rc::new(RefCell::new(vybe_host::SideEffectQueue::new()));
    if sandbox {
        eprintln!("[sandbox] Restricted mode");
        vybe_host::register_with_capabilities_and_gui(
            &mut vm, &vybe_host::Capabilities::safe(), queue.clone(),
        );
    } else {
        vybe_host::register_all_with_gui(&mut vm, queue.clone());
    }
    vybe_host::setup_namespaces(&mut vm);

    let chunks = match vybe_compiler_cobol::Compiler::new().compile(&program) {
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

fn run_ruby(path: &Path, dump: bool, emit_wasm: bool, sandbox: bool, _portable: bool) {
    let source = read_file(path);
    let program = match vybe_parser_ruby::parse(&source) {
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

    let chunks = match vybe_compiler_ruby::Compiler::new().compile(&program) {
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
    for (_i, chunk) in chunks.iter().enumerate() {
        println!("{}", vybe_bytecode::debug::disassemble(chunk));
    }
}
