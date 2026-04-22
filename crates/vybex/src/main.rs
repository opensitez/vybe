//! vybex — Universal compiler.
//!
//! Usage: vybex [flags] <file|project>
//!
//! Flags:
//!   --dump, -d        Disassemble bytecode and exit (no run)
//!   --emit-wasm, -w   Compile to .wasm binary and exit
//!   --sandbox, -s     Restricted mode (no filesystem/network/database)
//!   --portable, -p    Minimal WASI runtime only (no Vybe host optimizations)
//!   --trace, -t       Enable bytecode trace output
//!   --chunk <name>    Limit --dump/--trace output to a specific chunk
//!
//! Supports single source files (detected by extension), project files
//! (.vybe, .vbproj), and .wasm binaries. Language is determined automatically.

use std::path::Path;
use std::sync::{Arc, Mutex};
use vybe_bytecode::chunk::Chunk;
use vybe_bytecode::VM;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let mut dump = false;
    let mut emit_wasm = false;
    let mut sandbox = false;
    let mut portable = false;
    let mut trace = false;
    let mut chunk_filter = None;
    let mut file_arg = None;

    let mut iter = args[1..].iter();
    while let Some(arg) = iter.next() {
        match arg.as_str() {
            "--dump" | "-d" => dump = true,
            "--emit-wasm" | "-w" => emit_wasm = true,
            "--sandbox" | "-s" => sandbox = true,
            "--portable" | "-p" => portable = true,
            "--trace" | "-t" => trace = true,
            "--chunk" => {
                let Some(name) = iter.next() else {
                    eprintln!("Missing value for --chunk");
                    std::process::exit(1);
                };
                chunk_filter = Some(name.clone());
            }
            "--help" | "-h" => {
                print_usage();
                return;
            }
            _ if file_arg.is_none() && !arg.starts_with('-') => file_arg = Some(arg.clone()),
            _ => {
                eprintln!("Unknown flag: {arg}");
                std::process::exit(1);
            }
        }
    }

    let file_path = match file_arg {
        Some(f) => f,
        None => {
            print_usage();
            std::process::exit(1);
        }
    };

    let path = Path::new(&file_path);

    // ── Handle .wasm binaries directly ──────────────────────────────────────
    if path.extension().and_then(|e| e.to_str()) == Some("wasm") {
        run_wasm(path, dump, trace, chunk_filter.as_deref());
        return;
    }

    // ── Load project/source ─────────────────────────────────────────────────
    let bundle = match vybex::projects::load(path) {
        Ok(b) => b,
        Err(e) => { eprintln!("{e}"); std::process::exit(1); }
    };

    eprintln!("[vybex] Project '{}', sources={}, entry={:?}",
        bundle.name,
        bundle.sources.len(),
        bundle.entry_point);
    for s in &bundle.sources {
        eprintln!("  → {} ({} bytes)", s.path.display(), s.code.len());
    }

    // ── Compile ─────────────────────────────────────────────────────────────
    let chunks = match bundle.compile() {
        Ok(c) => c,
        Err(e) => { eprintln!("Compile error: {e}"); std::process::exit(1); }
    };

    // ── --dump: disassemble and exit ────────────────────────────────────────
    if dump {
        for chunk in filter_chunks(&chunks, chunk_filter.as_deref()) {
            println!("{}", vybe_bytecode::debug::disassemble(chunk));
        }
        return;
    }

    // ── --emit-wasm: write .wasm binary and exit ────────────────────────────
    if emit_wasm {
        let wasm_bytes = vybe_bytecode::wasm::write_wasm(&chunks);
        let out_path = path.with_extension("wasm");
        std::fs::write(&out_path, &wasm_bytes).unwrap();
        eprintln!("Wrote {} bytes to {}", wasm_bytes.len(), out_path.display());
        return;
    }

    // ── Set up VM ───────────────────────────────────────────────────────────
    let mut vm = VM::new();

    let gui = if sandbox {
        eprintln!("[sandbox] Restricted mode: no filesystem, network, or database access");
        vybe_host::register_with_capabilities_and_gui(
            &mut vm, &vybe_host::Capabilities::safe(),
        )
    } else if portable {
        eprintln!("[portable] Running with WASM stdlib only — no Vybe host optimizations");
        vm.register_host_fn("wasi:cli", "log", Box::new(|_ctx: &mut vybe_bytecode::HostContext, args: &[vybe_bytecode::Value]| {
            for a in args { print!("{}", a); }
            println!();
            vybe_bytecode::Value::Null
        }));
        vm.register_host_fn("wasi:cli", "readLine", Box::new(|_ctx: &mut vybe_bytecode::HostContext, _| {
            let mut line = String::new();
            std::io::stdin().read_line(&mut line).ok();
            vybe_bytecode::Value::String(std::sync::Arc::from(line.trim()))
        }));
        Arc::new(Mutex::new(vybe_host::gui_state::GuiState::new()))
    } else {
        vybe_host::register_all_with_gui(&mut vm)
    };

    if !portable {
        vybe_host::setup_namespaces(&mut vm);
    }

    if trace {
        vm.set_trace(true);
        vm.set_trace_chunk_filter(chunk_filter.clone());
    }

    // ── Register WASM function names as globals ─────────────────────────────
    // When a .vybe project includes .wasm files, their named functions are
    // appended after the compiled source chunks. Register them so the source
    // code can call them by name (e.g. `add(3, 4)` in VB).
    for (idx, chunk) in chunks.iter().enumerate() {
        if !chunk.name.is_empty()
            && chunk.name != "<script>"
            && chunk.name != "<bootstrap>"
            && !chunk.name.starts_with("__stdlib_")
        {
            use std::sync::{Arc as StdArc, Mutex as StdMutex};
            let func = vybe_bytecode::value::Function {
                name: Some(chunk.name.clone()),
                arity: chunk.arity,
                chunk_index: idx,
                upvalues: vec![],
            };
            let mut obj = vybe_bytecode::value::Object::new();
            obj.kind = vybe_bytecode::value::ObjectKind::Function(func);
            let val = vybe_bytecode::Value::Object(StdArc::new(StdMutex::new(obj)));
            vm.globals.insert(chunk.name.to_lowercase(), val);
        }
    }

    // ── Run ─────────────────────────────────────────────────────────────────
    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    if gui.lock().unwrap().should_run {
        vybex::gui_launch::launch_gui(vm, gui);
    }
}

fn run_wasm(path: &Path, dump: bool, trace: bool, chunk_filter: Option<&str>) {
    let data = match std::fs::read(path) {
        Ok(d) => d,
        Err(e) => { eprintln!("Error reading {}: {e}", path.display()); std::process::exit(1); }
    };
    eprintln!("Loading WASM: {} ({} bytes)", path.display(), data.len());

    let chunks = match vybe_bytecode::wasm::read_wasm(&data) {
        Ok(c) => c,
        Err(e) => { eprintln!("WASM error: {e}"); std::process::exit(1); }
    };
    eprintln!("Loaded {} chunks", chunks.len());

    if dump {
        for chunk in filter_chunks(&chunks, chunk_filter) {
            println!("{}", vybe_bytecode::debug::disassemble(chunk));
        }
        return;
    }

    let mut vm = VM::new();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    if trace {
        vm.set_trace(true);
        vm.set_trace_chunk_filter(chunk_filter.map(|s| s.to_string()));
    }

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => { eprintln!("Runtime error: {e}"); std::process::exit(1); }
    }

    if gui.lock().unwrap().should_run {
        vybex::gui_launch::launch_gui(vm, gui);
    }
}

fn filter_chunks<'a>(chunks: &'a [Chunk], chunk_filter: Option<&str>) -> Vec<&'a Chunk> {
    match chunk_filter {
        Some(filter) => {
            if let Ok(index) = filter.parse::<usize>() {
                chunks.get(index).into_iter().collect()
            } else {
                chunks.iter().filter(|chunk| chunk.name == filter).collect()
            }
        }
        None => chunks.iter().collect(),
    }
}

fn print_usage() {
    let exts = vybex::projects::supported_extensions();
    let ext_list: Vec<String> = exts.iter().map(|e| format!(".{e}")).collect();
    eprintln!("vybex — Universal compiler");
    eprintln!();
    eprintln!("Usage: vybex [flags] <file>");
    eprintln!();
    eprintln!("Flags:");
    eprintln!("  -d, --dump        Disassemble bytecode (no run)");
    eprintln!("  -w, --emit-wasm   Compile to .wasm binary");
    eprintln!("  -s, --sandbox     Restricted mode (safe capabilities only)");
    eprintln!("  -p, --portable    Minimal WASI runtime (no Vybe host)");
    eprintln!("  -t, --trace       Enable bytecode trace output");
    eprintln!("      --chunk NAME  Limit --dump/--trace output to a chunk name or index");
    eprintln!("  -h, --help        Show this help");
    eprintln!();
    eprintln!("Supported: {}", ext_list.join(", "));
}
