//! vybex — Universal compiler.
//!
//! Usage: vybex [flags] <file|project>
//!
//! Flags:
//!   --dump, -d        Disassemble bytecode and exit (no run)
//!   --dump-ast        Parse and print the prepared common AST, then exit
//!   --emit-wasm, -w   Compile to .wasm binary and exit
//!   --eval CODE       Compile and run source from a string
//!   --lang NAME       Language to use with --eval
//!   --virtual-path P  Virtual source path to use with --eval
//!   --sandbox, -s     Restricted mode (no filesystem/network/database)
//!   --portable, -p    Minimal WASI runtime only (no Vybe host optimizations)
//!   --trace, -t       Enable bytecode trace output
//!   --chunk <name>    Limit --dump/--trace output to a specific chunk
//!
//! Supports single source files (detected by extension), project files
//! (.vybe, .vbproj, .csproj, .pyproj/.ipyproj), and .wasm binaries.
//! Language is determined automatically.

use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use vybe_bytecode::VM;
use vybe_bytecode::chunk::Chunk;
use vybe_compiler::ast::{ExprKind, Literal, Module, StmtKind};

#[derive(Default)]
struct AstSummary {
    top_level_statements: usize,
    top_level_functions: usize,
    top_level_classes: usize,
    top_level_echoes: usize,
    string_bytes: usize,
    inline_html_echoes: usize,
}

fn summarize_module(module: &Module) -> AstSummary {
    let mut summary = AstSummary::default();
    for stmt in &module.body {
        summary.top_level_statements += 1;
        match &stmt.kind {
            StmtKind::FunctionDecl { .. } => summary.top_level_functions += 1,
            StmtKind::ClassDecl { .. }
            | StmtKind::StructDecl { .. }
            | StmtKind::ModuleDecl { .. }
            | StmtKind::InterfaceDecl { .. }
            | StmtKind::EnumDecl { .. } => summary.top_level_classes += 1,
            StmtKind::Echo(exprs) => {
                summary.top_level_echoes += 1;
                if exprs.len() == 1 {
                    if let ExprKind::Lit(Literal::Str(text)) = &exprs[0].kind {
                        summary.string_bytes += text.len();
                        if text.contains('<') || text.contains('>') || text.contains('\n') {
                            summary.inline_html_echoes += 1;
                        }
                    }
                }
            }
            _ => {}
        }
    }
    summary
}

fn print_ast_summary(module: &Module) {
    let summary = summarize_module(module);
    eprintln!(
        "[vybex] AST summary: top_level_stmts={}, top_level_funcs={}, top_level_classes={}, top_level_echoes={}, inline_html_echoes={}, inline_html_bytes={}",
        summary.top_level_statements,
        summary.top_level_functions,
        summary.top_level_classes,
        summary.top_level_echoes,
        summary.inline_html_echoes,
        summary.string_bytes,
    );
}

fn should_emit_full_ast(module: &Module) -> bool {
    if std::env::var_os("VYBEX_DUMP_AST_FULL").is_some() {
        return true;
    }
    let summary = summarize_module(module);
    summary.top_level_statements <= 250
        && summary.inline_html_echoes <= 16
        && summary.string_bytes <= 16000
}

fn print_ast_outline(module: &Module) {
    for (index, stmt) in module.body.iter().enumerate() {
        let label = match &stmt.kind {
            StmtKind::Expr(_) => "Expr",
            StmtKind::Block(_) => "Block",
            StmtKind::VarDecl { .. } => "VarDecl",
            StmtKind::FunctionDecl { name, .. } => {
                println!("[{index}] FunctionDecl {name}");
                continue;
            }
            StmtKind::ClassDecl { name, .. } => {
                println!("[{index}] ClassDecl {name}");
                continue;
            }
            StmtKind::InterfaceDecl { name, .. } => {
                println!("[{index}] InterfaceDecl {name}");
                continue;
            }
            StmtKind::EnumDecl { name, .. } => {
                println!("[{index}] EnumDecl {name}");
                continue;
            }
            StmtKind::StructDecl { name, .. } => {
                println!("[{index}] StructDecl {name}");
                continue;
            }
            StmtKind::NamespaceDecl { name, .. } => {
                println!("[{index}] NamespaceDecl {name}");
                continue;
            }
            StmtKind::If { .. } => "If",
            StmtKind::For { .. } => "For",
            StmtKind::ForIn { .. } => "ForIn",
            StmtKind::While { .. } => "While",
            StmtKind::DoWhile { .. } => "DoWhile",
            StmtKind::Switch { .. } => "Switch",
            StmtKind::Return(_) => "Return",
            StmtKind::Throw { .. } => "Throw",
            StmtKind::Echo(_) => "Echo",
            StmtKind::Try { .. } => "Try",
            StmtKind::Empty => "Empty",
            _ => "Other",
        };
        println!("[{index}] {label}");
    }
}

fn print_chunk_summary(chunks: &[Chunk], filter: Option<&str>) {
    let filtered = filter_chunks(chunks, filter);
    let total_instructions: usize = filtered.iter().map(|chunk| chunk.code.len()).sum();
    eprintln!(
        "[vybex] Chunk summary: total={}, selected={}, instructions={}",
        chunks.len(),
        filtered.len(),
        total_instructions,
    );
    for chunk in filtered.iter().take(20) {
        eprintln!(
            "  → chunk '{}' (arity={}, instructions={})",
            chunk.name,
            chunk.arity,
            chunk.code.len()
        );
    }
    if filtered.len() > 20 {
        eprintln!(
            "  … {} more chunks omitted from summary",
            filtered.len() - 20
        );
    }
}

/// Every bundled language, as a `vybe_plugin::Plugin`. This is the plugin list
/// the aggregator drives through the framework. When languages become dylibs,
/// each entry becomes a `dlopen` + the module's exported `Plugin` factory.
const LANGUAGE_PLUGINS: &[&dyn vybe_plugin::Plugin] = &[
    &vybe_language_c::Plugin,
    &vybe_language_cobol::Plugin,
    &vybe_language_csharp::Plugin,
    &vybe_language_dart::Plugin,
    &vybe_language_fortran::Plugin,
    &vybe_language_go::Plugin,
    &vybe_language_java::Plugin,
    &vybe_language_js::Plugin,
    &vybe_language_lua::Plugin,
    &vybe_language_pascal::Plugin,
    &vybe_language_php::Plugin,
    &vybe_language_python::Plugin,
    &vybe_language_ruby::Plugin,
    &vybe_language_vb::Plugin,
    &vybe_language_wast::Plugin,
];

/// Register every bundled language into the shared plugin registry by running
/// each language plugin's `init` through the framework (`vybe_plugin::init_all`).
/// Global/compile-time registration — no VM needed.
pub fn register_languages() {
    vybe_plugin::init_all(LANGUAGE_PLUGINS);
}

pub fn run() {
    register_languages();
    let args: Vec<String> = std::env::args().collect();
    let mut dump = false;
    let mut dump_ast = false;
    let mut emit_wasm = false;
    let mut eval_source: Option<String> = None;
    let mut eval_language: Option<String> = None;
    let mut eval_virtual_path: Option<String> = None;
    let mut sandbox = false;
    let mut portable = false;
    let mut trace = false;
    let mut debug = false;
    let mut dap_port: Option<u16> = None;
    let mut watch = false;
    let mut chunk_filter = None;
    let mut file_arg = None;

    // --serve flags. When `serve` is true, positional args are treated as
    // [BIND] [ROOT] instead of a single script path.
    let mut serve = false;
    let mut serve_no_sandbox = false;
    let mut serve_bind: Option<String> = None;
    let mut serve_positional: Vec<String> = Vec::new();

    let mut iter = args[1..].iter();
    while let Some(arg) = iter.next() {
        match arg.as_str() {
            "--dump" | "-d" => dump = true,
            "--dump-ast" => dump_ast = true,
            "--emit-wasm" | "-w" => emit_wasm = true,
            "--eval" => {
                let Some(code) = iter.next() else {
                    eprintln!("Missing value for --eval");
                    std::process::exit(1);
                };
                eval_source = Some(code.clone());
            }
            "--lang" => {
                let Some(name) = iter.next() else {
                    eprintln!("Missing value for --lang");
                    std::process::exit(1);
                };
                eval_language = Some(name.clone());
            }
            "--virtual-path" => {
                let Some(path) = iter.next() else {
                    eprintln!("Missing value for --virtual-path");
                    std::process::exit(1);
                };
                eval_virtual_path = Some(path.clone());
            }
            "--sandbox" | "-s" => sandbox = true,
            "--portable" | "-p" => portable = true,
            "--trace" | "-t" => trace = true,
            "--debug" | "-g" => debug = true,
            "--dap-port" => {
                let Some(p) = iter.next() else {
                    eprintln!("Missing value for --dap-port");
                    std::process::exit(1);
                };
                dap_port = p.parse().ok();
            }
            "--watch" | "-W" => watch = true,
            "--serve" => serve = true,
            "--bind" => {
                let Some(bind) = iter.next() else {
                    eprintln!("Missing value for --bind");
                    std::process::exit(1);
                };
                serve_bind = Some(bind.clone());
            }
            "--no-sandbox" => serve_no_sandbox = true,
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
            _ if serve && !arg.starts_with('-') => serve_positional.push(arg.clone()),
            _ if file_arg.is_none() && !arg.starts_with('-') => file_arg = Some(arg.clone()),
            _ => {
                eprintln!("Unknown flag: {arg}");
                std::process::exit(1);
            }
        }
    }

    if serve {
        let mut config = crate::server::ServeConfig::default();
        config.no_sandbox = serve_no_sandbox || !sandbox;
        // Positional parsing: first token that looks like an addr is bind;
        // first token that looks like a path is root. Order-insensitive.
        for p in &serve_positional {
            if looks_like_addr(p) {
                config.bind = p.clone();
            } else {
                config.root = std::path::PathBuf::from(p);
            }
        }
        if let Some(bind) = serve_bind {
            config.bind = bind;
        }
        crate::server::serve_directory(config);
    }

    // ── --watch: Phase-1 hot reload. Re-run the program in a fresh subprocess
    //    on every source change. Diverges (Ctrl-C to stop). ──────────────────
    if watch {
        let Some(entry) = file_arg.clone() else {
            eprintln!("--watch requires a source file");
            std::process::exit(1);
        };
        let child_args: Vec<String> = args[1..]
            .iter()
            .filter(|a| a.as_str() != "--watch" && a.as_str() != "-W")
            .cloned()
            .collect();
        crate::watch::run_watch(PathBuf::from(entry), child_args);
    }

    let dynamic_compile_caps = if sandbox {
        vybe_host::Capabilities::safe()
    } else {
        vybe_host::Capabilities::all()
    };

    if eval_source.is_some() && file_arg.is_some() {
        eprintln!("Use either a file path or --eval, not both");
        std::process::exit(1);
    }

    if eval_source.is_some() && !dynamic_compile_caps.has(vybe_host::Capability::DynamicCompile) {
        eprintln!(
            "Dynamic compilation is disabled in the current mode (missing Capability::DynamicCompile)"
        );
        std::process::exit(1);
    }

    let file_path = if eval_source.is_none() {
        match file_arg {
            Some(f) => Some(f),
            None => {
                print_usage();
                std::process::exit(1);
            }
        }
    } else {
        None
    };

    let source_path: PathBuf;
    let bundle = if let Some(source) = eval_source.as_ref() {
        let Some(language_name) = eval_language.as_ref() else {
            eprintln!("--eval requires --lang <name>");
            std::process::exit(1);
        };
        let Some(language) = vybe_compiler::languages::find_by_name(language_name) else {
            eprintln!("Unknown language for --lang: {language_name}");
            std::process::exit(1);
        };
        source_path = eval_virtual_path
            .map(PathBuf::from)
            .unwrap_or_else(|| PathBuf::from(format!("eval.{language_name}")));
        crate::dynamic::bundle_from_source(source.clone(), language, source_path.clone())
    } else {
        let file_path = file_path.expect("file path already checked");
        source_path = PathBuf::from(&file_path);
        let path = Path::new(&file_path);

        // ── Handle .wasm binaries directly ──────────────────────────────────
        if path.extension().and_then(|e| e.to_str()) == Some("wasm") {
            if dump_ast {
                eprintln!("AST dump is only available for source files and projects");
                std::process::exit(1);
            }
            run_wasm(path, dump, trace, chunk_filter.as_deref());
            return;
        }

        match vybe_compiler::projects::load(path) {
            Ok(b) => b,
            Err(e) => {
                eprintln!("{e}");
                std::process::exit(1);
            }
        }
    };

    eprintln!(
        "[vybex] Project '{}', sources={}, entry={:?}",
        bundle.name,
        bundle.sources.len(),
        bundle.entry_point
    );
    for s in &bundle.sources {
        eprintln!("  → {} ({} bytes)", s.path.display(), s.code.len());
    }

    if dump_ast {
        eprintln!("[vybex] Preparing AST...");
        let module = match bundle.prepared_module() {
            Ok(module) => module,
            Err(e) => {
                eprintln!("Parse error: {e}");
                std::process::exit(1);
            }
        };
        print_ast_summary(&module);
        if should_emit_full_ast(&module) {
            eprintln!("[vybex] Printing full AST...");
            println!("{:#?}", module);
        } else {
            eprintln!(
                "[vybex] AST is large; printing top-level outline. Set VYBEX_DUMP_AST_FULL=1 for full debug output."
            );
            print_ast_outline(&module);
        }
        return;
    }

    // ── Set up VM first (so adapter modules can be registered ──────────────
    // against the Synthetic modules they re-export from before the
    // user program is linked).
    let mut vm = VM::new();

    let gui = if sandbox {
        eprintln!("[sandbox] Restricted mode: no filesystem, network, or database access");
        vybe_host::register_with_capabilities_and_gui(&mut vm, &vybe_host::Capabilities::safe())
    } else if portable {
        eprintln!("[portable] Running with WASM stdlib only — no Vybe host optimizations");
        vm.register_host_fn(
            "wasi:cli",
            "log",
            Box::new(
                |_ctx: &mut vybe_bytecode::HostContext, args: &[vybe_bytecode::Value]| {
                    for a in args {
                        print!("{}", a);
                    }
                    println!();
                    vybe_bytecode::Value::Null
                },
            ),
        );
        vm.register_host_fn(
            "wasi:cli",
            "readLine",
            Box::new(|_ctx: &mut vybe_bytecode::HostContext, _| {
                let mut line = String::new();
                std::io::stdin().read_line(&mut line).ok();
                vybe_bytecode::Value::String(std::sync::Arc::from(line.trim()))
            }),
        );
        Arc::new(Mutex::new(vybe_host::gui_state::GuiState::new()))
    } else {
        vybe_host::register_all_with_gui(&mut vm)
    };

    if !portable {
        vybe_host::setup_namespaces(&mut vm);
    }

    // WAST script runtime — call_indirect and try/catch are WASM VM-level
    // constructs; the WAST walker routes them through these host stubs which
    // the vybex runner provides so WAT/WAST example files can run end-to-end.
    // Programmatic-mode server primitive: scripts can call
    // `vybe:http/server.listen(addr, handler)` to become a long-lived
    // HTTP server (Node/Flask/Sinatra style).
    crate::server::programmatic::register(&mut vm);

    // Register any in-language adapter modules whose targets are real.
    // Today this is intentionally empty; we do not install placeholder
    // adapters over non-existent `wasi:http/*` server surfaces.
    if let Err(e) = crate::adapters::register_all(&mut vm) {
        eprintln!("adapter registration error: {e}");
        std::process::exit(1);
    }

    // ── Compile ─────────────────────────────────────────────────────────────
    eprintln!("[vybex] Preparing and compiling module...");
    let compiled = {
        let mut runtime_compiler = crate::dynamic::RuntimeCompilerService::with_capabilities(
            &mut vm,
            dynamic_compile_caps.clone(),
        );
        match (&eval_source, &eval_language) {
            (Some(source), Some(language_name)) => {
                match runtime_compiler.compile_source_by_name(
                    source.clone(),
                    language_name,
                    source_path.clone(),
                ) {
                    Ok(c) => c,
                    Err(e) => {
                        eprintln!("Compile error: {e}");
                        std::process::exit(1);
                    }
                }
            }
            _ => match runtime_compiler.compile_bundle(&bundle) {
                Ok(c) => c,
                Err(e) => {
                    eprintln!("Compile error: {e}");
                    std::process::exit(1);
                }
            },
        }
    };

    // ── --dump: disassemble and exit ────────────────────────────────────────
    if dump {
        print_chunk_summary(&compiled.chunks, chunk_filter.as_deref());
        for chunk in filter_chunks(&compiled.chunks, chunk_filter.as_deref()) {
            println!("{}", vybe_bytecode::debug::disassemble(chunk));
        }
        return;
    }

    // ── --emit-wasm: write .wasm binary and exit ────────────────────────────
    if emit_wasm {
        let wasm_bytes = vybe_platform_wasm::write_wasm(&compiled.chunks);
        let out_path = source_path.with_extension("wasm");
        std::fs::write(&out_path, &wasm_bytes).unwrap();
        eprintln!("Wrote {} bytes to {}", wasm_bytes.len(), out_path.display());
        return;
    }

    // VM was set up above, before compilation, so adapter modules
    // could be registered against the Synthetic modules they re-export
    // from. Apply the trace flag now.
    if trace {
        vm.set_trace(true);
        vm.set_trace_chunk_filter(chunk_filter.clone());
    }

    // ── Debugger (--debug REPL or --dap-port VS Code): pause on entry ────────
    if debug || dap_port.is_some() {
        // Install the compiler-backed expression evaluator (for `p <expr>`,
        // conditional breakpoints, watches, and DAP `evaluate`). Faithful
        // semantics via an isolated mini-VM; never perturbs the paused VM.
        let eval_language = bundle.language;
        let eval_caps = dynamic_compile_caps.clone();
        vm.set_eval_hook(Box::new(move |live, expr, locals| {
            crate::dynamic::debug_eval_expression(live, expr, locals, eval_language, eval_caps.clone())
        }));

        // Install the hot-reload recompiler: re-read + recompile the source in a
        // FRESH VM set up exactly like the original compile, so unchanged
        // functions reproduce byte-for-byte (the live VM's module state has
        // drifted since startup and would produce spurious diffs). `apply_reload`
        // then swaps only the functions that actually changed.
        let reload_path = source_path.clone();
        let reload_caps = dynamic_compile_caps.clone();
        vm.set_reload_hook(Box::new(move |_live| {
            recompile_for_reload(&reload_path, reload_caps.clone())
        }));
        if let Some(port) = dap_port {
            crate::dap::attach(&mut vm, port, source_path.display().to_string());
        } else {
            crate::debug_repl::attach(&mut vm);
        }
    }

    // ── Run ─────────────────────────────────────────────────────────────────
    let mut runtime_compiler =
        crate::dynamic::RuntimeCompilerService::with_capabilities(&mut vm, dynamic_compile_caps);
    match runtime_compiler.run_compiled(compiled) {
        Ok(v) => {
            if debug {
                eprintln!("\n● program exited → {v}");
                std::process::exit(0);
            }
            if dap_port.is_some() {
                // The client sees the socket close and ends the session.
                std::process::exit(0);
            }
        }
        Err(e) if e.contains("__debug_quit__") => {
            eprintln!("\n● debugger quit");
            std::process::exit(0);
        }
        Err(e) if e.contains("__debug_restart__") => {
            // Replace this process with a fresh copy (same args) — clean restart.
            use std::os::unix::process::CommandExt;
            eprintln!("\n↻ restarting…");
            let args: Vec<String> = std::env::args().skip(1).collect();
            let exe = std::env::current_exe().unwrap_or_else(|_| PathBuf::from("vybex"));
            let _ = std::process::Command::new(exe).args(&args).exec();
            std::process::exit(0);
        }
        Err(e) => {
            eprintln!("Runtime error: {e}");
            std::process::exit(1);
        }
    }

    if gui.lock().unwrap().should_run {
        crate::gui_launch::launch_gui(vm, gui);
    }
}

/// Recompile the source for a hot reload in a FRESH VM whose host/module setup
/// mirrors the original compile (so unchanged functions reproduce identically).
/// Returns the fresh chunk set for `VM::debug_reload` to diff and swap.
fn recompile_for_reload(
    source_path: &Path,
    caps: vybe_host::Capabilities,
) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    let bundle = vybe_compiler::projects::load(source_path).map_err(|e| e.to_string())?;
    let mut tv = vybe_bytecode::VM::new();
    let _gui = vybe_host::register_all_with_gui(&mut tv);
    vybe_host::setup_namespaces(&mut tv);
    crate::server::programmatic::register(&mut tv);
    let _ = crate::adapters::register_all(&mut tv);
    let compiled = crate::dynamic::RuntimeCompilerService::with_capabilities(&mut tv, caps)
        .compile_bundle(&bundle)?;
    Ok(compiled.chunks)
}

fn run_wasm(path: &Path, dump: bool, trace: bool, chunk_filter: Option<&str>) {
    let data = match std::fs::read(path) {
        Ok(d) => d,
        Err(e) => {
            eprintln!("Error reading {}: {e}", path.display());
            std::process::exit(1);
        }
    };
    eprintln!("Loading WASM: {} ({} bytes)", path.display(), data.len());

    let chunks = match vybe_platform_wasm::read_wasm(&data) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("WASM error: {e}");
            std::process::exit(1);
        }
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
        Err(e) => {
            eprintln!("Runtime error: {e}");
            std::process::exit(1);
        }
    }

    if gui.lock().unwrap().should_run {
        crate::gui_launch::launch_gui(vm, gui);
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
    let exts = vybe_compiler::projects::supported_extensions();
    let ext_list: Vec<String> = exts.iter().map(|e| format!(".{e}")).collect();
    eprintln!("vybex — Universal compiler");
    eprintln!();
    eprintln!("Usage: vybex [flags] <file>");
    eprintln!("       vybex --eval CODE --lang NAME [--virtual-path PATH]");
    eprintln!("       vybex --serve [--bind BIND] [BIND] [ROOT]");
    eprintln!();
    eprintln!("Flags:");
    eprintln!("  -d, --dump        Disassemble bytecode (no run)");
    eprintln!("      --dump-ast    Parse and print the prepared common AST");
    eprintln!("  -w, --emit-wasm   Compile to .wasm binary");
    eprintln!("      --eval CODE   Compile source from a string");
    eprintln!("      --lang NAME   Language for --eval (js, php, python, vb, ...)");
    eprintln!("      --virtual-path PATH  Source path used for relative imports in --eval");
    eprintln!("  -s, --sandbox     Restricted mode (safe capabilities only)");
    eprintln!("  -p, --portable    Minimal WASI runtime (no Vybe host)");
    eprintln!("  -t, --trace       Enable bytecode trace output");
    eprintln!("  -g, --debug       Step debugger: pause on entry, REPL on stdin (h for help)");
    eprintln!("      --dap-port N  Debug Adapter Protocol server on 127.0.0.1:N (VS Code attach)");
    eprintln!("  -W, --watch       Re-run on source change (Phase-1 hot reload)");
    eprintln!("      --chunk NAME  Limit --dump/--trace output to a chunk name or index");
    eprintln!("      --serve       Start HTTP server for a directory (see httpserver.md)");
    eprintln!("      --bind ADDR   With --serve: bind to ADDR instead of 127.0.0.1:8080");
    eprintln!("                    BIND defaults to 127.0.0.1:8080, ROOT to current dir");
    eprintln!("      --no-sandbox  With --serve: keep full host access (default)");
    eprintln!("  -h, --help        Show this help");
    eprintln!();
    eprintln!("Supported: {}", ext_list.join(", "));
}

/// Heuristic to distinguish a bind address (`host:port` or `:port`) from a
/// filesystem root path in the positional args of `--serve`.
fn looks_like_addr(s: &str) -> bool {
    // Bare port `:3000`
    if let Some(rest) = s.strip_prefix(':') {
        return rest.chars().all(|c| c.is_ascii_digit());
    }
    // IPv6 bracketed `[::1]:8080`
    if s.starts_with('[') && s.contains("]:") {
        return true;
    }
    // host:port — exactly one colon, port is numeric, host has no slashes.
    if let Some((h, p)) = s.rsplit_once(':') {
        return !h.contains('/')
            && !h.contains('\\')
            && !p.is_empty()
            && p.chars().all(|c| c.is_ascii_digit());
    }
    false
}
