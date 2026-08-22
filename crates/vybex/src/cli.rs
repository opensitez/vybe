//! vybex — Universal compiler.
//!
//! Usage: vybex [flags] <file...|project>
//!
//! Flags:
//!   --dump, -d        Disassemble bytecode and exit (no run)
//!   --dump-ast        Parse and print the prepared common AST, then exit
//!   --emit-wasm, -w   Compile to .wasm binary and exit
//!   --entry, -e NAME  Override the entry symbol (ld -e style; Class.Method for static methods)
//!   --eval CODE       Compile and run source from a string
//!   --lang NAME       Language to use with --eval
//!   --virtual-path P  Virtual source path to use with --eval
//!   --sandbox, -s     Restricted mode (no filesystem/network/database)
//!   --portable, -p    Minimal WASI runtime only (no Vybe host optimizations)
//!   --trace, -t       Enable bytecode trace output
//!   --chunk <name>    Limit --dump/--trace output to a specific chunk
//!   --capture FILE    Render one GUI frame to a PNG instead of opening a window
//!   --capture-control N  Crop --capture to a single control
//!
//! Supports single source files (detected by extension), MULTIPLE source files
//! linked together like a C compiler (`vybex main.c util.c` — the first file is
//! the entry point and all must share one language), project files (.vybe,
//! .vbproj, .csproj, .pyproj/.ipyproj), and .wasm binaries.
//! Language is determined automatically.

use std::path::{Path, PathBuf};
use vybe_compiler::ast::{ExprKind, Literal, Module, StmtKind};
use vybe_runtime::VM;
use vybe_runtime::chunk::Chunk;

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

/// Print the full AST of just the top-level declarations named `name`
/// (functions, classes, structs, interfaces, enums, namespaces), matched
/// case-insensitively so it works for the case-folding languages too.
/// Returns how many matched.
fn print_ast_for_named(module: &Module, name: &str) -> usize {
    let mut matched = 0;
    for stmt in &module.body {
        let decl_name = match &stmt.kind {
            StmtKind::FunctionDecl { name, .. }
            | StmtKind::ClassDecl { name, .. }
            | StmtKind::InterfaceDecl { name, .. }
            | StmtKind::EnumDecl { name, .. }
            | StmtKind::StructDecl { name, .. }
            | StmtKind::NamespaceDecl { name, .. } => name.as_str(),
            _ => continue,
        };
        if decl_name.eq_ignore_ascii_case(name) {
            matched += 1;
            println!("{:#?}", stmt);
        }
    }
    matched
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

/// Registration for the whole binary: ONE loop over the ONE registry.
///
/// There is no plugin list here. Every plugin crate — language, platform,
/// host-function provider, all the same `vybe_runtime::Plugin` — submits
/// itself at link time, and this runs whatever `vybex` linked. Adding or
/// removing a plugin is a Cargo edit; this file never changes.

/// Every plugin registers into `vm` in a single loop. Non-GUI (drawing-only) —
/// installs `vybe:gui` no-op stubs so compiled control/form code doesn't hit
/// unresolved imports.
pub fn register_plugins(
    vm: &mut vybe_runtime::VM,
    caps: &vybe_runtime::capabilities::Capabilities,
) {
    vybe_runtime::init_all_registered(vm, caps);
    if vm
        .host_registry
        .get(&("vybe:gui".to_string(), "controlSetProperty".to_string()))
        .is_none()
    {
        vybe_platform_vybe::register_gui_stubs(vm);
    }
}

/// GUI variant of [`register_plugins`]: a fresh `GuiState` is installed before
/// the same one loop runs, and the shared handle is returned for the form
/// launcher.
pub fn register_plugins_with_gui(
    vm: &mut vybe_runtime::VM,
    caps: &vybe_runtime::capabilities::Capabilities,
) -> std::sync::Arc<std::sync::Mutex<vybe_platform_vybe::gui_state::GuiState>> {
    let vybe = vybe_platform_vybe::Plugin::with_gui();
    vybe_runtime::init_all_registered(vm, caps);
    vybe.gui_state()
        .expect("with_gui() always installs a GuiState")
}

pub fn run() {
    let args: Vec<String> = std::env::args().collect();
    let mut dump = false;
    let mut dump_ast = false;
    let mut emit_wasm = false;
    // `--check` compiles the program and reports diagnostics WITHOUT running it,
    // exiting non-zero on any parse/compile error. Intended for editors and CI.
    let mut check = false;
    let mut eval_source: Option<String> = None;
    let mut eval_language: Option<String> = None;
    let mut eval_virtual_path: Option<String> = None;
    let mut sandbox = false;
    let mut worker = false;
    let mut portable = false;
    let mut trace = false;
    let mut debug = false;
    let mut dap_port: Option<u16> = None;
    let mut watch = false;
    // `--capture <png>` renders one GUI frame offscreen instead of opening a
    // window; `--capture-control <name>` crops it to a single control.
    let mut capture: Option<String> = None;
    let mut capture_control: Option<String> = None;
    let mut chunk_filter = None;
    // `--entry <name>` — override the linker's default entry symbol, `ld -e`
    // style (`name` for a free function, `Class.Method` for a static method).
    // Removes the need for a project file just to pick the entry point.
    let mut entry_override: Option<String> = None;
    // Every non-flag positional is a source file. Several files link into one
    // bundle (C-compiler style); the first is the entry file.
    let mut file_args: Vec<String> = Vec::new();

    // --serve flags. When `serve` is true, positional args are treated as
    // [BIND] [ROOT] instead of a single script path.
    let mut serve = false;
    let mut serve_no_sandbox = false;
    let mut serve_cold = false;
    let mut serve_no_cache = false;
    let mut serve_pool: usize = 0;
    let mut serve_bind: Option<String> = None;
    let mut serve_positional: Vec<String> = Vec::new();

    let mut iter = args[1..].iter();
    while let Some(arg) = iter.next() {
        match arg.as_str() {
            "--dump" | "-d" => dump = true,
            "--dump-ast" => dump_ast = true,
            "--check" | "-c" => check = true,
            "--emit-wasm" | "-w" => emit_wasm = true,
            "--entry" | "-e" => {
                let Some(name) = iter.next() else {
                    eprintln!("Missing value for --entry");
                    std::process::exit(1);
                };
                entry_override = Some(name.clone());
            }
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
            "--worker" => worker = true,
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
            "--capture" => {
                let Some(p) = iter.next() else {
                    eprintln!("Missing value for --capture");
                    std::process::exit(1);
                };
                capture = Some(p.clone());
            }
            "--capture-control" => {
                let Some(c) = iter.next() else {
                    eprintln!("Missing value for --capture-control");
                    std::process::exit(1);
                };
                capture_control = Some(c.clone());
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
            "--cold" => serve_cold = true,
            "--no-cache" => serve_no_cache = true,
            "--pool" => {
                let Some(n) = iter.next().and_then(|v| v.parse::<usize>().ok()) else {
                    eprintln!("--pool requires a positive count");
                    std::process::exit(1);
                };
                serve_pool = n;
            }
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
            "--version" | "-V" => {
                println!("vybex {}", env!("CARGO_PKG_VERSION"));
                return;
            }
            "--list-languages" => {
                print_languages();
                return;
            }
            _ if serve && !arg.starts_with('-') => serve_positional.push(arg.clone()),
            _ if !arg.starts_with('-') => file_args.push(arg.clone()),
            _ => {
                eprintln!("Unknown flag: {arg}");
                std::process::exit(1);
            }
        }
    }

    if serve {
        let mut config = crate::server::ServeConfig::default();
        config.no_sandbox = serve_no_sandbox || !sandbox;
        config.cold = serve_cold;
        config.pool = serve_pool;
        config.no_cache = serve_no_cache;
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
        let Some(entry) = file_args.first().cloned() else {
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
        vybe_runtime::capabilities::Capabilities::safe()
    } else {
        vybe_runtime::capabilities::Capabilities::all()
    };

    // ── Warm execution mode ─────────────────────────────────────────────────
    // Boot once, then run program after program against a reset VM. Takes over
    // the process: jobs arrive on stdin, so there is no entry file to parse and
    // none of the flags below apply.
    if worker {
        crate::worker::run(dynamic_compile_caps);
    }

    // ── One registration ──────────────────────────────────────────────────
    // Create the VM and run THE single plugin loop (all 20 — languages AND
    // platforms) ONCE, before compiling. `find_by_name`/the compiler resolve
    // languages through `registry::all()` (populated by each language plugin's
    // `init`), so this must precede the compile below; the same pass registers
    // the platform host fns the runtime needs. Portable mode adds its minimal
    // `wasi:cli` stubs on top.
    let mut vm = VM::new();
    if sandbox {
        eprintln!("[sandbox] Restricted mode: no filesystem, network, or database access");
    } else if portable {
        eprintln!("[portable] Running with WASM stdlib only — no Vybe host optimizations");
    }
    let gui = register_plugins_with_gui(&mut vm, &dynamic_compile_caps);
    if portable {
        vm.register_host_fn(
            "wasi:cli",
            "log",
            Box::new(
                |_ctx: &mut vybe_runtime::HostContext, args: &[vybe_runtime::Value]| {
                    for a in args {
                        print!("{}", a);
                    }
                    println!();
                    vybe_runtime::Value::Null
                },
            ),
        );
        vm.register_host_fn(
            "wasi:cli",
            "readLine",
            Box::new(|_ctx: &mut vybe_runtime::HostContext, _| {
                let mut line = String::new();
                std::io::stdin().read_line(&mut line).ok();
                vybe_runtime::Value::String(std::sync::Arc::from(line.trim()))
            }),
        );
    }

    if eval_source.is_some() && !file_args.is_empty() {
        eprintln!("Use either a file path or --eval, not both");
        std::process::exit(1);
    }

    if eval_source.is_some()
        && !dynamic_compile_caps.has(vybe_runtime::capabilities::Capability::DynamicCompile)
    {
        eprintln!(
            "Dynamic compilation is disabled in the current mode (missing Capability::DynamicCompile)"
        );
        std::process::exit(1);
    }

    let file_paths: Vec<PathBuf> = if eval_source.is_none() {
        if file_args.is_empty() {
            print_usage();
            std::process::exit(1);
        }
        file_args.iter().map(PathBuf::from).collect()
    } else {
        Vec::new()
    };

    let source_path: PathBuf;
    // Units of a multi-language program that run BEFORE the entry unit —
    // each compiled through its own language, then linked in the shared VM.
    // Empty for the ordinary single-language case.
    let mut secondary_units: Vec<vybe_compiler::bundle::Bundle> = Vec::new();
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
        // The entry file (first positional) supplies the source path used for
        // diagnostics and for `.wasm` dispatch; the rest link alongside it.
        let entry = file_paths
            .first()
            .expect("file path already checked")
            .clone();
        source_path = entry.clone();
        let path = entry.as_path();

        // ── Handle .wasm binaries directly ──────────────────────────────────
        if path.extension().and_then(|e| e.to_str()) == Some("wasm") {
            if dump_ast {
                eprintln!("AST dump is only available for source files and projects");
                std::process::exit(1);
            }
            if check {
                eprintln!("--check is only available for source files and projects");
                std::process::exit(1);
            }
            run_wasm(
                path,
                dump,
                trace,
                chunk_filter.as_deref(),
                capture.as_deref(),
                capture_control.as_deref(),
            );
            return;
        }

        match vybe_compiler::projects::load_program(&file_paths) {
            Ok(p) => {
                secondary_units = p.units;
                secondary_units
                    .pop()
                    .expect("a program has at least one unit")
            }
            Err(e) => {
                eprintln!("{e}");
                std::process::exit(1);
            }
        }
    };

    let mut bundle = bundle;
    if let Some(spec) = entry_override.as_ref() {
        bundle.entry_point = match spec.split_once('.') {
            Some((class, method)) => {
                vybe_compiler::bundle::EntryPoint::Method(class.to_string(), method.to_string())
            }
            None => vybe_compiler::bundle::EntryPoint::Function(spec.clone()),
        };
    }
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
        // `--chunk` narrows --dump and --trace to one function; do the same
        // here. Without it, reading one function out of a large program means
        // scrolling tens of thousands of lines of `{:#?}`.
        if let Some(name) = chunk_filter.as_deref() {
            let matched = print_ast_for_named(&module, name);
            if matched == 0 {
                eprintln!("[vybex] no top-level declaration named `{name}`");
                std::process::exit(1);
            }
            return;
        }
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

    // (VM + the single plugin registration already happened above, before the
    // compile — see "One registration".)

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

    // ── --check: report success and exit WITHOUT running ────────────────────
    // Reaching here means every unit parsed and compiled; any parse/compile
    // error already exited non-zero above. No program code is executed.
    if check {
        println!(
            "OK: {} compiled successfully ({} chunk(s))",
            source_path.display(),
            compiled.chunks.len()
        );
        return;
    }

    // What the program DECLARED about presenting a UI
    // (`vybe_ast::Directives::app_shell`). Read here because `run_compiled`
    // consumes the compilation, and needed after it — the launch decision is
    // made once the program has run.
    let declared_shell = compiled.app_shell;

    // ── --dump: disassemble and exit ────────────────────────────────────────
    if dump {
        print_chunk_summary(&compiled.chunks, chunk_filter.as_deref());
        for chunk in filter_chunks(&compiled.chunks, chunk_filter.as_deref()) {
            println!("{}", vybe_runtime::debug::disassemble(chunk));
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
            crate::dynamic::debug_eval_expression(
                live,
                expr,
                locals,
                eval_language,
                eval_caps.clone(),
            )
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

        // Install the event simulator, so the debugger can fire a click or a
        // window-close with no OS window.
        //
        // `OnClick := h` IS `addEventListener("click", h)` for every frontend,
        // so the wiring is on the ELEMENT and a listener is invoked the way the
        // document invokes one: with an `Event` and nothing else. Its receiver
        // is already bound into the handler by `primitives/gui.rs`, which is
        // why no form object is looked up here — the same call the window makes
        // (`gui_launch::dispatch_document_events`), so the two cannot drift.
        //
        // `GuiState` is the fallback, and only for what is NOT a DOM event: a
        // form's `Load`, and a designer form's handlers. That path keeps the
        // arity rule it always had (0→[], 1→[me], 2→[me, sender]).
        let fire_gui = gui.clone();
        vm.set_event_fire_hook(Box::new(move |vm, control, event| {
            // A DOM type is lowercase where the debugger's word is `Click`;
            // `listeners_for` folds the case, so they are the same event.
            if let Some(node) = crate::gui_document::node_by_id(control) {
                if let Some(cb) = crate::gui_document::listeners_for(node, event)
                    .into_iter()
                    .next()
                {
                    let evt = crate::gui_document::event_object(event, node);
                    return Ok(vm.invoke_callback(&cb, &[evt]));
                }
            }
            let (handler, form_object) = {
                let g = fire_gui
                    .lock()
                    .map_err(|_| "gui state unavailable".to_string())?;
                (
                    g.get_event_handler(control, event).cloned(),
                    g.form_object.clone(),
                )
            };
            let Some(cb) = handler else {
                return Err(format!(
                    "no `{event}` handler on `{control}` (see `widgets` for wired events)"
                ));
            };
            let me = form_object
                .or_else(|| vm.global("__f").cloned())
                .unwrap_or(vybe_runtime::Value::Null);
            let sender = vybe_runtime::Value::String(std::sync::Arc::from(control));
            let args: Vec<vybe_runtime::Value> = match crate::gui_launch::fn_arity(&cb) {
                0 => vec![],
                1 => vec![me],
                2 => vec![me, sender],
                _ => vec![me, sender, vybe_runtime::Value::Null],
            };
            Ok(vm.invoke_callback(&cb, &args))
        }));
        if let Some(port) = dap_port {
            crate::dap::attach(&mut vm, port, source_path.display().to_string());
        } else {
            crate::debug_repl::attach(&mut vm, gui.clone());
        }
    }

    // ── Run ─────────────────────────────────────────────────────────────────
    let mut runtime_compiler =
        crate::dynamic::RuntimeCompilerService::with_capabilities(&mut vm, dynamic_compile_caps);

    // Link the other languages in first. Each was parsed and compiled by its
    // own front-end; loading it here puts its functions and classes in the
    // shared global table and runs its top-level code, so the entry unit
    // starts with everything already defined.
    for unit in &secondary_units {
        eprintln!("[vybex] Linking {} ({})", unit.name, unit.language.name);
        if let Err(e) = runtime_compiler.run_program_unit(unit) {
            eprintln!("Error in {} ({}): {e}", unit.name, unit.language.name);
            std::process::exit(1);
        }
    }

    match runtime_compiler.run_compiled(compiled) {
        Ok(v) => {
            // A GUI program hasn't really finished when `run_compiled` returns —
            // the window/event loop is launched below. Under the debugger we
            // must let that happen (so breakpoints in click handlers fire, via
            // `vm.invoke` re-entering the instrumented dispatch loop); only
            // report "exited" for a program with no GUI. Without this, `--debug`
            // on a GUI app exited before the window ever showed.
            //
            // The SAME question the launch gate asks, so the debugger cannot
            // decide a run is over while the gate goes on to open a window.
            let gui_should_run =
                should_present(gui.lock().unwrap().should_run, declared_shell);
            if debug && !gui_should_run {
                eprintln!("\n● program exited → {v}");
                std::process::exit(0);
            }
            if dap_port.is_some() && !gui_should_run {
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

    // The status the guest handed `wasi:cli/exit.exit-with-code` — `sys.exit(3)`,
    // `System.exit(4)`, `halt(2)`, `STOP RUN`. The VM only CARRIES it: it ends
    // the run and hands control back, exactly as the component model requires of
    // a guest instance, and never calls `process::exit` itself (that would kill
    // the embedder — a test binary, the server — on the first `exit`). Turning a
    // status into a process exit is the embedder's job, which is here. Before
    // this the code was dropped at the host boundary and all four exited 0.
    //
    // Read BEFORE the GUI check: a program that exits non-zero must exit, not
    // open a window. Zero needs no action — falling off the end is already
    // exit 0, and taking this branch would skip a legitimate GUI launch.
    if vm.pending_exit_code != 0 {
        std::process::exit(vm.pending_exit_code);
    }

    if should_present(gui.lock().unwrap().should_run, declared_shell) {
        match capture {
            Some(path) => run_capture(vm, gui, &path, capture_control.as_deref()),
            None => crate::gui_launch::launch_gui(vm, gui),
        }
    }
}

/// Does this run end in a window?
///
/// Three answers, because three different things know, and only asking one of
/// them was the bug:
///
/// - **The program declared it.** `AppShell::Windowed` covers a UI built LATER —
///   from a timer or an event handler — which no test taken at this instant can
///   see. `Headless` is the only way to say "builds controls, must not present",
///   and it wins outright.
/// - **The document has content.** A page is not told to run; it runs because it
///   HAS a document. This is what every converted frontend relies on, and it is
///   the same test `gui_document::with_live` and `render_into` paint by.
/// - **`GuiState::should_run`.** The legacy answer, set by
///   `vybe:gui.runApplication`. Kept while frontends still call it, so nothing
///   regresses as they convert one at a time.
///
/// Asking only the third meant a converted program drew every frame and exited
/// without ever showing a window.
fn should_present(
    gui_should_run: bool,
    declared: Option<vybe_compiler::ast::AppShell>,
) -> bool {
    match declared {
        Some(vybe_compiler::ast::AppShell::Headless) => false,
        Some(vybe_compiler::ast::AppShell::Windowed) => true,
        // A BROWSING CONTEXT is the window (HTML §7.1). A browser that opens
        // one shows a tab before any content arrives — `about:blank` is a
        // window with nothing in it — so the question is whether a context was
        // opened, never what it ended up containing.
        //
        // This used to ask `has_content()` (`control_count() > 0`), which is a
        // different and wrong question: it closed the window on any page that
        // builds its UI later — from `load`, a timer, a promise — and on any
        // page that legitimately has none yet. It also made a program's
        // UI-ness depend on how far it happened to get.
        //
        // Opening the context is deliberate: `active_document` creates on first
        // use, so a program that never touches the DOM never has one and stays
        // a console program.
        None => gui_should_run || vybe_platform_web::html::has_browsing_context(),
    }
}

/// Write one offscreen GUI frame to `path`, reporting the result on stderr.
/// A capture that finds no frame is an ERROR exit, not a silent empty file —
/// the whole point is to be usable as a check.
fn run_capture(
    vm: vybe_runtime::VM,
    gui: std::sync::Arc<std::sync::Mutex<vybe_platform_vybe::gui_state::GuiState>>,
    path: &str,
    control: Option<&str>,
) {
    match crate::gui_launch::capture_gui(vm, gui, path, control) {
        Ok((w, h)) => eprintln!("[vybex] captured {w}x{h} → {path}"),
        Err(e) => {
            eprintln!("[vybex] capture failed: {e}");
            std::process::exit(1);
        }
    }
}

/// Recompile the source for a hot reload in a FRESH VM whose host/module setup
/// mirrors the original compile (so unchanged functions reproduce identically).
/// Returns the fresh chunk set for `VM::debug_reload` to diff and swap.
fn recompile_for_reload(
    source_path: &Path,
    caps: vybe_runtime::capabilities::Capabilities,
) -> Result<Vec<vybe_runtime::Chunk>, String> {
    let bundle = vybe_compiler::projects::load(source_path).map_err(|e| e.to_string())?;
    let mut tv = vybe_runtime::VM::new();
    let _gui = register_plugins_with_gui(&mut tv, &vybe_runtime::capabilities::Capabilities::all());
    crate::server::programmatic::register(&mut tv);
    let _ = crate::adapters::register_all(&mut tv);
    let compiled = crate::dynamic::RuntimeCompilerService::with_capabilities(&mut tv, caps)
        .compile_bundle(&bundle)?;
    Ok(compiled.chunks)
}

fn run_wasm(
    path: &Path,
    dump: bool,
    trace: bool,
    chunk_filter: Option<&str>,
    capture: Option<&str>,
    capture_control: Option<&str>,
) {
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
            println!("{}", vybe_runtime::debug::disassemble(chunk));
        }
        return;
    }

    let mut vm = VM::new();
    let gui = register_plugins_with_gui(&mut vm, &vybe_runtime::capabilities::Capabilities::all());

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

    // Same contract as the source path: the guest's `exit-with-code` status
    // becomes this process's status, and only the embedder does that.
    if vm.pending_exit_code != 0 {
        std::process::exit(vm.pending_exit_code);
    }

    // A prebuilt `.wasm` carries no directives — the declaration lives in the
    // AST, and this path starts from bytecode — so the document answers alone.
    if should_present(gui.lock().unwrap().should_run, None) {
        match capture {
            Some(path) => run_capture(vm, gui, path, capture_control),
            None => crate::gui_launch::launch_gui(vm, gui),
        }
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
    eprintln!("Usage: vybex [flags] <file...>   (several files link together)");
    eprintln!("       vybex --eval CODE --lang NAME [--virtual-path PATH]");
    eprintln!("       vybex --serve [--bind BIND] [BIND] [ROOT]");
    eprintln!();
    eprintln!("Flags:");
    eprintln!("  -d, --dump        Disassemble bytecode (no run)");
    eprintln!("      --dump-ast    Parse and print the prepared common AST");
    eprintln!("  -c, --check       Compile and report errors without running (exit 1 on error)");
    eprintln!("  -w, --emit-wasm   Compile to .wasm binary");
    eprintln!("      --eval CODE   Compile source from a string");
    eprintln!("      --lang NAME   Language for --eval (js, php, python, vb, ...)");
    eprintln!("      --virtual-path PATH  Source path used for relative imports in --eval");
    eprintln!(
        "  -s, --sandbox     Restricted mode (safe capabilities only)
      --worker      Warm mode: boot once, run a program per stdin line
                    (reset between each — no relaunch, no re-registration)"
    );
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
    eprintln!("      --pool N      With --serve: warm VM workers (default: one per core)");
    eprintln!("      --cold        With --serve: fresh VM per request instead of the warm pool");
    eprintln!("      --no-cache    With --serve: recompile on every request");
    eprintln!("      --list-languages  List registered language frontends and their extensions");
    eprintln!("  -V, --version     Print the vybex version and exit");
    eprintln!("  -h, --help        Show this help");
    eprintln!();
    eprintln!("Supported: {}", ext_list.join(", "));
}

/// `--list-languages` — print every registered language frontend and the file
/// extensions it claims, grouped by language. Both are read from the live
/// plugin registry (`languages::all` / `supported_extensions`), so the list
/// always matches the frontends actually compiled into this build.
fn print_languages() {
    use std::collections::BTreeMap;

    // Seed with every registered language so a frontend with no extension
    // mapping still appears, then attach each supported extension to its owner.
    let mut by_lang: BTreeMap<&'static str, Vec<String>> = BTreeMap::new();
    for lang in vybe_compiler::languages::all() {
        by_lang.entry(lang.name).or_default();
    }
    for ext in vybe_compiler::languages::supported_extensions() {
        if let Some(lang) = vybe_compiler::languages::find_by_extension(&ext) {
            by_lang.entry(lang.name).or_default().push(format!(".{ext}"));
        }
    }

    println!("vybex — {} registered language frontends:", by_lang.len());
    println!();
    let width = by_lang.keys().map(|n| n.len()).max().unwrap_or(0);
    for (name, exts) in &by_lang {
        let exts = if exts.is_empty() {
            "(no file extensions)".to_string()
        } else {
            exts.join(", ")
        };
        println!("  {:<width$}  {exts}", name);
    }
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
