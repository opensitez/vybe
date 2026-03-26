use std::fs;
use std::path::{Path, PathBuf};

use dioxus::prelude::*;
use dioxus::desktop::{Config, WindowBuilder};

use vybe_parser_basic::parse_program;
use vybe_project::Project;
use vybe_runtime::{Interpreter, ResourceEntry, RuntimeSideEffect};

use crate::runtime_panel::RuntimeProject;
use crate::FormRunner;

// ---------------------------------------------------------------------------
// Thread-local used to pass the Project into the named Dioxus App component.
// (Dioxus `launch()` requires a plain fn-pointer, so we can't use a closure.)
// ---------------------------------------------------------------------------
thread_local! {
    pub static LAUNCH_PROJECT: std::cell::RefCell<Option<Project>> = std::cell::RefCell::new(None);
    pub static LAUNCH_TITLE: std::cell::RefCell<String> = std::cell::RefCell::new(String::new());
}

// ---------------------------------------------------------------------------
// Public entry point – the ONLY function the shell binary calls.
// ---------------------------------------------------------------------------

/// Run a Visual Basic or JavaScript file or project.
///
/// * `.vb`    → parse & run as console program
/// * `.vbp`   → load VB6 project, run as form or console
/// * `.vbproj` → load VB.NET project, run as form or console
/// * `.js`    → parse, compile to bytecode & run via VM
///
/// `extra_args` are the command-line arguments passed *after* the project file,
/// available to the VB program via `Command()` or `Environment.GetCommandLineArgs()`.
pub fn run(path: &Path, extra_args: &[String]) {
    let ext = path
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default();

    match ext.as_str() {
        "vb" => run_vb_file(path, extra_args),
        "vbp" | "vbproj" => run_project(path, extra_args),
        "js" => run_js_file(path),
        _ => {
            eprintln!(
                "Error: unsupported file type '.{}'. Expected .vb, .vbp, .vbproj, or .js",
                ext
            );
            std::process::exit(1);
        }
    }
}

// ---------------------------------------------------------------------------
// Internals
// ---------------------------------------------------------------------------

/// Run a standalone .js file via the bytecode VM.
fn run_js_file(path: &Path) {
    let code = match fs::read_to_string(path) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("Error reading file: {e}");
            std::process::exit(1);
        }
    };

    let program = match vybe_parser_js::parse(&code) {
        Ok(p) => p,
        Err(e) => {
            eprintln!("JS parse error: {e}");
            std::process::exit(1);
        }
    };

    // 1. Create a bare VM
    let mut vm = vybe_bytecode::VM::new();

    // 2. Register JS-specific host functions (console.log, coercion, etc.)
    let host = vybe_compiler_js::setup_js_runtime(&mut vm);

    // 3. Compile JS → bytecode using the host function table
    let chunks = match vybe_compiler_js::compile(&program, host) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("JS compile error: {e}");
            std::process::exit(1);
        }
    };

    // 4. Run
    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => {
            eprintln!("JS runtime error: {e}");
            std::process::exit(1);
        }
    }
}

/// Run a standalone .vb file as a console program.
fn run_vb_file(path: &Path, extra_args: &[String]) {
    let code = match fs::read_to_string(path) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("Error reading file: {e}");
            std::process::exit(1);
        }
    };

    let program = match parse_program(&code) {
        Ok(p) => p,
        Err(e) => {
            eprintln!("Parse error: {:?}", e);
            std::process::exit(1);
        }
    };

    let mut interp = Interpreter::new();
    interp.direct_console = true;
    interp.set_command_line_args(extra_args.to_vec());

    if let Err(e) = interp.run(&program) {
        eprintln!("Runtime error: {:?}", e);
        std::process::exit(1);
    }

    match interp.call_procedure(&vybe_parser_basic::ast::Identifier::new("main"), &[]) {
        Ok(_) => {}
        Err(vybe_runtime::RuntimeError::Exit(_)) => {}
        Err(vybe_runtime::RuntimeError::Return(_)) => {}
        Err(vybe_runtime::RuntimeError::Continue(_)) => {}
        Err(vybe_runtime::RuntimeError::UndefinedFunction(_)) => {} // no Main sub found
        Err(e) => {
            drain_console_effects(&mut interp);
            eprintln!("Runtime error: {:?}", e);
            std::process::exit(1);
        }
    }
    drain_console_effects(&mut interp);
}

/// Run a .vbp / .vbproj project.
fn run_project(path: &Path, extra_args: &[String]) {
    let project = match vybe_project::load_project_auto(path) {
        Ok(p) => p,
        Err(e) => {
            eprintln!("Error loading project: {e}");
            std::process::exit(1);
        }
    };

    let has_forms = !project.forms.is_empty();
    let mut starts_with_main = project.starts_with_main();

    // Fallback: if startup_object is None, scan code for Sub Main
    if !starts_with_main && !has_forms {
        for cf in &project.code_files {
            if cf.code.to_uppercase().contains("SUB MAIN") {
                starts_with_main = true;
                break;
            }
        }
    }

    if has_forms {
        // Has forms → launch the GUI (handles Sub Main inside FormRunner too)
        run_form_project(project);
    } else if starts_with_main {
        // Pure console project
        run_console_project(&project, extra_args);
    } else {
        eprintln!("Error: project has no forms and no Sub Main entry point");
        std::process::exit(1);
    }
}

/// Run a console-only project (Sub Main, no forms).
fn run_console_project(project: &Project, extra_args: &[String]) {
    let mut interp = Interpreter::new();
    interp.direct_console = true;
    interp.set_command_line_args(extra_args.to_vec());

    let entries = collect_resource_entries(project);
    interp.register_resource_entries(entries);

    for code_file in &project.code_files {
        match parse_program(&code_file.code) {
            Ok(program) => {
                if let Err(e) = interp.load_code_file(&program) {
                    eprintln!("Runtime error loading '{}': {:?}", code_file.name, e);
                }
            }
            Err(e) => {
                eprintln!("Parse error in '{}': {:?}", code_file.name, e);
            }
        }
    }

    match interp.call_procedure(&vybe_parser_basic::ast::Identifier::new("main"), &[]) {
        Ok(_) => {}
        Err(e) => {
            eprintln!("Sub Main error: {:?}", e);
            std::process::exit(1);
        }
    }

    drain_console_effects(&mut interp);
}

/// Launch a Dioxus desktop window showing the form runtime.
/// Uses the shared FormRunner – the exact same renderer the editor uses.
fn run_form_project(project: Project) {
    let title = project
        .get_startup_form()
        .map(|f| {
            if f.form.text.is_empty() {
                f.form.name.clone()
            } else {
                f.form.text.clone()
            }
        })
        .unwrap_or_else(|| project.name.clone());

    LAUNCH_PROJECT.with(|cell| *cell.borrow_mut() = Some(project));
    LAUNCH_TITLE.with(|cell| *cell.borrow_mut() = title.clone());

    let config = Config::new()
        .with_resource_directory(PathBuf::from("."))
        .with_window(
            WindowBuilder::new()
                .with_title(&title)
                .with_resizable(true),
        );

    LaunchBuilder::desktop()
        .with_cfg(config)
        .launch(ShellApp);
}

/// Top-level Dioxus component for the standalone shell.
#[component]
fn ShellApp() -> Element {
    let project = LAUNCH_PROJECT
        .with(|cell| cell.borrow().clone())
        .expect("LAUNCH_PROJECT must be set before launching");

    use_context_provider(|| RuntimeProject {
        project: Signal::new(Some(project)),
        finished: Signal::new(false),
    });

    rsx! { FormRunner {} }
}

/// Drain console side-effects from the interpreter and print them to stdout.
fn drain_console_effects(interp: &mut Interpreter) {
    while let Some(effect) = interp.side_effects.pop_front() {
        match effect {
            RuntimeSideEffect::ConsoleOutput(msg) => {
                print!("{msg}");
                use std::io::Write;
                let _ = std::io::stdout().flush();
            }
            RuntimeSideEffect::InputBox { .. } => {}
    
            RuntimeSideEffect::ConsoleClear => {}
            RuntimeSideEffect::MsgBox(msg) => println!("[MsgBox] {msg}"),
            RuntimeSideEffect::PropertyChange { .. } => {}
            RuntimeSideEffect::DataSourceChanged { .. } => {}
            RuntimeSideEffect::BindingPositionChanged { .. } => {}
            RuntimeSideEffect::FormClose { .. } => {}
            RuntimeSideEffect::FormShowDialog { .. } => {}
            RuntimeSideEffect::AddControl { .. } => {}
            RuntimeSideEffect::RunApplication { .. } => {}
            RuntimeSideEffect::Repaint { .. } => {}
        }
    }
}

/// Collect all resource entries from the project (resource_files + form-level resources)
/// into a flat Vec of ResourceEntry for the runtime.
pub fn collect_resource_entries(project: &Project) -> Vec<ResourceEntry> {
    let mut entries = Vec::new();

    // Project-level resource files
    for mgr in &project.resource_files {
        for item in &mgr.resources {
            let rt = format!("{:?}", item.resource_type).to_lowercase();
            entries.push(ResourceEntry {
                name: item.name.clone(),
                value: item.value.clone(),
                resource_type: rt,
                file_path: item.file_name.clone(),
            });
        }
    }

    // Legacy: also include old single resources field (backward compat)
    for item in &project.resources.resources {
        let rt = format!("{:?}", item.resource_type).to_lowercase();
        // Avoid duplicates (if already in resource_files)
        if !entries.iter().any(|e| e.name == item.name) {
            entries.push(ResourceEntry {
                name: item.name.clone(),
                value: item.value.clone(),
                resource_type: rt,
                file_path: item.file_name.clone(),
            });
        }
    }

    // Form-level resources
    for form_mod in &project.forms {
        for item in &form_mod.resources.resources {
            let rt = format!("{:?}", item.resource_type).to_lowercase();
            // Prefix form resources with form name to avoid collisions
            let key = format!("{}_{}", form_mod.form.name, item.name);
            entries.push(ResourceEntry {
                name: key,
                value: item.value.clone(),
                resource_type: rt,
                file_path: item.file_name.clone(),
            });
        }
    }

    entries
}
