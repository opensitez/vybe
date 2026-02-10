use std::fs;
use std::path::{Path, PathBuf};
use std::collections::HashMap;

use dioxus::prelude::*;
use dioxus_desktop::{Config, WindowBuilder};

use irys_parser::parse_program;
use irys_project::Project;
use irys_runtime::{Interpreter, RuntimeSideEffect};

use crate::runtime_panel::RuntimeProject;
use crate::FormRunner;

// ---------------------------------------------------------------------------
// Thread-local used to pass the Project into the named Dioxus App component.
// (Dioxus `launch()` requires a plain fn-pointer, so we can't use a closure.)
// ---------------------------------------------------------------------------
thread_local! {
    static LAUNCH_PROJECT: std::cell::RefCell<Option<Project>> = std::cell::RefCell::new(None);
    static LAUNCH_TITLE: std::cell::RefCell<String> = std::cell::RefCell::new(String::new());
}

// ---------------------------------------------------------------------------
// Public entry point – the ONLY function the shell binary calls.
// ---------------------------------------------------------------------------

/// Run a Visual Basic file or project.
///
/// * `.vb`    → parse & run as console program
/// * `.vbp`   → load VB6 project, run as form or console
/// * `.vbproj` → load VB.NET project, run as form or console
pub fn run(path: &Path) {
    let ext = path
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default();

    match ext.as_str() {
        "vb" => run_vb_file(path),
        "vbp" | "vbproj" => run_project(path),
        _ => {
            eprintln!(
                "Error: unsupported file type '.{}'. Expected .vb, .vbp, or .vbproj",
                ext
            );
            std::process::exit(1);
        }
    }
}

// ---------------------------------------------------------------------------
// Internals
// ---------------------------------------------------------------------------

/// Run a standalone .vb file as a console program.
fn run_vb_file(path: &Path) {
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

    if let Err(e) = interp.run(&program) {
        eprintln!("Runtime error: {:?}", e);
        std::process::exit(1);
    }

    let _ = interp.call_procedure(&irys_parser::ast::Identifier::new("main"), &[]);
    drain_console_effects(&mut interp);
}

/// Run a .vbp / .vbproj project.
fn run_project(path: &Path) {
    let project = match irys_project::load_project_auto(path) {
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
        run_console_project(&project);
    } else {
        eprintln!("Error: project has no forms and no Sub Main entry point");
        std::process::exit(1);
    }
}

/// Run a console-only project (Sub Main, no forms).
fn run_console_project(project: &Project) {
    let mut interp = Interpreter::new();

    let mut res_map = HashMap::new();
    for item in &project.resources.resources {
        res_map.insert(item.name.clone(), item.value.clone());
    }
    interp.register_resources(res_map);

    for code_file in &project.code_files {
        if let Ok(program) = parse_program(&code_file.code) {
            let _ = interp.load_module(&code_file.name, &program);
        }
    }

    match interp.call_procedure(&irys_parser::ast::Identifier::new("main"), &[]) {
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
            if f.form.caption.is_empty() {
                f.form.name.clone()
            } else {
                f.form.caption.clone()
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
    });

    rsx! { FormRunner {} }
}

/// Drain console side-effects from the interpreter and print them to stdout.
fn drain_console_effects(interp: &mut Interpreter) {
    while let Some(effect) = interp.side_effects.pop_front() {
        match effect {
            RuntimeSideEffect::ConsoleOutput(msg) => println!("{msg}"),
            RuntimeSideEffect::ConsoleClear => {}
            RuntimeSideEffect::MsgBox(msg) => println!("[MsgBox] {msg}"),
            RuntimeSideEffect::PropertyChange { .. } => {}
        }
    }
}
