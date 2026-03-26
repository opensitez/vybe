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
/// Supports multi-file modules (import/export).
/// If the JS program emits gui.runApplication(), a Dioxus window is launched.
fn run_js_file(path: &Path) {
    // 1. Create VM + shared side effect queue
    let mut vm = vybe_bytecode::VM::new();
    let queue = std::rc::Rc::new(std::cell::RefCell::new(vybe_host::SideEffectQueue::new()));

    // 2. Register all host modules (vybe:* + js:coerce)
    vybe_host::register_all_with_gui(&mut vm, queue.clone());
    vybe_compiler_js::register_js_coercion(&mut vm);

    // 3. Load, resolve imports, and compile all modules
    let chunks = match vybe_compiler_js::load_and_compile(path) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("JS error: {e}");
            std::process::exit(1);
        }
    };

    // 4. Run the top-level JS code (sets up form, controls, etc.)
    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => {
            eprintln!("JS runtime error: {e}");
            std::process::exit(1);
        }
    }

    // 5. Check the side effect queue — did the JS request a GUI?
    let effects = queue.borrow_mut().drain();
    let mut form_name = None;
    let mut form_title = String::new();
    let mut form = vybe_forms::Form::new("JSForm");
    form.width = 800;
    form.height = 600;

    for effect in &effects {
        match effect {
            vybe_host::SideEffect::RunApplication { form_name: name } => {
                form_name = Some(name.clone());
            }
            vybe_host::SideEffect::AddControl {
                control_name, control_type, left, top, width, height, ..
            } => {
                let ct = vybe_forms::control::ControlType::from_name(control_type)
                    .unwrap_or(vybe_forms::control::ControlType::Label);
                let mut ctrl = vybe_forms::Control::new(ct, control_name.clone(), *left, *top);
                ctrl.bounds = vybe_forms::Bounds::new(*left, *top, *width, *height);
                form.controls.push(ctrl);
            }
            vybe_host::SideEffect::PropertyChange { object, property, value } => {
                let val_str = value.as_string();
                // Form-level property
                if Some(object.clone()) == form_name || object == "JSForm"
                    || form_name.as_ref().is_some_and(|n| n == object)
                {
                    match property.as_str() {
                        "Text" => { form.text = val_str.clone(); form_title = val_str; }
                        "Width" => { form.width = val_str.parse().unwrap_or(800); }
                        "Height" => { form.height = val_str.parse().unwrap_or(600); }
                        _ => {}
                    }
                } else {
                    // Control property
                    if let Some(ctrl) = form.controls.iter_mut().find(|c| c.name == *object) {
                        ctrl.properties.set(property.clone(), val_str);
                    }
                }
            }
            vybe_host::SideEffect::ConsoleOutput(msg) => {
                print!("{msg}");
            }
            vybe_host::SideEffect::MsgBox { text, title } => {
                println!("[MsgBox] {}: {}", title, text);
            }
            _ => {}
        }
    }

    // 6. If RunApplication was called, launch the JS form with event support
    if let Some(name) = form_name {
        if form_title.is_empty() {
            form_title = name.clone();
        }
        form.name = name;

        // Pass data to the Dioxus component via thread-locals
        JS_LAUNCH_FORM.with(|cell| *cell.borrow_mut() = Some(form));
        JS_LAUNCH_TITLE.with(|cell| *cell.borrow_mut() = form_title.clone());
        JS_LAUNCH_VM.with(|cell| *cell.borrow_mut() = Some(vm));
        JS_LAUNCH_QUEUE.with(|cell| *cell.borrow_mut() = Some(queue));

        let config = Config::new()
            .with_resource_directory(PathBuf::from("."))
            .with_window(
                WindowBuilder::new()
                    .with_title(&form_title)
                    .with_resizable(true),
            );

        LaunchBuilder::desktop()
            .with_cfg(config)
            .launch(JsFormApp);
    }
}

// ---------------------------------------------------------------------------
// JS Form runner — Dioxus component with event dispatch to bytecode VM
// ---------------------------------------------------------------------------

thread_local! {
    static JS_LAUNCH_FORM: std::cell::RefCell<Option<vybe_forms::Form>> = std::cell::RefCell::new(None);
    static JS_LAUNCH_TITLE: std::cell::RefCell<String> = std::cell::RefCell::new(String::new());
    static JS_LAUNCH_VM: std::cell::RefCell<Option<vybe_bytecode::VM>> = std::cell::RefCell::new(None);
    static JS_LAUNCH_QUEUE: std::cell::RefCell<Option<std::rc::Rc<std::cell::RefCell<vybe_host::SideEffectQueue>>>> = std::cell::RefCell::new(None);
}

#[component]
fn JsFormApp() -> Element {
    // use_hook runs only on first render — safe to take() from thread-locals
    let (initial_form, vm_cell, queue_cell) = use_hook(|| {
        let form = JS_LAUNCH_FORM.with(|c| c.borrow_mut().take()).expect("JS_LAUNCH_FORM not set");
        let vm = JS_LAUNCH_VM.with(|c| c.borrow_mut().take()).expect("JS_LAUNCH_VM not set");
        let queue = JS_LAUNCH_QUEUE.with(|c| c.borrow_mut().take()).expect("JS_LAUNCH_QUEUE not set");
        (
            form,
            std::rc::Rc::new(std::cell::RefCell::new(vm)),
            queue,
        )
    });

    let form_width = initial_form.width;
    let form_height = initial_form.height;

    let runtime_form = use_signal(|| initial_form.clone());
    let vm_cell = vm_cell.clone();
    let queue_cell = queue_cell.clone();

    // Event handler: fires when a control is clicked
    let handle_event = {
        let vm_cell = vm_cell.clone();
        let queue_cell = queue_cell.clone();
        let mut runtime_form = runtime_form.clone();
        move |control_name: String, event_name: String| {
            let callback = {
                let q = queue_cell.borrow();
                q.get_event_handler(&control_name, &event_name).cloned()
            };
            if let Some(cb) = callback {
                let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                    let mut vm = vm_cell.borrow_mut();
                    vm.invoke(&cb, &[])
                }));
                match result {
                    Ok(Ok(_)) => {}
                    Ok(Err(e)) => eprintln!("Event handler error: {e}"),
                    Err(panic) => {
                        let msg = if let Some(s) = panic.downcast_ref::<&str>() {
                            s.to_string()
                        } else if let Some(s) = panic.downcast_ref::<String>() {
                            s.clone()
                        } else {
                            "unknown panic".to_string()
                        };
                        eprintln!("Event handler panic: {msg}");
                    }
                }

                // Process any new side effects from the callback
                let new_effects = queue_cell.borrow_mut().drain();
                let mut form = runtime_form.write();
                for effect in new_effects {
                    match effect {
                        vybe_host::SideEffect::PropertyChange { object, property, value } => {
                            let val_str = value.as_string();
                            if object == form.name {
                                match property.as_str() {
                                    "Text" => form.text = val_str,
                                    _ => {}
                                }
                            } else if let Some(ctrl) = form.controls.iter_mut().find(|c| c.name == object) {
                                ctrl.properties.set(property, val_str);
                            }
                        }
                        vybe_host::SideEffect::ConsoleOutput(msg) => {
                            print!("{msg}");
                        }
                        vybe_host::SideEffect::MsgBox { text, title } => {
                            println!("[MsgBox] {}: {}", title, text);
                        }
                        _ => {}
                    }
                }
            }
        }
    };

    // Collect control data — use shared control definitions from vybe_host
    let controls: Vec<(String, String, i32, i32, i32, i32, String, String)> = {
        let f = runtime_form.read();
        f.controls.iter().map(|ctrl| {
            let type_name = format!("{:?}", ctrl.control_type);
            let def = vybe_host::get_def(&type_name);
            // Collect properties for CSS generation
            let mut props = std::collections::HashMap::new();
            for (k, v) in ctrl.properties.iter() {
                if let Some(s) = v.as_string() {
                    props.insert(k.clone(), s.to_string());
                }
            }
            let css = (def.css_fn)(&props);
            (
                ctrl.name.clone(),
                type_name,
                ctrl.bounds.x, ctrl.bounds.y, ctrl.bounds.width, ctrl.bounds.height,
                ctrl.properties.get_string("Text").unwrap_or_default().to_string(),
                css,
            )
        }).collect()
    };

    rsx! {
        div {
            style: "width: {form_width}px; height: {form_height}px; position: relative; background: #f0f0f0; font-family: 'Segoe UI', sans-serif; font-size: 13px;",
            {controls.iter().map(|(ctrl_name, type_name, x, y, w, h, text, ctrl_css)| {
                let click_name = ctrl_name.clone();
                let mut handle = handle_event.clone();
                let def = vybe_host::get_def(type_name);

                let pos_style = format!(
                    "position: absolute; left: {}px; top: {}px; width: {}px; height: {}px; {}",
                    x, y, w, h, ctrl_css
                );

                // Generic rendering based on control definition tag
                match def.tag {
                    "button" => rsx! {
                        button {
                            key: "{ctrl_name}",
                            style: "{pos_style}",
                            onclick: move |_| handle(click_name.clone(), "Click".into()),
                            "{text}"
                        }
                    },
                    "input" => {
                        let input_type = def.input_type.unwrap_or("text");
                        if def.inner_tag == Some("input") {
                            // Checkbox/Radio: label wrapping input
                            rsx! {
                                label {
                                    key: "{ctrl_name}",
                                    style: "{pos_style}",
                                    onclick: move |_| handle(click_name.clone(), "Click".into()),
                                    input { r#type: "{input_type}" }
                                    "{text}"
                                }
                            }
                        } else {
                            rsx! {
                                input {
                                    key: "{ctrl_name}",
                                    style: "{pos_style}",
                                    r#type: "{input_type}",
                                    value: "{text}",
                                }
                            }
                        }
                    },
                    "select" => rsx! {
                        select {
                            key: "{ctrl_name}",
                            style: "{pos_style}",
                            onchange: move |_| handle(click_name.clone(), "SelectedIndexChanged".into()),
                        }
                    },
                    "progress" => rsx! {
                        progress {
                            key: "{ctrl_name}",
                            style: "{pos_style}",
                            value: "{text}",
                            max: "100",
                        }
                    },
                    "table" => rsx! {
                        div {
                            key: "{ctrl_name}",
                            style: "{pos_style}",
                            "[DataGrid]"
                        }
                    },
                    "nav" | "iframe" | "img" => rsx! {
                        div {
                            key: "{ctrl_name}",
                            style: "{pos_style}",
                            "{text}"
                        }
                    },
                    // div, a, label — all rendered as div with appropriate styles
                    _ => rsx! {
                        div {
                            key: "{ctrl_name}",
                            style: "{pos_style}",
                            onclick: move |_| handle(click_name.clone(), "Click".into()),
                            "{text}"
                        }
                    },
                }
            })}
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
