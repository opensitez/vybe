use std::fs;
use std::path::Path;
use std::sync::{Arc, Mutex};

use vybe_parser_basic::parse_program;

// ---------------------------------------------------------------------------
// Thread-local used to pass the Project into the named Dioxus App component.
// (Dioxus `launch()` requires a plain fn-pointer, so we can't use a closure.)
// ---------------------------------------------------------------------------
// Dioxus removed from this crate; use tiny-skia renderer instead.

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
        "cs" => run_cs_file(path),
        _ => {
            eprintln!(
                "Error: unsupported file type '.{}'. Expected .vb, .vbp, .vbproj, .js, or .cs",
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
    let mut vm = vybe_bytecode::VM::new();

    let gui = vybe_host::register_all_with_gui(&mut vm);
    vybe_compiler_js::register_js_coercion(&mut vm);

    let chunks = match vybe_compiler_js::load_and_compile(path) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("JS error: {e}");
            std::process::exit(1);
        }
    };

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => {
            eprintln!("JS runtime error: {e}");
            std::process::exit(1);
        }
    }

    // Drain any pending MsgBox dialogs (before GUI window opens)
    {
        let dialogs: Vec<(String, String)> = gui.lock().unwrap().pending_dialogs.drain(..).collect();
        for (text, title) in dialogs {
            println!("[MsgBox] {title}: {text}");
        }
    }

    // If runApplication was called, GuiState already has the real widgets
    if gui.lock().unwrap().should_run {
        crate::vybewidget_form::launch_gui(vm, gui);
    }
}

// ---------------------------------------------------------------------------
// Launch a form from a project's Form struct (already parsed from designer)
// ---------------------------------------------------------------------------

/// Wire event handlers by finding compiled chunks that match method names with Handles clauses.
#[allow(dead_code)]
fn wire_handles_from_chunks(
    program: &vybe_parser_basic::ast::Program,
    vm: &mut vybe_bytecode::VM,
    gui: &Arc<Mutex<vybe_host::GuiState>>,
) {
    use vybe_parser_basic::ast::*;
    for decl in &program.declarations {
        let methods = match decl {
            Declaration::Class(c) => &c.methods,
            _ => continue,
        };
        for method in methods {
            let (name, handles) = match method {
                MethodDecl::Sub(s) => (&s.name, &s.handles),
                MethodDecl::Function(f) => (&f.name, &f.handles),
            };
            if let Some(handle_list) = handles {
                let method_lower = name.as_str().to_lowercase();
                // Find the compiled chunk by name
                let chunk_idx = vm.chunks.iter().position(|c| c.name.to_lowercase() == method_lower);
                if let Some(idx) = chunk_idx {
                    // Create a closure value for this chunk
                    let func = vybe_bytecode::value::Function {
                        name: Some(method_lower.clone()),
                        arity: vm.chunks[idx].arity,
                        chunk_index: idx,
                        upvalues: Vec::new(),
                    };
                    let obj = vybe_bytecode::value::Object {
                        properties: std::collections::HashMap::new(),
                        kind: vybe_bytecode::value::ObjectKind::Function(func),
                        type_id: 0, fields: Vec::new(),
                    };
                    let func_val = vybe_bytecode::Value::Object(Arc::new(Mutex::new(obj)));

                    for handle in handle_list {
                        let parts: Vec<&str> = handle.splitn(2, '.').collect();
                        if parts.len() == 2 {
                            let ctrl = parts[0].to_lowercase();
                            let event = parts[1].to_string();
                            gui.lock().unwrap().register_event(&ctrl, &event, func_val.clone());
                        }
                    }
                }
            }
        }
    }
}

/// Scan parsed AST for Handles clauses and register event handlers with GuiState.
/// Looks up the compiled function in VM globals and wires it to the control.event.
#[allow(dead_code)]
fn wire_handles_from_ast(
    program: &vybe_parser_basic::ast::Program,
    vm: &vybe_bytecode::VM,
    gui: &std::sync::Arc<std::sync::Mutex<vybe_host::GuiState>>,
) {
    use vybe_parser_basic::ast::*;
    for decl in &program.declarations {
        let (class_name, methods) = match decl {
            Declaration::Class(c) => (c.name.as_str().to_lowercase(), &c.methods),
            _ => continue,
        };
        for method in methods {
            let (name, handles) = match method {
                MethodDecl::Sub(s) => (&s.name, &s.handles),
                MethodDecl::Function(f) => (&f.name, &f.handles),
            };
            if let Some(handle_list) = handles {
                let method_lower = name.as_str().to_lowercase();
                eprintln!("[wire] checking method={} handles={:?}", method_lower, handle_list);
                // Look up the method: first try as global, then on the class constructor
                let func_val = vm.globals.get(&method_lower).cloned().or_else(|| {
                    // Method is on the class — look up class global, get method property
                    if let Some(vybe_bytecode::Value::Object(class_obj)) = vm.globals.get(&class_name) {
                        let o = class_obj.lock().unwrap();
                        o.properties.get(&method_lower).cloned()
                    } else {
                        None
                    }
                });
                // Also check class properties
                if func_val.is_none() {
                    if let Some(class_val) = vm.globals.get(&class_name) {
                        eprintln!("[wire] class={} type={}", class_name, class_val.type_tag());
                        if let vybe_bytecode::Value::Object(class_obj) = class_val {
                            let o = class_obj.lock().unwrap();
                            eprintln!("[wire] class={} kind={:?} props={:?}", class_name, std::mem::discriminant(&o.kind), o.properties.keys().collect::<Vec<_>>());
                        }
                    } else {
                        eprintln!("[wire] class={} NOT FOUND in globals", class_name);
                    }
                }
                eprintln!("[wire] method={} func_found={}", method_lower, func_val.is_some());
                if let Some(func) = func_val {
                    for handle in handle_list {
                        let parts: Vec<&str> = handle.splitn(2, '.').collect();
                        if parts.len() == 2 {
                            let ctrl = if parts[0].eq_ignore_ascii_case("Me") {
                                class_name.clone()
                            } else {
                                parts[0].to_lowercase()
                            };
                            let event = parts[1].to_string();
                            gui.lock().unwrap().register_event(&ctrl, &event, func.clone());
                        }
                    }
                }
            }
        }
    }
}


pub fn launch_project_form(
    form: vybe_forms::Form,
    vm: vybe_bytecode::VM,
    gui: std::sync::Arc<std::sync::Mutex<vybe_host::GuiState>>,
) {
    crate::vybewidget_form::launch_vybewidget_form(vm, gui, &form);
}

// ---------------------------------------------------------------------------
// Public API: launch a bytecode VM form from side effects
// ---------------------------------------------------------------------------

/// Launch a form after VM has run.
/// For programmatic forms, GuiState already has the real widgets.
/// For designer forms, pass `initial_form` to build widgets from the model.
pub fn launch_vm_form(
    vm: vybe_bytecode::VM,
    gui: std::sync::Arc<std::sync::Mutex<vybe_host::GuiState>>,
    initial_form: Option<vybe_forms::Form>,
) {
    // Drain any pending MsgBox dialogs produced before the window opens
    {
        let dialogs: Vec<(String, String)> = gui.lock().unwrap().pending_dialogs.drain(..).collect();
        for (text, title) in dialogs {
            println!("[MsgBox] {title}: {text}");
        }
    }

    let should_launch = gui.lock().unwrap().should_run || initial_form.is_some();

    if should_launch {
        if let Some(form) = initial_form {
            // Designer form path — convert vybe_forms into real widgets
            crate::vybewidget_form::launch_vybewidget_form(vm, gui, &form);
        } else {
            // Programmatic form path — GuiState already has all widgets
            crate::vybewidget_form::launch_gui(vm, gui);
        }
    }
}

// Dioxus UI removed from this crate; JS form rendering uses the tiny-skia renderer.

/// Run a standalone .cs file via the bytecode VM.
fn run_cs_file(path: &Path) {
    let code = match fs::read_to_string(path) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("Error reading file: {e}");
            std::process::exit(1);
        }
    };

    let unit = match vybe_parser_csharp::parse(&code) {
        Ok(u) => u,
        Err(e) => {
            eprintln!("Parse error: {e}");
            std::process::exit(1);
        }
    };

    let mut vm = vybe_bytecode::VM::new();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    let chunks = match vybe_compiler_csharp::Compiler::new().compile(&unit) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("Compile error: {e}");
            std::process::exit(1);
        }
    };

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => {
            eprintln!("Runtime error: {e}");
            std::process::exit(1);
        }
    }

    launch_vm_form(vm, gui, None);
}

/// Run a standalone .vb file via the bytecode VM.
fn run_vb_file(path: &Path, _extra_args: &[String]) {
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

    let mut vm = vybe_bytecode::VM::new();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    let chunks = match vybe_compiler_vb::Compiler::new().compile(&program) {
        Ok(c) => c,
        Err(e) => {
            eprintln!("Compile error: {e}");
            std::process::exit(1);
        }
    };

    match vm.run(chunks) {
        Ok(_) => {}
        Err(e) => {
            eprintln!("Runtime error: {e}");
            std::process::exit(1);
        }
    }

    launch_vm_form(vm, gui, None);
}

/// Map internal ControlType debug names to host-recognized constructor names.
fn control_type_name(ctrl: &vybe_forms::Control) -> String {
    match format!("{:?}", ctrl.control_type).as_str() {
        "Frame" => "GroupBox".into(),
        "BindingSourceComponent" => "BindingSource".into(),
        "DataSetComponent" => "DataSet".into(),
        "DataTableComponent" => "DataTable".into(),
        "DataAdapterComponent" => "DataAdapter".into(),
        other => other.into(),
    }
}

/// Build the combined VB source code for a project (designer + user code + entry point).
/// Returns (all_code, startup_form) — startup_form is Some if this is a forms project.
pub fn build_project_code(project: &vybe_project::Project) -> (String, Option<vybe_forms::Form>) {
    let mut all_code = String::new();
    for cf in &project.code_files {
        all_code.push_str(&cf.code);
        all_code.push('\n');
    }
    for fm in &project.forms {
        let user_code = fm.get_user_code();
        let user_upper = user_code.to_uppercase();
        let has_init = user_upper.contains("INITIALIZECOMPONENT");

        if has_init {
            // User code already has InitializeComponent — use it directly.
            // The compiler handles fully-qualified .NET names via:
            // - compile_new_expr prefix stripping (New System.Windows.Forms.Panel → new_Panel)
            // - resolve_interface_call (System.Drawing.Color.FromArgb → vybe:gui/color.fromargb)
            // - runtime namespace objects (System.Windows.Forms.BorderStyle.FixedSingle → F64)
            all_code.push_str(user_code);
            all_code.push('\n');
        } else {
            // No user InitializeComponent — generate designer from the form model
            // using the full-featured codegen that handles colors, fonts, nesting, etc.
            let designer = vybe_forms::serialization::designer_codegen::generate_designer_code(&fm.form);
            all_code.push_str(&designer);
            all_code.push('\n');
            all_code.push_str(user_code);
            all_code.push('\n');
        }
    }

    let startup_form_name = match &project.startup_object {
        vybe_project::StartupObject::Form(name) => Some(name.clone()),
        vybe_project::StartupObject::None if !project.forms.is_empty() => {
            Some(project.forms.first().unwrap().form.name.clone())
        }
        _ => None,
    };

    let startup_form = startup_form_name.as_ref().and_then(|_| {
        project.get_startup_form().map(|fm| fm.form.clone())
            .or_else(|| project.forms.first().map(|fm| fm.form.clone()))
    });

    let is_sub_main = project.starts_with_main() || all_code.to_uppercase().contains("SUB MAIN");

    if let Some(ref form) = startup_form {
        if !is_sub_main {
            all_code.push_str(&format!("\nDim __f As New {}()\nApplication.Run(__f)\n", form.name));
        }
    }

    (all_code, startup_form)
}

/// Run an already-loaded project (used by vybe_editor).
/// Must be called on the main thread — opens a native window for form projects.
pub fn run_project_in_memory(project: &vybe_project::Project) {
    let (all_code, startup_form) = build_project_code(project);

    if all_code.trim().is_empty() {
        eprintln!("Error: no code to run");
        return;
    }

    let mut vm = vybe_bytecode::VM::new();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    match parse_program(&all_code) {
        Ok(program) => {
            match vybe_compiler_vb::Compiler::new().compile(&program) {
                Ok(chunks) => {
                    if let Err(e) = vm.run(chunks) {
                        let msg = format!("{}", e);
                        if !msg.starts_with("__") {
                            eprintln!("Runtime error: {e}");
                        }
                    }
                }
                Err(e) => { eprintln!("Compile error: {e}"); return; }
            }
        }
        Err(e) => { eprintln!("Parse error: {:?}", e); return; }
    }

    launch_vm_form(vm, gui, startup_form);
}

/// Run a .vbp / .vbproj project.
fn run_project(path: &Path, _extra_args: &[String]) {
    let project = match vybe_project::load_project_auto(path) {
        Ok(p) => p,
        Err(e) => {
            eprintln!("Error loading project: {e}");
            std::process::exit(1);
        }
    };

    let (all_code, startup_form) = build_project_code(&project);

    if startup_form.is_none() && !all_code.to_uppercase().contains("SUB MAIN") {
        eprintln!("Error: project has no forms and no Sub Main entry point");
        std::process::exit(1);
    }

    if all_code.trim().is_empty() && startup_form.is_none() {
        eprintln!("Error: no code to run");
        std::process::exit(1);
    }

    // Set up VM + compile ALL code (class defs + entry point together)
    let mut vm = vybe_bytecode::VM::new();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vybe_host::setup_namespaces(&mut vm);

    if !all_code.trim().is_empty() {
        match parse_program(&all_code) {
            Ok(program) => {
                match vybe_compiler_vb::Compiler::new().compile(&program) {
                    Ok(chunks) => {
                        if let Err(e) = vm.run(chunks) {
                            let msg = format!("{}", e);
                            if !msg.starts_with("__") {
                                eprintln!("Runtime error: {e}");
                            }
                        }
                    }
                    Err(e) => eprintln!("Compile error: {e}"),
                }
            }
            Err(e) => eprintln!("Parse error: {:?}", e),
        }
    }

    // Pass the parsed form model directly — no side effects needed for layout.
    // Side effects are only used for runtime property changes during events.
    launch_vm_form(vm, gui, startup_form);
}
