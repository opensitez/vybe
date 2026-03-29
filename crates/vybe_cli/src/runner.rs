use std::fs;
use std::path::Path;
use std::rc::Rc;
use std::cell::RefCell;

use vybe_parser_basic::parse_program;
use vybe_project::Project;

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
            vybe_host::SideEffect::RunApplication { form_name: name, .. } => {
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

        // Launch using egui renderer (replaces Dioxus UI in this crate).
        crate::egui_form::launch_egui_form(vm, queue, &form, &form_title);
    }
}

// ---------------------------------------------------------------------------
// Launch a form from a project's Form struct (already parsed from designer)
// ---------------------------------------------------------------------------

/// Wire event handlers by finding compiled chunks that match method names with Handles clauses.
fn wire_handles_from_chunks(
    program: &vybe_parser_basic::ast::Program,
    vm: &mut vybe_bytecode::VM,
    queue: &Rc<RefCell<vybe_host::SideEffectQueue>>,
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
                    let func_val = vybe_bytecode::Value::Object(Rc::new(RefCell::new(obj)));

                    for handle in handle_list {
                        let parts: Vec<&str> = handle.splitn(2, '.').collect();
                        if parts.len() == 2 {
                            let ctrl = parts[0].to_lowercase();
                            let event = parts[1].to_string();
                            queue.borrow_mut().register_event(&ctrl, &event, func_val.clone());
                        }
                    }
                }
            }
        }
    }
}

/// Scan parsed AST for Handles clauses and register event handlers with the side effect queue.
/// Looks up the compiled function in VM globals and wires it to the control.event.
fn wire_handles_from_ast(
    program: &vybe_parser_basic::ast::Program,
    vm: &vybe_bytecode::VM,
    queue: &std::rc::Rc<std::cell::RefCell<vybe_host::SideEffectQueue>>,
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
                        let o = class_obj.borrow();
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
                            let o = class_obj.borrow();
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
                            queue.borrow_mut().register_event(&ctrl, &event, func.clone());
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
    queue: std::rc::Rc<std::cell::RefCell<vybe_host::SideEffectQueue>>,
) {
    // Launch using egui renderer (replaces Dioxus UI in this crate).
    let title = if form.text.is_empty() { form.name.clone() } else { form.text.clone() };
    crate::egui_form::launch_egui_form(vm, queue, &form, &title);
}

// ---------------------------------------------------------------------------
// Public API: launch a bytecode VM form from side effects
// ---------------------------------------------------------------------------

/// Build a form from side effects and launch a Dioxus window.
/// Called by both `run_js_file` and `vybec` after compiling and running bytecode.
/// If no RunApplication side effect was emitted, just prints console output.
pub fn launch_vm_form(
    mut vm: vybe_bytecode::VM,
    queue: std::rc::Rc<std::cell::RefCell<vybe_host::SideEffectQueue>>,
    initial_form: Option<vybe_forms::Form>,
) {
    let effects = queue.borrow_mut().drain();

    // Process side effects for form object reference and console output
    for effect in &effects {
        match effect {
            vybe_host::SideEffect::RunApplication { form_object, .. } => {
                if let Some(obj) = form_object {
                    vm.globals.insert("__f".into(), obj.clone());
                }
            }
            vybe_host::SideEffect::ConsoleOutput(msg) => print!("{msg}"),
            vybe_host::SideEffect::MsgBox { text, title } => println!("[MsgBox] {title}: {text}"),
            _ => {}
        }
    }

    // Build the form — use the parsed model if available, otherwise from side effects
    let (form, form_title) = if let Some(mut f) = initial_form {
        // Designer form — apply any runtime property changes from side effects
        let fname = f.name.to_lowercase();
        for effect in &effects {
            if let vybe_host::SideEffect::PropertyChange { object, property, value } = effect {
                let val_str = value.as_string();
                if object.to_lowercase() == fname {
                    match property.as_str() {
                        "Text" | "Caption" => { f.text = val_str; }
                        _ => {}
                    }
                } else if let Some(ctrl) = f.controls.iter_mut().find(|c| c.name.eq_ignore_ascii_case(object)) {
                    ctrl.properties.set(property.clone(), val_str);
                }
            }
        }
        let title = if f.text.is_empty() { f.name.clone() } else { f.text.clone() };
        (f, title)
    } else {
        // Programmatic form — build entirely from side effects
        let mut form = vybe_forms::Form::new("VMForm");
        form.width = 800;
        form.height = 600;
        let mut run_form_name = None;
        let mut form_title = String::new();
        for effect in &effects {
            match effect {
                vybe_host::SideEffect::RunApplication { form_name: name, .. } => {
                    run_form_name = Some(name.clone());
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
                    if Some(object.clone()) == run_form_name || object == "VMForm" {
                        match property.as_str() {
                            "Text" | "Caption" => { form.text = val_str.clone(); form_title = val_str; }
                            "Width" => { form.width = val_str.parse().unwrap_or(800); }
                            "Height" => { form.height = val_str.parse().unwrap_or(600); }
                            _ => {}
                        }
                    } else if let Some(ctrl) = form.controls.iter_mut().find(|c| c.name == *object) {
                        ctrl.properties.set(property.clone(), val_str);
                    }
                }
                _ => {}
            }
        }
        if let Some(name) = run_form_name {
            if form_title.is_empty() { form_title = name.clone(); }
            form.name = name;
        }
        (form, form_title)
    };

    eprintln!("[LAUNCH] title={:?} controls={} width={} height={}", form_title, form.controls.len(), form.width, form.height);
    if !form_title.is_empty() || !form.controls.is_empty() {
        // Use tiny-skia renderer — no webview
        crate::skia_form::launch_skia_form(vm, queue, &form, &form_title);
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
    let queue = std::rc::Rc::new(std::cell::RefCell::new(vybe_host::SideEffectQueue::new()));
    vybe_host::register_all_with_gui(&mut vm, queue.clone());
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

    launch_vm_form(vm, queue, None);
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
    let queue = std::rc::Rc::new(std::cell::RefCell::new(vybe_host::SideEffectQueue::new()));
    vybe_host::register_all_with_gui(&mut vm, queue.clone());
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

    launch_vm_form(vm, queue, None);
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
        let form = &fm.form;
        let mut designer = format!("Partial Class {}\n", form.name);
        designer.push_str("    Inherits System.Windows.Forms.Form\n\n");
        for ctrl in &form.controls {
            designer.push_str(&format!("    Friend WithEvents {} As {}\n", ctrl.name, control_type_name(ctrl)));
        }
        let user_code = fm.get_user_code();
        let has_ctor = user_code.to_uppercase().contains("SUB NEW");
        if !has_ctor {
            designer.push_str("\n    Public Sub New()\n");
            designer.push_str("        InitializeComponent()\n");
            designer.push_str("    End Sub\n");
        }
        designer.push_str("\n    Private Sub InitializeComponent()\n");
        for ctrl in &form.controls {
            designer.push_str(&format!("        Me.{} = New {}()\n", ctrl.name, control_type_name(ctrl)));
        }
        for ctrl in &form.controls {
            let is_non_visual = ctrl.control_type.is_non_visual();
            designer.push_str(&format!("        Me.{}.Name = \"{}\"\n", ctrl.name, ctrl.name));
            if let Some(text) = ctrl.properties.get_string("Text") {
                designer.push_str(&format!("        Me.{}.Text = \"{}\"\n", ctrl.name, text));
            }
            for (key, val) in ctrl.properties.iter() {
                if let Some(s) = val.as_string() {
                    let k = key.as_str();
                    if matches!(k, "Name" | "Text" | "Location" | "Size" | "TabIndex"
                        | "Enabled" | "Visible" | "BackColor" | "ForeColor" | "Font") {
                        continue;
                    }
                    if k.starts_with("DataBindings.") { continue; }
                    if !s.is_empty() {
                        if k == "DataSource" {
                            designer.push_str(&format!("        Me.{}.DataSource = Me.{}\n", ctrl.name, s));
                        } else if k == "BindingSource" {
                            designer.push_str(&format!("        Me.{}.BindingSource = Me.{}\n", ctrl.name, s));
                        } else if k.starts_with("DataBinding:") {
                            let parts: Vec<&str> = k.splitn(2, ':').collect();
                            if parts.len() == 2 {
                                let prop = parts[1];
                                let binding_parts: Vec<&str> = s.splitn(2, |c| c == '|' || c == '.').collect();
                                if binding_parts.len() == 2 {
                                    designer.push_str(&format!(
                                        "        Me.{}.DataBindings.Add(\"{}\", Me.{}, \"{}\")\n",
                                        ctrl.name, prop, binding_parts[0], binding_parts[1]
                                    ));
                                }
                            }
                        } else {
                            designer.push_str(&format!("        Me.{}.{} = \"{}\"\n", ctrl.name, k, s));
                        }
                    }
                }
            }
            if !is_non_visual {
                designer.push_str(&format!("        Me.{}.Location = New Point({}, {})\n", ctrl.name, ctrl.bounds.x, ctrl.bounds.y));
                designer.push_str(&format!("        Me.{}.Size = New Size({}, {})\n", ctrl.name, ctrl.bounds.width, ctrl.bounds.height));
                designer.push_str(&format!("        Me.Controls.Add(Me.{})\n", ctrl.name));
            }
        }
        designer.push_str(&format!("        Me.Name = \"{}\"\n", form.name));
        if !form.text.is_empty() {
            designer.push_str(&format!("        Me.Text = \"{}\"\n", form.text));
        }
        designer.push_str("    End Sub\n");
        designer.push_str("End Class\n");
        all_code.push_str(&designer);
        all_code.push('\n');
        all_code.push_str(&fm.get_user_code());
        all_code.push('\n');
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
    let queue = std::rc::Rc::new(std::cell::RefCell::new(vybe_host::SideEffectQueue::new()));
    vybe_host::register_all_with_gui(&mut vm, queue.clone());
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

    launch_vm_form(vm, queue, startup_form);
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

    // Compile ALL code — modules, classes, forms — into one program.
    // For forms, generate simplified InitializeComponent from the parsed form model
    // instead of re-including the raw designer code (which uses fully-qualified names
    // like System.Windows.Forms.FormStartPosition.CenterScreen that our compiler can't handle).
    let mut all_code = String::new();
    for cf in &project.code_files {
        all_code.push_str(&cf.code);
        all_code.push('\n');
    }
    for fm in &project.forms {
        // Generate simple designer code from the parsed form model
        let form = &fm.form;
        let mut designer = format!("Partial Class {}\n", form.name);
        designer.push_str("    Inherits System.Windows.Forms.Form\n\n");
        // Field declarations
        for ctrl in &form.controls {
            designer.push_str(&format!("    Friend WithEvents {} As {}\n", ctrl.name, control_type_name(ctrl)));
        }
        // Auto-inject constructor that calls InitializeComponent if user code doesn't have one
        let user_code = fm.get_user_code();
        let has_ctor = user_code.to_uppercase().contains("SUB NEW");
        if !has_ctor {
            designer.push_str("\n    Public Sub New()\n");
            designer.push_str("        InitializeComponent()\n");
            designer.push_str("    End Sub\n");
        }
        designer.push_str("\n    Private Sub InitializeComponent()\n");
        // Create controls using bare names (not fully-qualified)
        for ctrl in &form.controls {
            designer.push_str(&format!("        Me.{} = New {}()\n", ctrl.name, control_type_name(ctrl)));
        }
        // Set properties on all controls, but only Controls.Add for visual ones
        for ctrl in &form.controls {
            let is_non_visual = ctrl.control_type.is_non_visual();
            designer.push_str(&format!("        Me.{}.Name = \"{}\"\n", ctrl.name, ctrl.name));
            if let Some(text) = ctrl.properties.get_string("Text") {
                designer.push_str(&format!("        Me.{}.Text = \"{}\"\n", ctrl.name, text));
            }
            // Emit all string properties from the parsed form model
            // (ConnectionString, DataSource, DataMember, DbType, etc.)
            for (key, val) in ctrl.properties.iter() {
                if let Some(s) = val.as_string() {
                    let k = key.as_str();
                    // Skip already emitted or layout properties
                    if matches!(k, "Name" | "Text" | "Location" | "Size" | "TabIndex"
                        | "Enabled" | "Visible" | "BackColor" | "ForeColor" | "Font") {
                        continue;
                    }
                    // Skip DataBindings.* properties — handled by skia_form data binding system
                    if k.starts_with("DataBindings.") {
                        continue;
                    }
                    if !s.is_empty() {
                        if k == "DataSource" {
                            // DataSource is a reference to another control: Me.bs1.DataSource = Me.da1
                            designer.push_str(&format!("        Me.{}.DataSource = Me.{}\n", ctrl.name, s));
                        } else if k == "BindingSource" {
                            designer.push_str(&format!("        Me.{}.BindingSource = Me.{}\n", ctrl.name, s));
                        } else if k.starts_with("DataBinding:") {
                            // DataBindings.Add("Text", bs1, "ColumnName")
                            let parts: Vec<&str> = k.splitn(2, ':').collect();
                            if parts.len() == 2 {
                                let prop = parts[1];
                                // s = "bs1|ColumnName" or "bs1.ColumnName"
                                let binding_parts: Vec<&str> = s.splitn(2, |c| c == '|' || c == '.').collect();
                                if binding_parts.len() == 2 {
                                    designer.push_str(&format!(
                                        "        Me.{}.DataBindings.Add(\"{}\", Me.{}, \"{}\")\n",
                                        ctrl.name, prop, binding_parts[0], binding_parts[1]
                                    ));
                                }
                            }
                        } else {
                            // Generic property: Me.ctrl.Prop = "value"
                            designer.push_str(&format!("        Me.{}.{} = \"{}\"\n", ctrl.name, k, s));
                        }
                    }
                }
            }
            if !is_non_visual {
                designer.push_str(&format!(
                    "        Me.{}.Location = New Point({}, {})\n",
                    ctrl.name, ctrl.bounds.x, ctrl.bounds.y
                ));
                designer.push_str(&format!(
                    "        Me.{}.Size = New Size({}, {})\n",
                    ctrl.name, ctrl.bounds.width, ctrl.bounds.height
                ));
                designer.push_str(&format!("        Me.Controls.Add(Me.{})\n", ctrl.name));
            }
        }
        // Form properties
        designer.push_str(&format!("        Me.Name = \"{}\"\n", form.name));
        if !form.text.is_empty() {
            designer.push_str(&format!("        Me.Text = \"{}\"\n", form.text));
        }
        designer.push_str("    End Sub\n");
        designer.push_str("End Class\n");
        eprintln!("[GENERATED-DESIGNER]\n{}", designer);
        all_code.push_str(&designer);
        all_code.push('\n');
        // User code (event handlers etc.)
        all_code.push_str(&fm.get_user_code());
        all_code.push('\n');
    }

    // Determine startup mode
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

    let is_sub_main = project.starts_with_main()
        || all_code.to_uppercase().contains("SUB MAIN");

    // For form projects, we compile the class then instantiate it with
    // `New FormName()` which calls InitializeComponent and wires events.

    if startup_form.is_none() && !is_sub_main {
        eprintln!("Error: project has no forms and no Sub Main entry point");
        std::process::exit(1);
    }

    if all_code.trim().is_empty() && startup_form.is_none() {
        eprintln!("Error: no code to run");
        std::process::exit(1);
    }

    // For form projects, instantiate the class and run the app.
    // InitializeComponent creates controls (emits AddControl side effects),
    // sets properties, and wires Handles events — all via the VM.
    if let Some(ref form) = startup_form {
        all_code.push_str(&format!(
            "\nDim __f As New {}()\nApplication.Run(__f)\n",
            form.name
        ));
    }

    // Set up VM + compile ALL code (class defs + entry point together)
    let mut vm = vybe_bytecode::VM::new();
    let queue = std::rc::Rc::new(std::cell::RefCell::new(vybe_host::SideEffectQueue::new()));
    vybe_host::register_all_with_gui(&mut vm, queue.clone());
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
    // The form model IS the source of truth for rendering.
    // Side effects are only used for runtime property changes during events.
    if let Some(form) = startup_form {
        launch_vm_form(vm, queue, Some(form));
    } else {
        launch_vm_form(vm, queue, None);
    }
}
