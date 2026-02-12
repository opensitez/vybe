use dioxus::prelude::*;
use irys_forms::{ControlType, Form, EventType};
use irys_project::Project;
use irys_runtime::{Interpreter, RuntimeSideEffect, Value, ObjectData, ConsoleMessage};
use std::collections::HashMap;
use std::cell::RefCell;
use std::rc::Rc;
use std::sync::mpsc;
use irys_parser::parse_program;

// ---------------------------------------------------------------------------
// Console color helpers
// ---------------------------------------------------------------------------

/// Map .NET ConsoleColor value (0-15) to a CSS color string.
fn console_color_to_css(color: i32) -> &'static str {
    match color {
        0  => "#0c0c0c",  // Black
        1  => "#0037da",  // DarkBlue
        2  => "#13a10e",  // DarkGreen
        3  => "#3a96dd",  // DarkCyan
        4  => "#c50f1f",  // DarkRed
        5  => "#881798",  // DarkMagenta
        6  => "#c19c00",  // DarkYellow
        7  => "#cccccc",  // Gray
        8  => "#767676",  // DarkGray
        9  => "#3b78ff",  // Blue
        10 => "#16c60c",  // Green
        11 => "#61d6d6",  // Cyan
        12 => "#e74856",  // Red
        13 => "#b4009e",  // Magenta
        14 => "#f9f1a5",  // Yellow
        15 => "#f2f2f2",  // White
        _  => "#cccccc",  // default Gray
    }
}

/// Escape HTML special characters.
fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
     .replace('<', "&lt;")
     .replace('>', "&gt;")
     .replace('"', "&quot;")
}

// ---------------------------------------------------------------------------
// Public context wrapper – both the editor and the CLI provide this via
// `use_context_provider`.  It is the single source-of-truth the FormRunner
// reads from, so there is exactly one code path.
// ---------------------------------------------------------------------------

/// Wrapper around a `Signal<Option<Project>>` that the FormRunner reads.
/// Both the IDE and the standalone shell provide this via `use_context_provider`.
#[derive(Clone, Copy)]
pub struct RuntimeProject {
    pub project: Signal<Option<Project>>,
    /// Set to `true` by the FormRunner when a console (Sub Main) project
    /// finishes executing.  The editor watches this to leave run-mode
    /// automatically; the standalone shell can ignore it.
    pub finished: Signal<bool>,
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn event_type_from_name(name: &str) -> Option<EventType> {
    match name.to_lowercase().as_str() {
        "click" => Some(EventType::Click),
        "dblclick" => Some(EventType::DblClick),
        "doubleclick" => Some(EventType::DoubleClick),
        "load" => Some(EventType::Load),
        "unload" => Some(EventType::Unload),
        "change" => Some(EventType::Change),
        "textchanged" => Some(EventType::TextChanged),
        "selectedindexchanged" => Some(EventType::SelectedIndexChanged),
        "checkedchanged" => Some(EventType::CheckedChanged),
        "valuechanged" => Some(EventType::ValueChanged),
        "keypress" => Some(EventType::KeyPress),
        "keydown" => Some(EventType::KeyDown),
        "keyup" => Some(EventType::KeyUp),
        "mouseclick" => Some(EventType::MouseClick),
        "mousedoubleclick" => Some(EventType::MouseDoubleClick),
        "mousedown" => Some(EventType::MouseDown),
        "mouseup" => Some(EventType::MouseUp),
        "mousemove" => Some(EventType::MouseMove),
        "mouseenter" => Some(EventType::MouseEnter),
        "mouseleave" => Some(EventType::MouseLeave),
        "mousewheel" => Some(EventType::MouseWheel),
        "gotfocus" => Some(EventType::GotFocus),
        "lostfocus" => Some(EventType::LostFocus),
        "enter" => Some(EventType::Enter),
        "leave" => Some(EventType::Leave),
        "validated" => Some(EventType::Validated),
        "validating" => Some(EventType::Validating),
        "resize" => Some(EventType::Resize),
        "paint" => Some(EventType::Paint),
        "formclosing" => Some(EventType::FormClosing),
        "formclosed" => Some(EventType::FormClosed),
        "shown" => Some(EventType::Shown),
        "activated" => Some(EventType::Activated),
        "deactivate" => Some(EventType::Deactivate),
        "tick" => Some(EventType::Tick),
        "elapsed" => Some(EventType::Elapsed),
        "scroll" => Some(EventType::Scroll),
        "selectedvaluechanged" => Some(EventType::SelectedValueChanged),
        "cellclick" => Some(EventType::CellClick),
        "celldoubleclick" => Some(EventType::CellDoubleClick),
        "cellvaluechanged" => Some(EventType::CellValueChanged),
        "selectionchanged" => Some(EventType::SelectionChanged),
        _ => None,
    }
}

/// Map Dioxus keyboard Key to Windows Forms Virtual Key code (VK_*).
fn dioxus_key_to_vk(key: &dioxus::prelude::Key) -> i32 {
    use dioxus::prelude::Key;
    match key {
        Key::Backspace => 8,
        Key::Tab => 9,
        Key::Enter => 13,
        Key::Shift => 16,
        Key::Control => 17,
        Key::Alt => 18,
        Key::Pause => 19,
        Key::CapsLock => 20,
        Key::Escape => 27,
        Key::PageUp => 33,
        Key::PageDown => 34,
        Key::End => 35,
        Key::Home => 36,
        Key::ArrowLeft => 37,
        Key::ArrowUp => 38,
        Key::ArrowRight => 39,
        Key::ArrowDown => 40,
        Key::Insert => 45,
        Key::Delete => 46,
        Key::Character(c) => {
            let ch = c.chars().next().unwrap_or('\0');
            match ch {
                '0'..='9' => ch as i32,              // 0x30–0x39
                'a'..='z' => (ch as i32) - 32,       // VK uses uppercase: 0x41–0x5A
                'A'..='Z' => ch as i32,
                ' ' => 32,
                '+' | '=' => 187,  // VK_OEM_PLUS
                '-' | '_' => 189,  // VK_OEM_MINUS
                ',' | '<' => 188,  // VK_OEM_COMMA
                '.' | '>' => 190,  // VK_OEM_PERIOD
                '/' | '?' => 191,  // VK_OEM_2
                ';' | ':' => 186,  // VK_OEM_1
                '\'' | '"' => 222, // VK_OEM_7
                '[' | '{' => 219,  // VK_OEM_4
                ']' | '}' => 221,  // VK_OEM_6
                '\\' | '|' => 220, // VK_OEM_5
                '`' | '~' => 192, // VK_OEM_3
                _ => ch as i32,
            }
        }
        Key::F1 => 112,
        Key::F2 => 113,
        Key::F3 => 114,
        Key::F4 => 115,
        Key::F5 => 116,
        Key::F6 => 117,
        Key::F7 => 118,
        Key::F8 => 119,
        Key::F9 => 120,
        Key::F10 => 121,
        Key::F11 => 122,
        Key::F12 => 123,
        Key::NumLock => 144,
        Key::ScrollLock => 145,
        _ => 0,
    }
}

fn process_side_effects(
    interp: &mut Interpreter,
    rp: RuntimeProject,
    runtime_form: &mut Signal<Option<Form>>,
    msgbox_content: &mut Signal<Option<String>>,
) {
    while let Some(effect) = interp.side_effects.pop_front() {
        match effect {
            RuntimeSideEffect::MsgBox(msg) => {
                msgbox_content.set(Some(msg));
            }
            RuntimeSideEffect::PropertyChange { object, property, value } => {
                let mut switched = false;

                let project_read = rp.project.read();
                if let Some(proj) = project_read.as_ref() {
                    if let Some(other_form_module) = proj.forms.iter().find(|f| f.form.name.eq_ignore_ascii_case(&object)) {
                        let current_is_it = if let Some(cf) = &*runtime_form.peek() {
                            cf.name.eq_ignore_ascii_case(&object)
                        } else {
                            false
                        };

                        if !current_is_it {
                            if property.eq_ignore_ascii_case("Visible")
                                && (value.as_bool().unwrap_or(false) || value.as_string() == "True")
                            {
                                let switch_code = if other_form_module.is_vbnet() {
                                    format!("{}\n{}", other_form_module.get_designer_code(), other_form_module.get_user_code())
                                } else {
                                    other_form_module.get_user_code().to_string()
                                };
                                match parse_program(&switch_code) {
                                    Ok(prog) => {
                                        if let Err(e) = interp.load_module(&other_form_module.form.name, &prog) {
                                            println!("Error loading new form code: {:?}", e);
                                        } else {
                                            let load_args = interp.make_event_handler_args(&other_form_module.form.name, "Load");
                            let _ = interp.call_event_handler(&format!("{}_Load", other_form_module.form.name), &load_args);
                                            runtime_form.set(Some(other_form_module.form.clone()));
                                            switched = true;
                                        }
                                    }
                                    Err(_) => {
                                        println!("Failed to parse new form code");
                                    }
                                }
                            }
                        }
                    }
                }

                if !switched {
                    if let Some(frm) = runtime_form.write().as_mut() {
                        let (form_part, control_part) = if let Some(idx) = object.find('.') {
                            (&object[..idx], &object[idx + 1..])
                        } else {
                            ("", object.as_str())
                        };

                        let is_for_this_form = form_part.is_empty() || form_part.eq_ignore_ascii_case(&frm.name);

                        if is_for_this_form {
                            if control_part.eq_ignore_ascii_case(&frm.name)
                                || (form_part.is_empty() && object.eq_ignore_ascii_case(&frm.name))
                            {
                                if property.eq_ignore_ascii_case("Caption") {
                                    frm.caption = value.as_string();
                                }
                            } else {
                                if let Some(ctrl) = frm.get_control_by_name_mut(control_part) {
                                    match property.to_lowercase().as_str() {
                                        "text" => {
                                            let text_val = value.as_string();
                                            ctrl.set_text(text_val.clone());
                                            if ctrl.control_type == irys_forms::ControlType::ListBox {
                                                let items: Vec<String> = text_val
                                                    .split('|')
                                                    .map(|s| s.trim().to_string())
                                                    .filter(|s| !s.is_empty())
                                                    .collect();
                                                if !items.is_empty() {
                                                    ctrl.set_list_items(items);
                                                }
                                            }
                                        }
                                        "caption" => ctrl.set_caption(value.as_string()),
                                        "left" => { if let irys_runtime::Value::Integer(v) = &value { ctrl.bounds.x = *v; } },
                                        "top" => { if let irys_runtime::Value::Integer(v) = &value { ctrl.bounds.y = *v; } },
                                        "width" => { if let irys_runtime::Value::Integer(v) = &value { ctrl.bounds.width = *v; } },
                                        "height" => { if let irys_runtime::Value::Integer(v) = &value { ctrl.bounds.height = *v; } },
                                        "visible" => { ctrl.properties.set_raw("Visible", irys_forms::PropertyValue::Boolean(value.as_bool().unwrap_or(true))); },
                                        "enabled" => { ctrl.properties.set_raw("Enabled", irys_forms::PropertyValue::Boolean(value.as_bool().unwrap_or(true))); },
                                        "backcolor" => ctrl.set_back_color(value.as_string()),
                                        "forecolor" => ctrl.set_fore_color(value.as_string()),
                                        "font" => ctrl.set_font(value.as_string()),
                                        "url" => {
                                            ctrl.properties.set("URL", value.as_string());
                                            let url = value.as_string();
                                            let _ = document::eval(&format!(
                                                r#"
                                                const iframe = document.getElementById('{}');
                                                if (iframe) {{
                                                    iframe.src = '{}';
                                                }}
                                                "#,
                                                control_part, url
                                            ));
                                        }
                                        "html" => {
                                            ctrl.properties.set("HTML", value.as_string());
                                            let html = value.as_string();
                                            let rtb_id = format!("rtb_{}", control_part);
                                            let _ = document::eval(&format!(
                                                r#"
                                                const editor = document.getElementById('{}');
                                                if (editor) {{
                                                    editor.innerHTML = '{}';
                                                }}
                                                "#,
                                                rtb_id,
                                                html.replace("'", "\\'").replace("\n", "\\n")
                                            ));
                                        }
                                        _ => {
                                            let prop_val = match value {
                                                irys_runtime::Value::Integer(i) => irys_forms::PropertyValue::Integer(i),
                                                irys_runtime::Value::String(s) => irys_forms::PropertyValue::String(s),
                                                irys_runtime::Value::Boolean(b) => irys_forms::PropertyValue::Boolean(b),
                                                irys_runtime::Value::Double(d) => irys_forms::PropertyValue::Double(d),
                                                _ => irys_forms::PropertyValue::String(value.to_string()),
                                            };
                                            ctrl.properties.set_raw(&property, prop_val);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            RuntimeSideEffect::ConsoleOutput(msg) => {
                print!("{}", msg);
            }
            RuntimeSideEffect::ConsoleClear => {
                // Console cleared
            }
            RuntimeSideEffect::DataSourceChanged { control_name, columns, rows } => {
                // Update the DataGridView control's grid data
                if let Some(frm) = runtime_form.write().as_mut() {
                    if let Some(ctrl) = frm.get_control_by_name_mut(&control_name) {
                        // Store columns and rows as serialized JSON in properties
                        ctrl.properties.set_raw("__grid_columns",
                            irys_forms::PropertyValue::StringArray(columns.clone()));
                        let row_strs: Vec<String> = rows.iter()
                            .map(|r| r.join("\t"))
                            .collect();
                        ctrl.properties.set_raw("__grid_rows",
                            irys_forms::PropertyValue::StringArray(row_strs));
                    }
                }
            }
            RuntimeSideEffect::BindingPositionChanged { binding_source_name, position, count } => {
                // Update BindingNavigator display for navigators linked to this BindingSource
                if let Some(frm) = runtime_form.write().as_mut() {
                    for ctrl in &mut frm.controls {
                        if matches!(ctrl.control_type, irys_forms::ControlType::BindingNavigator) {
                            let ctrl_bs = ctrl.properties.get_string("BindingSource").unwrap_or_default();
                            if ctrl_bs.eq_ignore_ascii_case(&binding_source_name) {
                                let count_text = format!("{} of {}", position + 1, count);
                                ctrl.set_text(count_text);
                            }
                        }
                    }
                }
            }
            RuntimeSideEffect::FormClose { form_name } => {
                // Fire FormClosing event, check Cancel, then fire FormClosed, then hide
                let closing_args = interp.make_event_handler_args(&form_name, "FormClosing");
                let _ = interp.call_event_handler(&format!("{}_FormClosing", form_name), &closing_args);
                // Check if Cancel was set to True on the EventArgs
                let cancel = if let Value::Object(ref ea) = closing_args[1] {
                    ea.borrow().fields.get("cancel").map(|v| v.as_bool().unwrap_or(false)).unwrap_or(false)
                } else {
                    false
                };
                if !cancel {
                    let closed_args = interp.make_event_handler_args(&form_name, "FormClosed");
                    let _ = interp.call_event_handler(&format!("{}_FormClosed", form_name), &closed_args);
                    // Hide the form — read first, then set
                    let should_close = runtime_form.read().as_ref()
                        .map(|frm| frm.name.eq_ignore_ascii_case(&form_name))
                        .unwrap_or(false);
                    if should_close {
                        runtime_form.set(None);
                    }
                }
            }
            RuntimeSideEffect::FormShowDialog { form_name } => {
                // Show another form as modal (same as Show for now — full modal requires blocking)
                let project_read = rp.project.read();
                if let Some(proj) = project_read.as_ref() {
                    if let Some(form_module) = proj.forms.iter().find(|f| f.form.name.eq_ignore_ascii_case(&form_name)) {
                        let switch_code = if form_module.is_vbnet() {
                            format!("{}\n{}", form_module.get_designer_code(), form_module.get_user_code())
                        } else {
                            form_module.get_user_code().to_string()
                        };
                        if let Ok(prog) = parse_program(&switch_code) {
                            let _ = interp.load_module(&form_module.form.name, &prog);
                            let load_args = interp.make_event_handler_args(&form_module.form.name, "Load");
                            let _ = interp.call_event_handler(&format!("{}_Load", form_module.form.name), &load_args);
                            runtime_form.set(Some(form_module.form.clone()));
                        }
                    }
                }
            }
            RuntimeSideEffect::AddControl { form_name, control_name, control_type, left, top, width, height } => {
                // Dynamically add a control at runtime
                if let Some(frm) = runtime_form.write().as_mut() {
                    if frm.name.eq_ignore_ascii_case(&form_name) || form_name.is_empty() {
                        let ct = irys_forms::ControlType::from_name(&control_type);
                        if let Some(ct) = ct {
                            let mut ctrl = irys_forms::Control::new(ct, control_name.clone(), left, top);
                            ctrl.bounds.width = width;
                            ctrl.bounds.height = height;
                            frm.controls.push(ctrl);
                        }
                    }
                }
            }
        }
    }

    // Sync Step: Poll interpreter for control state changes to reflect code updates in UI
    if let Some(frm) = runtime_form.write().as_mut() {
        let instance_name = format!("{}Instance", frm.name);

        if let Ok(irys_runtime::Value::Object(form_obj)) = interp.env.get(&instance_name) {
            let form_borrow = form_obj.borrow();

            for control in frm.controls.iter_mut() {
                if let Some(irys_runtime::Value::Object(ctrl_obj)) =
                    form_borrow.fields.get(&control.name.to_lowercase())
                {
                    let ctrl_fields = &ctrl_obj.borrow().fields;

                    if let Some(irys_runtime::Value::String(s)) = ctrl_fields.get("caption") {
                        control.properties.set_raw("Caption", irys_forms::PropertyValue::String(s.clone()));
                        control.properties.set_raw("Text", irys_forms::PropertyValue::String(s.clone()));
                    }

                    if let Some(irys_runtime::Value::String(s)) = ctrl_fields.get("text") {
                        control.properties.set_raw("Text", irys_forms::PropertyValue::String(s.clone()));
                    }

                    if let Some(val) = ctrl_fields.get("enabled") {
                        match val {
                            irys_runtime::Value::Boolean(b) => {
                                control.properties.set_raw("Enabled", irys_forms::PropertyValue::Boolean(*b));
                            }
                            irys_runtime::Value::Integer(i) => {
                                control.properties.set_raw("Enabled", irys_forms::PropertyValue::Boolean(*i != 0));
                            }
                            _ => {}
                        }
                    }

                    if let Some(val) = ctrl_fields.get("visible") {
                        match val {
                            irys_runtime::Value::Boolean(b) => {
                                control.properties.set_raw("Visible", irys_forms::PropertyValue::Boolean(*b));
                            }
                            irys_runtime::Value::Integer(i) => {
                                control.properties.set_raw("Visible", irys_forms::PropertyValue::Boolean(*i != 0));
                            }
                            _ => {}
                        }
                    }
                }
            }
        }
    }
}

// ---------------------------------------------------------------------------
// FormRunner – the single, shared Dioxus component used by both the editor
// and the standalone CLI shell.
// ---------------------------------------------------------------------------

#[component]
pub fn FormRunner() -> Element {
    let mut rp = use_context::<RuntimeProject>();

    let mut interpreter = use_signal(|| None::<Interpreter>);
    let mut runtime_form = use_signal(|| None::<Form>);
    let mut msgbox_content = use_signal(|| None::<String>);
    let mut parse_error = use_signal(|| None::<String>);

    // ── Console mode state ──────────────────────────────────────────────
    let mut console_output = use_signal(String::new);
    let mut console_waiting_input = use_signal(|| false);
    let mut console_finished = use_signal(|| false);
    let mut console_input_line = use_signal(String::new);
    let mut is_console_mode = use_signal(|| false);

    // Channel endpoints wrapped so they can live in signals.
    type TxWrap = Rc<RefCell<Option<mpsc::Sender<String>>>>;
    type RxWrap = Rc<RefCell<Option<mpsc::Receiver<ConsoleMessage>>>>;
    let mut console_input_tx: Signal<Option<TxWrap>> = use_signal(|| None);
    let mut console_rx: Signal<Option<RxWrap>> = use_signal(|| None);

    // ── Initialize Runtime ──────────────────────────────────────────────
    use_effect(move || {
        if interpreter.read().is_none() && !*is_console_mode.read() {
            let project_read = rp.project.read();
            if let Some(proj) = project_read.as_ref() {
                // ── Sub Main mode (interactive console) ─────────────────
                if proj.starts_with_main() {
                    // Collect all code + resources as owned Strings so we can
                    // move them to the background thread.
                    let resource_entries = crate::runner::collect_resource_entries(proj);
                    let code_files: Vec<String> = proj.code_files.iter()
                        .map(|cf| cf.code.clone())
                        .collect();
                    let form_sources: Vec<(String, String)> = proj.forms.iter()
                        .map(|fm| {
                            let code = if fm.is_vbnet() {
                                format!("{}\n{}", fm.get_designer_code(), fm.get_user_code())
                            } else {
                                fm.get_user_code().to_string()
                            };
                            (fm.form.name.clone(), code)
                        })
                        .collect();
                    drop(project_read);

                    // Set up channels
                    let (msg_tx, msg_rx) = mpsc::channel::<ConsoleMessage>();
                    let (input_tx, input_rx) = mpsc::channel::<String>();

                    // Store channel endpoints for the UI
                    console_rx.set(Some(Rc::new(RefCell::new(Some(msg_rx)))));
                    console_input_tx.set(Some(Rc::new(RefCell::new(Some(input_tx)))));
                    is_console_mode.set(true);

                    // Spawn the interpreter on a background thread
                    std::thread::spawn(move || {
                        let mut interp = Interpreter::new();
                        interp.console_tx = Some(msg_tx.clone());
                        interp.console_input_rx = Some(input_rx);
                        interp.register_resource_entries(resource_entries);

                        // Load all code files
                        for code in &code_files {
                            if let Ok(program) = parse_program(code) {
                                let _ = interp.load_code_file(&program);
                            }
                        }
                        // Load form modules
                        for (name, code) in &form_sources {
                            if let Ok(program) = parse_program(code) {
                                let _ = interp.load_module(name, &program);
                            }
                        }

                        // Run Sub Main
                        match interp.call_procedure(&irys_parser::ast::Identifier::new("main"), &[]) {
                            Ok(_) => { let _ = msg_tx.send(ConsoleMessage::Finished); }
                            Err(e) => { let _ = msg_tx.send(ConsoleMessage::Error(format!("{:?}", e))); }
                        }
                    });

                    return;
                }

                // ── Form mode ───────────────────────────────────────────
                if let Some(startup_form_module) = proj.get_startup_form() {
                    let form = startup_form_module.form.clone();
                    let form_code = if startup_form_module.is_vbnet() {
                        format!("{}\n{}", startup_form_module.get_designer_code(), startup_form_module.get_user_code())
                    } else {
                        startup_form_module.get_user_code().to_string()
                    };
                    let is_vbnet_form = startup_form_module.is_vbnet();
                    drop(project_read);

                    runtime_form.set(Some(form.clone()));

                    let mut interp = Interpreter::new();

                    // Register resources from all project resource files + form resources
                    if let Some(proj) = rp.project.read().as_ref() {
                        let entries = crate::runner::collect_resource_entries(proj);
                        interp.register_resource_entries(entries);
                    }

                    // Load all code files (global scope + class definitions)
                    let project_read = rp.project.read();
                    if let Some(proj) = project_read.as_ref() {
                        for code_file in &proj.code_files {
                            if let Ok(program) = parse_program(&code_file.code) {
                                let _ = interp.run(&program);
                            }
                        }
                    }
                    drop(project_read);

                    // Now load the startup form code
                    match parse_program(&form_code) {
                        Ok(program) => {
                            parse_error.set(None);
                            if let Err(e) = interp.run(&program) {
                                parse_error.set(Some(format!("Runtime Load Error: {:?}", e)));
                            } else {
                                // Register form controls as variables
                                if let Some(runtime_form_data) = runtime_form.read().as_ref() {
                                    for control in &runtime_form_data.controls {
                                        let control_var = irys_runtime::Value::String(control.name.clone());
                                        interp.env.define(&control.name, control_var);
                                    }
                                }

                                // For VB.NET, create an instance of the form
                                if is_vbnet_form {
                                    let runtime_controls = r#"
                                        Module RuntimeControls
                                            Public Class Control
                                                Public Name As String
                                                Public Caption As String
                                                Public Text As String
                                                Public Value As Integer
                                                Public Enabled As Boolean
                                                Public Visible As Boolean

                                                Public Sub Click()
                                                End Sub
                                            End Class
                                        End Module
                                    "#;
                                    if let Ok(prog) = parse_program(runtime_controls) {
                                        interp.load_module("RuntimeControls", &prog).ok();
                                    }

                                    let global_module = format!(
                                        r#"
                                        Module RuntimeGlobals
                                            Public {}Instance As New {}

                                            Public Sub InitControls()
                                                ' Auto-generated control initialization
                                                {}
                                            End Sub
                                        End Module
                                    "#,
                                        form.name,
                                        form.name,
                                        if let Some(runtime_form_data) = runtime_form.read().as_ref() {
                                            runtime_form_data
                                                .controls
                                                .iter()
                                                .map(|c| {
                                                    let cap = c
                                                        .get_caption()
                                                        .map(|s| s.to_string())
                                                        .unwrap_or_else(|| c.name.clone());
                                                    let en = c.is_enabled();
                                                    format!(
                                                    "RuntimeGlobals.{0}Instance.{1} = New RuntimeControls.Control\nRuntimeGlobals.{0}Instance.{1}.Name = \"{1}\"\nRuntimeGlobals.{0}Instance.{1}.Caption = \"{2}\"\nRuntimeGlobals.{0}Instance.{1}.Enabled = {3}\n",
                                                    form.name,
                                                    c.name,
                                                    cap.replace("\"", "\"\""),
                                                    if en { "True" } else { "False" }
                                                )
                                                })
                                                .collect::<Vec<_>>()
                                                .join("\n")
                                        } else {
                                            String::new()
                                        }
                                    );
                                    if let Ok(prog) = parse_program(&global_module) {
                                        if let Err(e) = interp.load_module("RuntimeGlobals", &prog) {
                                            println!("Error creating runtime globals: {:?}", e);
                                        } else {
                                            let _ = interp.call_procedure(
                                                &irys_parser::ast::Identifier::new("InitControls"),
                                                &[],
                                            );
                                        }
                                    }

                                    let instance_name = format!("RuntimeGlobals.{}Instance", form.name);
                                    let form_name_lower = form.name.to_lowercase();

                                    match interp.call_instance_method(&instance_name, "InitializeComponent", &[]) {
                                        Ok(_) => {
                                            if let Some(Value::Object(form_obj)) =
                                                interp.env.get_global(&instance_name)
                                            {
                                                let mut form_ref = form_obj.borrow_mut();
                                                if let Some(runtime_form_data) = runtime_form.read().as_ref() {
                                                    for ctrl in &runtime_form_data.controls {
                                                        let key = ctrl.name.to_lowercase();
                                                        let needs_create = matches!(
                                                            form_ref.fields.get(&key),
                                                            None | Some(Value::Nothing) | Some(Value::String(_))
                                                        );
                                                        if needs_create {
                                                            let mut ctrl_fields = HashMap::new();
                                                            ctrl_fields.insert(
                                                                "name".to_string(),
                                                                Value::String(ctrl.name.clone()),
                                                            );
                                                            let caption = ctrl
                                                                .get_caption()
                                                                .map(|s| s.to_string())
                                                                .unwrap_or_else(|| ctrl.name.clone());
                                                            ctrl_fields.insert(
                                                                "caption".to_string(),
                                                                Value::String(caption),
                                                            );
                                                            let text = ctrl
                                                                .get_text()
                                                                .map(|s| s.to_string())
                                                                .unwrap_or_default();
                                                            ctrl_fields.insert(
                                                                "text".to_string(),
                                                                Value::String(text),
                                                            );
                                                            ctrl_fields.insert(
                                                                "enabled".to_string(),
                                                                Value::Boolean(ctrl.is_enabled()),
                                                            );
                                                            let ctrl_obj = Value::Object(Rc::new(
                                                                RefCell::new(ObjectData {
                                                                    class_name: "RuntimeControls.Control".to_string(),
                                                                    fields: ctrl_fields,
                                                                }),
                                                            ));
                                                            form_ref.fields.insert(key.clone(), ctrl_obj.clone());
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        Err(e) => println!("InitializeComponent error: {:?}", e),
                                    }

                                    let load_args = interp.make_event_handler_args(&form.name, "Load");
                                    if let Some(load_handler) =
                                        interp.find_handles_method(&form_name_lower, "Me", "Load")
                                    {
                                        let _ = interp.call_instance_method(&instance_name, &load_handler, &load_args);
                                    } else {
                                        let _ = interp.call_instance_method(
                                            &instance_name,
                                            &format!("{}_Load", form.name),
                                            &load_args,
                                        );
                                        let _ = interp.call_instance_method(&instance_name, "Form_Load", &load_args);
                                    }

                                    let _ = interp.call_event_handler("Form_Load", &load_args);
                                } else {
                                    let load_args = interp.make_event_handler_args("Form", "Load");
                                    match interp.call_event_handler("Form_Load", &load_args) {
                                        Ok(_) => {}
                                        Err(irys_runtime::RuntimeError::UndefinedFunction(_)) => {}
                                        Err(e) => println!("Form_Load Error: {:?}", e),
                                    }
                                }
                                process_side_effects(&mut interp, rp, &mut runtime_form, &mut msgbox_content);

                                // Fire Shown event after the form is first displayed
                                let shown_form_name = runtime_form.read().as_ref().map(|frm| frm.name.clone());
                                if let Some(fname) = shown_form_name {
                                    let shown_args = interp.make_event_handler_args(&fname, "Shown");
                                    let _ = interp.call_event_handler(&format!("{}_Shown", fname), &shown_args);
                                    process_side_effects(&mut interp, rp, &mut runtime_form, &mut msgbox_content);
                                }
                            }
                        }
                        Err(e) => {
                            parse_error.set(Some(format!("Parse Error: {:?}", e)));
                        }
                    }
                    interpreter.set(Some(interp));
                } else {
                    println!("No startup form defined in project");
                }
            }
        }
    });

    // ── Event handler ───────────────────────────────────────────────────
    let mut handle_event = move |control_name: String, event_name: String, event_data: Option<irys_runtime::EventData>| {
        if parse_error.read().is_some() {
            return;
        }
        if msgbox_content.read().is_some() {
            return;
        }

        if let Some(interp) = interpreter.write().as_mut() {
            if let Some(frm) = runtime_form.read().as_ref() {
                let is_vbnet = rp
                    .project
                    .read()
                    .as_ref()
                    .and_then(|p| p.get_startup_form())
                    .map(|f| f.is_vbnet())
                    .unwrap_or(false);

                // Common global sync
                interp
                    .env
                    .set(&format!("{}.Caption", frm.name), irys_runtime::Value::String(frm.caption.clone()))
                    .ok();
                for ctrl in &frm.controls {
                    let ctrl_name = ctrl.name.clone();
                    let caption = ctrl.get_caption().map(|s| s.to_string()).unwrap_or(ctrl_name.clone());
                    let text = ctrl.get_text().map(|s| s.to_string()).unwrap_or_default();
                    let val = ctrl.properties.get_int("Value").unwrap_or(0);

                    interp
                        .env
                        .set(&format!("{}.Caption", ctrl_name), irys_runtime::Value::String(caption.clone()))
                        .ok();
                    interp
                        .env
                        .set(&format!("{}.Text", ctrl_name), irys_runtime::Value::String(text.clone()))
                        .ok();
                    interp
                        .env
                        .set(&format!("{}.Value", ctrl_name), irys_runtime::Value::Integer(val))
                        .ok();

                    if is_vbnet {
                        let instance_name = format!("RuntimeGlobals.{}Instance", frm.name);
                        let sync_script = format!(
                            r#"
                            {}.{}.Caption = "{}"
                            {}.{}.Text = "{}"
                         "#,
                            instance_name,
                            ctrl_name,
                            caption.replace("\"", "\"\""),
                            instance_name,
                            ctrl_name,
                            text.replace("\"", "\"\"")
                        );
                        if let Ok(prog) = parse_program(&sync_script) {
                            interp.load_module("SyncState", &prog).ok();
                        }
                    }
                }
            }

            let handler_name = if let Some(frm) = runtime_form.read().as_ref() {
                if let Some(event_type) = event_type_from_name(&event_name) {
                    if control_name.eq_ignore_ascii_case(&frm.name) {
                        if let Some(handler) = frm.get_event_handler(&control_name, &event_type) {
                            handler.to_string()
                        } else {
                            format!("{}_{}", frm.name, event_name)
                        }
                    } else {
                        if let Some(handler) = frm.get_event_handler(&control_name, &event_type) {
                            handler.to_string()
                        } else {
                            format!("{}_{}", control_name, event_name)
                        }
                    }
                } else {
                    format!("{}_{}", control_name, event_name)
                }
            } else {
                format!("{}_{}", control_name, event_name)
            };

            let handler_lower = handler_name.trim().to_lowercase();

            let mut executed = false;

            let is_vbnet = rp
                .project
                .read()
                .as_ref()
                .and_then(|p| p.get_startup_form())
                .map(|f| f.is_vbnet())
                .unwrap_or(false);

            if is_vbnet {
                if let Some(frm) = runtime_form.read().as_ref() {
                    let instance_name = format!("RuntimeGlobals.{}Instance", frm.name);
                    let form_class = frm.name.to_lowercase();

                    let event_args = interp.make_event_handler_args_with_data(&control_name, &event_name, event_data.as_ref());

                    if let Some(method_name) =
                        interp.find_handles_method(&form_class, &control_name, &event_name)
                    {
                        match interp.call_instance_method(&instance_name, &method_name, &event_args) {
                            Ok(_) => executed = true,
                            Err(_) => {}
                        }
                    }

                    if !executed {
                        match interp.call_instance_method(&instance_name, &handler_name, &event_args) {
                            Ok(_) => executed = true,
                            Err(_) => {}
                        }
                    }
                    if !executed && control_name.eq_ignore_ascii_case(&frm.name) {
                        let conv_name = format!("{}_{}", frm.name, event_name);
                        match interp.call_instance_method(&instance_name, &conv_name, &event_args) {
                            Ok(_) => executed = true,
                            Err(_) => {}
                        }
                    }
                }
            }

            if !executed {
                let handler_key = interp
                    .subs
                    .keys()
                    .find(|key| *key == &handler_lower || key.ends_with(&format!(".{}", handler_lower)))
                    .cloned();

                if let Some(key) = handler_key {
                    let event_args = interp.make_event_handler_args_with_data(&control_name, &event_name, event_data.as_ref());
                    let _ = interp.call_event_handler(&key, &event_args);
                }
            }

            // BindingNavigator auto-delegation: if the event is a navigation
            // action (MoveFirst, MoveNext, …) and the control is a BindingNavigator,
            // forward the call to its BindingSource.
            if let Some(frm) = runtime_form.read().as_ref() {
                let is_nav_event = matches!(
                    event_name.as_str(),
                    "MoveFirst" | "MoveNext" | "MovePrevious" | "MoveLast" | "AddNew" | "Delete"
                );
                if is_nav_event {
                    if let Some(ctrl) = frm.controls.iter().find(|c| c.name == control_name) {
                        if matches!(ctrl.control_type, irys_forms::ControlType::BindingNavigator) {
                            let bs_name = ctrl.properties.get_string("BindingSource").unwrap_or_default();
                            if !bs_name.is_empty() {
                                let instance_name = format!("RuntimeGlobals.{}Instance", frm.name);
                                // Call bs.MoveFirst() etc. through the interpreter
                                let nav_script = format!("{}.{}.{}()", instance_name, bs_name, event_name);
                                eprintln!("[Nav] script: {}", nav_script);
                                if let Ok(prog) = parse_program(&nav_script) {
                                    let _ = interp.load_module("NavAction", &prog);
                                }
                            }
                        }
                    }
                }
            }

            process_side_effects(interp, rp, &mut runtime_form, &mut msgbox_content);
        }
    };

    // ── Console message polling ─────────────────────────────────────────
    // In console mode, poll the channel from the interpreter thread every
    // 50 ms and update signals accordingly.
    use_future(move || async move {
        loop {
            tokio::time::sleep(std::time::Duration::from_millis(50)).await;
            if !*is_console_mode.read() { continue; }

            // Drain all available messages
            let rx_opt = console_rx.read().clone();
            if let Some(rx_rc) = rx_opt {
                let rx_ref = rx_rc.borrow();
                if let Some(rx) = rx_ref.as_ref() {
                    loop {
                        match rx.try_recv() {
                            Ok(ConsoleMessage::Output { text, fg, bg }) => {
                                let escaped = html_escape(&text);
                                // Convert newlines to <br> since we use innerHTML
                                let with_br = escaped.replace('\n', "<br>");
                                let fg_css = console_color_to_css(fg);
                                if bg == 0 {
                                    // No background (transparent) for default
                                    console_output.write().push_str(
                                        &format!("<span style=\"color:{}\">{}</span>", fg_css, with_br)
                                    );
                                } else {
                                    let bg_css = console_color_to_css(bg);
                                    console_output.write().push_str(
                                        &format!("<span style=\"color:{};background:{}\">{}</span>", fg_css, bg_css, with_br)
                                    );
                                }
                            }
                            Ok(ConsoleMessage::Clear) => {
                                console_output.set(String::new());
                            }
                            Ok(ConsoleMessage::InputRequest) => {
                                console_waiting_input.set(true);
                            }
                            Ok(ConsoleMessage::Finished) => {
                                console_finished.set(true);
                                rp.finished.set(true);
                                break;
                            }
                            Ok(ConsoleMessage::Error(e)) => {
                                console_output.write().push_str(
                                    &format!("<br><span style=\"color:#e74856\">--- Error: {} ---</span><br>", html_escape(&e))
                                );
                                console_finished.set(true);
                                rp.finished.set(true);
                                break;
                            }
                            Err(mpsc::TryRecvError::Empty) => break,
                            Err(mpsc::TryRecvError::Disconnected) => {
                                console_finished.set(true);
                                rp.finished.set(true);
                                break;
                            }
                        }
                    }
                }
            }
        }
    });

    // ── Render ───────────────────────────────────────────────────────────
    let form_opt = runtime_form.read().clone();
    let msgbox_visible = msgbox_content.read().is_some();
    let msgbox_text = msgbox_content.read().clone().unwrap_or_default();
    let error_text = parse_error.read().clone();

    rsx! {
        div {
            class: "runtime-panel",
            style: "flex: 1; background: #e0e0e0; display: flex; align-items: center; justify-content: center; overflow: auto; position: relative;",

            {
                if *is_console_mode.read() {
                    // ── Interactive Console UI ──────────────────────────
                    let output_text = console_output.read().clone();
                    let waiting = *console_waiting_input.read();
                    let finished = *console_finished.read();
                    let input_val = console_input_line.read().clone();

                    rsx! {
                        div {
                            style: "
                                width: 700px;
                                height: 480px;
                                background: #1e1e1e;
                                color: #d4d4d4;
                                border: 1px solid #444;
                                box-shadow: 0 0 10px rgba(0,0,0,0.5);
                                display: flex;
                                flex-direction: column;
                                font-family: 'Consolas', 'Courier New', monospace;
                                font-size: 14px;
                            ",

                            // Title bar
                            div {
                                style: "
                                    background: #2d2d2d;
                                    color: #ccc;
                                    padding: 4px 10px;
                                    font-size: 12px;
                                    border-bottom: 1px solid #444;
                                    user-select: none;
                                ",
                                "Console Output"
                            }

                            // Output area
                            div {
                                id: "console-output",
                                style: "
                                    flex: 1;
                                    overflow-y: auto;
                                    padding: 8px 10px;
                                    white-space: pre-wrap;
                                    word-break: break-all;
                                ",
                                dangerous_inner_html: "{output_text}",
                            }

                            // Input area (shown when waiting for input, or always visible with prompt)
                            if waiting && !finished {
                                div {
                                    style: "
                                        display: flex;
                                        border-top: 1px solid #444;
                                        background: #252526;
                                    ",
                                    span {
                                        style: "padding: 6px 4px 6px 10px; color: #569cd6;",
                                        ">"
                                    }
                                    input {
                                        id: "console-input",
                                        style: "
                                            flex: 1;
                                            background: #252526;
                                            color: #d4d4d4;
                                            border: none;
                                            outline: none;
                                            padding: 6px 8px;
                                            font-family: 'Consolas', 'Courier New', monospace;
                                            font-size: 14px;
                                        ",
                                        value: "{input_val}",
                                        autofocus: true,
                                        oninput: move |evt| {
                                            console_input_line.set(evt.value().clone());
                                        },
                                        onkeypress: move |evt: KeyboardEvent| {
                                            if evt.key() == Key::Enter {
                                                let line = console_input_line.read().clone();
                                                // Echo the input to the output
                                                console_output.write().push_str(&format!("{}\n", line));
                                                // Send to the interpreter thread
                                                if let Some(tx_rc) = console_input_tx.read().as_ref() {
                                                    if let Some(tx) = tx_rc.borrow().as_ref() {
                                                        let _ = tx.send(line);
                                                    }
                                                }
                                                console_input_line.set(String::new());
                                                console_waiting_input.set(false);
                                            }
                                        },
                                    }
                                }
                            }

                            // Status bar
                            div {
                                style: "
                                    background: #007acc;
                                    color: white;
                                    padding: 2px 10px;
                                    font-size: 11px;
                                ",
                                if finished {
                                    "Program finished"
                                } else if waiting {
                                    "Waiting for input..."
                                } else {
                                    "Running..."
                                }
                            }
                        }
                    }
                } else if let Some(err) = error_text {
                    rsx! {
                        div {
                            style: "color: red; padding: 20px; background: #fee; border: 1px solid red;",
                            "Error: {err}"
                        }
                    }
                } else if let Some(form) = form_opt {
                    let width = form.width;
                    let height = form.height;
                    let caption = form.caption.clone();
                    let form_back = form.back_color.clone().unwrap_or_else(|| "#f8fafc".to_string());
                    let form_fore = form.fore_color.clone().unwrap_or_else(|| "#0f172a".to_string());
                    let form_font = form.font.clone().unwrap_or_else(|| "Segoe UI, 12px".to_string());

                    rsx! {
                        div {
                            style: "
                                position: relative;
                                width: {width}px;
                                height: {height}px;
                                background: {form_back};
                                color: {form_fore};
                                border: 1px solid #999;
                                box-shadow: 0 0 10px rgba(0,0,0,0.5);
                                font: {form_font};
                            ",

                            // Title bar
                            div {
                                style: "
                                    height: 30px;
                                    background: linear-gradient(to bottom, #0078d4, #005a9e);
                                    color: {form_fore};
                                    padding: 4px 8px;
                                    display: flex;
                                    align-items: center;
                                    font-weight: bold;
                                ",
                                span { style: "flex: 1;", "{caption}" }
                                // Close button
                                button {
                                    style: "background: transparent; border: none; color: white; font-size: 16px; cursor: pointer; padding: 0 4px; margin-left: auto; font-weight: bold; line-height: 1;",
                                    onclick: move |_| {
                                        let form_name_close = form.name.clone();
                                        handle_event(form_name_close, "FormClosing".to_string(), None);
                                    },
                                    "\u{00D7}"
                                }
                            }

                            // Controls (skip non-visual components)
                            for control in form.controls.into_iter().filter(|c| !c.control_type.is_non_visual()) {
                                {
                                    let control_type = control.control_type;
                                    let x = control.bounds.x;
                                    let y = control.bounds.y;
                                    let w = control.bounds.width;
                                    let h = control.bounds.height;
                                    let name = control.name.clone();
                                    let name_clone = name.clone();
                                    let name_mousemove = name.clone();
                                    let name_mouseenter = name.clone();
                                    let name_mouseleave = name.clone();
                                    let name_mousedown = name.clone();
                                    let name_mouseup = name.clone();
                                    let name_wheel = name.clone();
                                    let name_keydown = name.clone();
                                    let name_keyup = name.clone();
                                    let name_keypress = name.clone();
                                    let name_dblclick = name.clone();
                                    let name_focusin = name.clone();
                                    let name_focusout = name.clone();

                                    let caption = control.get_caption().map(|s| s.to_string()).unwrap_or(name.clone());
                                    let text = control.get_text().map(|s| s.to_string()).unwrap_or_default();
                                    let is_enabled = control.is_enabled();
                                    let is_visible = control.is_visible();

                                    let base_font = "font-family: 'Segoe UI', sans-serif; color: #0f172a;";
                                    let base_field_bg = "background: #f8fafc;";
                                    let base_button_bg = "background: linear-gradient(90deg, #2563eb, #1d4ed8); color: #f8fafc; border: 1px solid #1d4ed8;";
                                    let base_frame_border = "border: 1px solid #cbd5e1;";

                                    let back_color = control.get_back_color().map(|s| s.to_string());
                                    let fore_color = control.get_fore_color().map(|s| s.to_string());
                                    let font_family = control.get_font().map(|s| s.to_string());

                                    let mut style_font = base_font.to_string();
                                    if let Some(f) = &font_family {
                                        style_font = format!("font: {};", f);
                                    }
                                    let mut style_fore = String::new();
                                    if let Some(fc) = &fore_color {
                                        style_fore = format!("color: {};", fc);
                                    }
                                    let mut style_back = String::new();
                                    if let Some(bc) = &back_color {
                                        style_back = format!("background: {};", bc);
                                    }
                                    let button_bg = if let Some(bc) = &back_color {
                                        let color_part = if let Some(fc) = &fore_color {
                                            format!("color: {};", fc)
                                        } else {
                                            String::new()
                                        };
                                        format!("background: {}; {}; border: 1px solid #cbd5e1;", bc, color_part)
                                    } else {
                                        base_button_bg.to_string()
                                    };

                                    // Compute Dock-based positioning
                                    let dock_val = control.properties.get_int("Dock").unwrap_or(0);
                                    let wrapper_style = match dock_val {
                                        1 => "position: absolute; left: 0; top: 0; width: 100%; outline: none;".to_string(), // DockStyle.Top
                                        2 => "position: absolute; left: 0; bottom: 0; width: 100%; outline: none;".to_string(), // DockStyle.Bottom
                                        3 => "position: absolute; left: 0; top: 0; height: 100%; outline: none;".to_string(), // DockStyle.Left
                                        4 => "position: absolute; right: 0; top: 0; height: 100%; outline: none;".to_string(), // DockStyle.Right
                                        5 => "position: absolute; left: 0; top: 0; width: 100%; height: 100%; outline: none;".to_string(), // DockStyle.Fill
                                        _ => format!("position: absolute; left: {}px; top: {}px; width: {}px; height: {}px; outline: none;", x, y, w, h),
                                    };

                                    rsx! {
                                        if is_visible {
                                            div {
                                                tabindex: "0",
                                                style: "{wrapper_style}",
                                                onmousemove: move |evt: MouseEvent| {
                                                    let data = irys_runtime::EventData::Mouse {
                                                        button: 0, clicks: 0,
                                                        x: evt.client_coordinates().x as i32,
                                                        y: evt.client_coordinates().y as i32,
                                                        delta: 0,
                                                    };
                                                    handle_event(name_mousemove.clone(), "MouseMove".to_string(), Some(data));
                                                },
                                                onmouseenter: move |evt: MouseEvent| {
                                                    let data = irys_runtime::EventData::Mouse {
                                                        button: 0, clicks: 0,
                                                        x: evt.client_coordinates().x as i32,
                                                        y: evt.client_coordinates().y as i32,
                                                        delta: 0,
                                                    };
                                                    handle_event(name_mouseenter.clone(), "MouseEnter".to_string(), Some(data));
                                                },
                                                onmouseleave: move |evt: MouseEvent| {
                                                    handle_event(name_mouseleave.clone(), "MouseLeave".to_string(), None);
                                                },
                                                onmousedown: move |evt: MouseEvent| {
                                                    let btn = match evt.trigger_button() {
                                                        Some(dioxus::html::input_data::MouseButton::Primary) => 0x100000,
                                                        Some(dioxus::html::input_data::MouseButton::Secondary) => 0x200000,
                                                        Some(dioxus::html::input_data::MouseButton::Auxiliary) => 0x400000,
                                                        _ => 0,
                                                    };
                                                    let data = irys_runtime::EventData::Mouse {
                                                        button: btn, clicks: 1,
                                                        x: evt.client_coordinates().x as i32,
                                                        y: evt.client_coordinates().y as i32,
                                                        delta: 0,
                                                    };
                                                    handle_event(name_mousedown.clone(), "MouseDown".to_string(), Some(data));
                                                },
                                                onmouseup: move |evt: MouseEvent| {
                                                    let btn = match evt.trigger_button() {
                                                        Some(dioxus::html::input_data::MouseButton::Primary) => 0x100000,
                                                        Some(dioxus::html::input_data::MouseButton::Secondary) => 0x200000,
                                                        Some(dioxus::html::input_data::MouseButton::Auxiliary) => 0x400000,
                                                        _ => 0,
                                                    };
                                                    let data = irys_runtime::EventData::Mouse {
                                                        button: btn, clicks: 1,
                                                        x: evt.client_coordinates().x as i32,
                                                        y: evt.client_coordinates().y as i32,
                                                        delta: 0,
                                                    };
                                                    handle_event(name_mouseup.clone(), "MouseUp".to_string(), Some(data));
                                                },
                                                ondoubleclick: move |evt: MouseEvent| {
                                                    let data = irys_runtime::EventData::Mouse {
                                                        button: 0x100000, clicks: 2,
                                                        x: evt.client_coordinates().x as i32,
                                                        y: evt.client_coordinates().y as i32,
                                                        delta: 0,
                                                    };
                                                    handle_event(name_dblclick.clone(), "DoubleClick".to_string(), Some(data));
                                                },
                                                onwheel: move |evt: WheelEvent| {
                                                    let data = irys_runtime::EventData::Mouse {
                                                        button: 0, clicks: 0, x: 0, y: 0,
                                                        delta: evt.delta().strip_units().y as i32,
                                                    };
                                                    handle_event(name_wheel.clone(), "MouseWheel".to_string(), Some(data));
                                                },
                                                onkeydown: move |evt: KeyboardEvent| {
                                                    let data = irys_runtime::EventData::Key {
                                                        key_code: dioxus_key_to_vk(&evt.key()),
                                                        shift: evt.modifiers().shift(),
                                                        ctrl: evt.modifiers().ctrl(),
                                                        alt: evt.modifiers().alt(),
                                                    };
                                                    handle_event(name_keydown.clone(), "KeyDown".to_string(), Some(data));
                                                },
                                                onkeyup: move |evt: KeyboardEvent| {
                                                    let data = irys_runtime::EventData::Key {
                                                        key_code: dioxus_key_to_vk(&evt.key()),
                                                        shift: evt.modifiers().shift(),
                                                        ctrl: evt.modifiers().ctrl(),
                                                        alt: evt.modifiers().alt(),
                                                    };
                                                    handle_event(name_keyup.clone(), "KeyUp".to_string(), Some(data));
                                                },
                                                onkeypress: move |evt: KeyboardEvent| {
                                                    // KeyPress fires for character keys — extract the char
                                                    let ch = match evt.key() {
                                                        dioxus::prelude::Key::Character(ref s) => s.chars().next().unwrap_or('\0'),
                                                        dioxus::prelude::Key::Enter => '\r',
                                                        dioxus::prelude::Key::Tab => '\t',
                                                        dioxus::prelude::Key::Backspace => '\x08',
                                                        _ => '\0',
                                                    };
                                                    if ch != '\0' {
                                                        let data = irys_runtime::EventData::KeyPress { key_char: ch };
                                                        handle_event(name_keypress.clone(), "KeyPress".to_string(), Some(data));
                                                    }
                                                },
                                                onfocusin: move |_evt: FocusEvent| {
                                                    handle_event(name_focusin.clone(), "GotFocus".to_string(), None);
                                                },
                                                onfocusout: move |_evt: FocusEvent| {
                                                    handle_event(name_focusout.clone(), "LostFocus".to_string(), None);
                                                },

                                                {match control_type {
                                                    ControlType::Button => rsx! {
                                                        button {
                                                            style: "width: 100%; height: 100%; padding: 6px 10px; {button_bg} {style_font}; border-radius: 6px; box-shadow: 0 2px 4px rgba(37,99,235,0.12);",
                                                            disabled: !is_enabled,
                                                            onclick: move |evt: MouseEvent| {
                                                                let data = irys_runtime::EventData::Mouse {
                                                                    button: 0x100000, clicks: 1,
                                                                    x: evt.client_coordinates().x as i32,
                                                                    y: evt.client_coordinates().y as i32,
                                                                    delta: 0,
                                                                };
                                                                handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                                            },
                                                            "{caption}"
                                                        }
                                                    },
                                                    ControlType::TextBox => rsx! {
                                                        input {
                                                            style: "width: 100%; height: 100%; padding: 6px 8px; border: 1px solid #cbd5e1; border-radius: 6px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                                            disabled: !is_enabled,
                                                            value: "{text}",
                                                            oninput: move |evt| {
                                                                if let Some(frm) = runtime_form.write().as_mut() {
                                                                    if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                                        ctrl.set_text(evt.value());
                                                                    }
                                                                }
                                                                handle_event(name_clone.clone(), "Change".to_string(), None);
                                                            }
                                                        }
                                                    },
                                                    ControlType::Label => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; padding: 4px 2px; {style_font} {style_fore} {style_back};",
                                                            onclick: move |evt: MouseEvent| {
                                                                let data = irys_runtime::EventData::Mouse {
                                                                    button: 0x100000, clicks: 1,
                                                                    x: evt.client_coordinates().x as i32,
                                                                    y: evt.client_coordinates().y as i32,
                                                                    delta: 0,
                                                                };
                                                                handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                                            },
                                                            "{caption}"
                                                        }
                                                    },
                                                    ControlType::CheckBox => rsx! {
                                                        div {
                                                            style: "display: flex; align-items: center; gap: 6px; {style_font} {style_fore} {style_back};",
                                                            input {
                                                                r#type: "checkbox",
                                                                disabled: !is_enabled,
                                                                checked: control.properties.get_int("Value").unwrap_or(0) == 1,
                                                                onclick: move |evt: MouseEvent| {
                                                                    let data = irys_runtime::EventData::Mouse {
                                                                        button: 0x100000, clicks: 1,
                                                                        x: evt.client_coordinates().x as i32,
                                                                        y: evt.client_coordinates().y as i32,
                                                                        delta: 0,
                                                                    };
                                                                    handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                                                }
                                                            }
                                                            span { "{caption}" }
                                                        }
                                                    },
                                                    ControlType::RadioButton => rsx! {
                                                        div {
                                                            style: "display: flex; align-items: center; gap: 6px; {style_font} {style_fore} {style_back};",
                                                            input {
                                                                r#type: "radio",
                                                                name: "radio_group",
                                                                disabled: !is_enabled,
                                                                checked: control.properties.get_int("Value").unwrap_or(0) == 1,
                                                                onclick: move |evt: MouseEvent| {
                                                                    let data = irys_runtime::EventData::Mouse {
                                                                        button: 0x100000, clicks: 1,
                                                                        x: evt.client_coordinates().x as i32,
                                                                        y: evt.client_coordinates().y as i32,
                                                                        delta: 0,
                                                                    };
                                                                    handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                                                }
                                                            }
                                                            span { "{caption}" }
                                                        }
                                                    },
                                                    ControlType::Frame => rsx! {
                                                        fieldset {
                                                            style: "width: 100%; height: 100%; {base_frame_border} margin: 0; padding: 0; border-radius: 8px; {style_back} {style_font} {style_fore};",
                                                            onclick: move |evt: MouseEvent| {
                                                                let data = irys_runtime::EventData::Mouse {
                                                                    button: 0x100000, clicks: 1,
                                                                    x: evt.client_coordinates().x as i32,
                                                                    y: evt.client_coordinates().y as i32,
                                                                    delta: 0,
                                                                };
                                                                handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                                            },
                                                            legend { "{caption}" }
                                                        }
                                                    },
                                                    ControlType::ListBox => rsx! {
                                                        select {
                                                            style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; border-radius: 8px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                                            multiple: true,
                                                            disabled: !is_enabled,
                                                            onchange: move |_| handle_event(name_clone.clone(), "Click".to_string(), None),
                                                            {
                                                                let mut items = control.get_list_items();
                                                                if items.is_empty() {
                                                                    let raw = text.clone();
                                                                    if !raw.is_empty() {
                                                                        items = raw.split('|').map(|s| s.trim().to_string()).collect();
                                                                    }
                                                                }

                                                                if items.is_empty() {
                                                                    rsx! { option { "(empty)" } }
                                                                } else {
                                                                    rsx! {
                                                                        for item in items {
                                                                            option { "{item}" }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::ComboBox => rsx! {
                                                        select {
                                                            style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; border-radius: 8px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                                            disabled: !is_enabled,
                                                            onchange: move |evt| {
                                                                if let Some(frm) = runtime_form.write().as_mut() {
                                                                    if let Some(ctrl) = frm.get_control_by_name_mut(&name_clone) {
                                                                        ctrl.set_text(evt.value());
                                                                    }
                                                                }
                                                                handle_event(name_clone.clone(), "Change".to_string(), None);
                                                            },
                                                            {
                                                                let items = control.get_list_items();
                                                                if items.is_empty() {
                                                                    rsx! { option { value: "", "{text}" } }
                                                                } else {
                                                                    rsx! {
                                                                        for item in items {
                                                                            option { "{item}" }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::RichTextBox => rsx! {
                                                        {
                                                            let html = control.properties.get_string("HTML")
                                                                .map(|s| s.to_string())
                                                                .or_else(|| control.get_text().map(|s| s.to_string()))
                                                                .unwrap_or_default();
                                                            let toolbar_visible = control.properties.get_bool("ToolbarVisible").unwrap_or(true);
                                                            let rtb_id = format!("rtb_{}", name_clone);
                                                            let rtb_id_bold = rtb_id.clone();
                                                            let rtb_id_italic = rtb_id.clone();
                                                            let rtb_id_underline = rtb_id.clone();
                                                            let rtb_id_ul = rtb_id.clone();
                                                            let rtb_id_ol = rtb_id.clone();
                                                            let rtb_id_mount = rtb_id.clone();
                                                            rsx! {
                                                                div {
                                                                    style: "width: 100%; height: 100%; display: flex; flex-direction: column; border: 1px inset #999; background: white; {style_back} {style_font} {style_fore};",
                                                                    if toolbar_visible {
                                                                        div {
                                                                            style: "display: flex; gap: 2px; padding: 4px; background: #f0f0f0; border-bottom: 1px solid #ccc;",
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; font-weight: bold;",
                                                                                title: "Bold (Ctrl+B)",
                                                                                onclick: move |_| {
                                                                                    let _ = document::eval(&format!("document.execCommand('bold', false, null); document.getElementById('{}').focus();", rtb_id_bold));
                                                                                },
                                                                                "B"
                                                                            }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; font-style: italic;",
                                                                                title: "Italic (Ctrl+I)",
                                                                                onclick: move |_| {
                                                                                    let _ = document::eval(&format!("document.execCommand('italic', false, null); document.getElementById('{}').focus();", rtb_id_italic));
                                                                                },
                                                                                "I"
                                                                            }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; text-decoration: underline;",
                                                                                title: "Underline (Ctrl+U)",
                                                                                onclick: move |_| {
                                                                                    let _ = document::eval(&format!("document.execCommand('underline', false, null); document.getElementById('{}').focus();", rtb_id_underline));
                                                                                },
                                                                                "U"
                                                                            }
                                                                            div { style: "width: 1px; background: #ccc; margin: 0 4px;" }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                                title: "Bullet List",
                                                                                onclick: move |_| {
                                                                                    let _ = document::eval(&format!("document.execCommand('insertUnorderedList', false, null); document.getElementById('{}').focus();", rtb_id_ul));
                                                                                },
                                                                                "• List"
                                                                            }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                                title: "Numbered List",
                                                                                onclick: move |_| {
                                                                                    let _ = document::eval(&format!("document.execCommand('insertOrderedList', false, null); document.getElementById('{}').focus();", rtb_id_ol));
                                                                                },
                                                                                "1. List"
                                                                            }
                                                                        }
                                                                    }
                                                                    div {
                                                                        id: "{rtb_id}",
                                                                        contenteditable: if is_enabled { "true" } else { "false" },
                                                                        style: "flex: 1; padding: 8px; overflow: auto; outline: none; background: white; {style_back} {style_font} {style_fore};",
                                                                        dangerous_inner_html: "{html}",
                                                                        onmounted: move |_| {
                                                                            let _ = document::eval(&format!(r#"
                                                                                (function() {{
                                                                                    const editor = document.getElementById('{}');
                                                                                    if (editor) {{
                                                                                        editor.addEventListener('keydown', function(e) {{
                                                                                            if (e.ctrlKey || e.metaKey) {{
                                                                                                switch(e.key.toLowerCase()) {{
                                                                                                    case 'b':
                                                                                                        e.preventDefault();
                                                                                                        document.execCommand('bold', false, null);
                                                                                                        break;
                                                                                                    case 'i':
                                                                                                        e.preventDefault();
                                                                                                        document.execCommand('italic', false, null);
                                                                                                        break;
                                                                                                    case 'u':
                                                                                                        e.preventDefault();
                                                                                                        document.execCommand('underline', false, null);
                                                                                                        break;
                                                                                                }}
                                                                                            }}
                                                                                        }});
                                                                                    }}
                                                                                }})();
                                                                            "#, rtb_id_mount));
                                                                        },
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::WebBrowser => rsx! {
                                                        {
                                                            let url = control.properties.get_string("URL").map(|s| s.to_string()).unwrap_or_else(|| "about:blank".to_string());
                                                            rsx! {
                                                                iframe {
                                                                    id: "{name_clone}",
                                                                    style: "width: 100%; height: 100%; border: 1px inset #999; background: white; {style_back};",
                                                                    src: "{url}",
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::ListView => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px inset #999; background: white; overflow: auto; {style_back} {style_font} {style_fore};",
                                                            table {
                                                                style: "width: 100%; border-collapse: collapse; font-size: 12px; {style_font}; {style_fore};",
                                                                thead {
                                                                    tr {
                                                                        th { style: "border: 1px solid #ccc; padding: 2px; background: #f0f0f0; text-align: left; {style_fore}; {style_font};", "ColumnHeader" }
                                                                    }
                                                                }
                                                                tbody {
                                                                    tr {
                                                                        td { style: "border: 1px solid #eee; padding: 2px; {style_font}; {style_fore};", "ListItem" }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::TreeView => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px inset #999; background: white; overflow: auto; padding: 4px; {style_back} {style_font} {style_fore};",
                                                            div { style: "{style_font} {style_fore};", "Node0" }
                                                            div { style: "padding-left: 20px; {style_font} {style_fore};", "Node1" }
                                                        }
                                                    },
                                                    ControlType::DataGridView => rsx! {
                                                        {
                                                            // Get dynamic grid data from properties (set by DataSourceChanged)
                                                            let grid_columns = control.properties.get_string_array("__grid_columns")
                                                                .cloned()
                                                                .unwrap_or_default();
                                                            let grid_row_strs = control.properties.get_string_array("__grid_rows")
                                                                .cloned()
                                                                .unwrap_or_default();
                                                            let grid_rows: Vec<Vec<String>> = grid_row_strs.iter()
                                                                .map(|s| s.split('\t').map(|c| c.to_string()).collect())
                                                                .collect();
                                                            let has_data = !grid_columns.is_empty();

                                                            rsx! {
                                                                div {
                                                                    style: "width: 100%; height: 100%; border: 1px solid #999; background: #f0f0f0; padding: 1px; overflow: auto;",
                                                                    table {
                                                                        style: "width: 100%; background: white; border-collapse: separate; border-spacing: 0; font-size: 12px;",
                                                                        thead {
                                                                            tr {
                                                                                th { style: "background: #e8e8e8; border-right: 1px solid #999; border-bottom: 2px solid #999; padding: 4px 6px; width: 30px; text-align: center; font-weight: normal; color: #333;", "" }
                                                                                if has_data {
                                                                                    for col in &grid_columns {
                                                                                        th { style: "background: #e8e8e8; border-right: 1px solid #ccc; border-bottom: 2px solid #999; padding: 4px 8px; text-align: left; font-weight: bold; color: #222; cursor: default; white-space: nowrap;", "{col}" }
                                                                                    }
                                                                                } else {
                                                                                    th { style: "background: #e8e8e8; border-right: 1px solid #ccc; border-bottom: 2px solid #999; padding: 4px 8px;", "Column1" }
                                                                                    th { style: "background: #e8e8e8; border-right: 1px solid #ccc; border-bottom: 2px solid #999; padding: 4px 8px;", "Column2" }
                                                                                }
                                                                            }
                                                                        }
                                                                        tbody {
                                                                            if has_data {
                                                                                for (ri, row) in grid_rows.iter().enumerate() {
                                                                                    tr {
                                                                                        td { style: "background: #e8e8e8; border-right: 1px solid #999; border-bottom: 1px solid #ddd; text-align: center; padding: 2px 4px; color: #333; width: 30px; height: 22px;", "{ri}" }
                                                                                        for cell in row {
                                                                                            td { style: "border-right: 1px solid #eee; border-bottom: 1px solid #eee; padding: 3px 6px; white-space: nowrap; height: 22px;", "{cell}" }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            } else {
                                                                                tr {
                                                                                    td { style: "background: #e8e8e8; border-right: 1px solid #999; border-bottom: 1px solid #ddd; text-align: center; padding: 2px 4px; width: 30px; height: 22px;", "" }
                                                                                    td { style: "border-right: 1px solid #eee; border-bottom: 1px solid #eee; padding: 3px 6px; height: 22px;", "" }
                                                                                    td { style: "border-right: 1px solid #eee; border-bottom: 1px solid #eee; padding: 3px 6px; height: 22px;", "" }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::BindingNavigator => rsx! {
                                                        {
                                                            let nav_first = name.clone();
                                                            let nav_prev = name.clone();
                                                            let nav_next = name.clone();
                                                            let nav_last = name.clone();
                                                            let nav_add = name.clone();
                                                            let nav_del = name.clone();
                                                            rsx! {
                                                                div {
                                                                    style: "width: 100%; height: 100%; display: flex; align-items: center; gap: 2px; background: #f0f0f0; border: 1px solid #ccc; padding: 2px 4px; font-size: 11px;",
                                                                    button {
                                                                        style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                                        title: "Move first",
                                                                        onclick: move |_| handle_event(nav_first.clone(), "MoveFirst".to_string(), None),
                                                                        "⏮"
                                                                    }
                                                                    button {
                                                                        style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                                        title: "Move previous",
                                                                        onclick: move |_| handle_event(nav_prev.clone(), "MovePrevious".to_string(), None),
                                                                        "◀"
                                                                    }
                                                                    span {
                                                                        style: "padding: 0 4px; min-width: 40px; text-align: center; border: 1px solid #ccc; background: white;",
                                                                        "{text}"
                                                                    }
                                                                    button {
                                                                        style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                                        title: "Move next",
                                                                        onclick: move |_| handle_event(nav_next.clone(), "MoveNext".to_string(), None),
                                                                        "▶"
                                                                    }
                                                                    button {
                                                                        style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                                        title: "Move last",
                                                                        onclick: move |_| handle_event(nav_last.clone(), "MoveLast".to_string(), None),
                                                                        "⏭"
                                                                    }
                                                                    div { style: "width: 1px; height: 16px; background: #aaa; margin: 0 2px;" }
                                                                    button {
                                                                        style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                                        title: "Add new",
                                                                        onclick: move |_| handle_event(nav_add.clone(), "AddNew".to_string(), None),
                                                                        "➕"
                                                                    }
                                                                    button {
                                                                        style: "padding: 1px 6px; border: 1px solid #aaa; background: #e8e8e8; cursor: pointer; font-size: 11px;",
                                                                        title: "Delete",
                                                                        onclick: move |_| handle_event(nav_del.clone(), "Delete".to_string(), None),
                                                                        "❌"
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::Panel => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px solid #ccc; overflow: hidden; {style_back} {style_font} {style_fore};",
                                                            onclick: move |evt: MouseEvent| {
                                                                let data = irys_runtime::EventData::Mouse {
                                                                    button: 0x100000, clicks: 1,
                                                                    x: evt.client_coordinates().x as i32,
                                                                    y: evt.client_coordinates().y as i32,
                                                                    delta: 0,
                                                                };
                                                                handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                                            },
                                                        }
                                                    },
                                                    ControlType::PictureBox => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px solid #e2e8f0; overflow: hidden; {style_back};",
                                                            onclick: move |evt: MouseEvent| {
                                                                let data = irys_runtime::EventData::Mouse {
                                                                    button: 0x100000, clicks: 1,
                                                                    x: evt.client_coordinates().x as i32,
                                                                    y: evt.client_coordinates().y as i32,
                                                                    delta: 0,
                                                                };
                                                                handle_event(name_clone.clone(), "Click".to_string(), Some(data));
                                                            },
                                                        }
                                                    },
                                                    ControlType::TabControl => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px solid #adb5bd; display: flex; flex-direction: column;",
                                                            div {
                                                                style: "display: flex; background: #e9ecef; border-bottom: 1px solid #adb5bd;",
                                                                div {
                                                                    style: "padding: 4px 12px; background: white; border: 1px solid #adb5bd; border-bottom: none; cursor: pointer; font-size: 12px;",
                                                                    "Tab 1"
                                                                }
                                                            }
                                                            div {
                                                                style: "flex: 1; padding: 8px; background: white;",
                                                            }
                                                        }
                                                    },
                                                    ControlType::ProgressBar => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; background: #e9ecef; border: 1px solid #adb5bd; overflow: hidden;",
                                                            div {
                                                                style: "height: 100%; background: #0d6efd; width: 0%; transition: width 0.3s;",
                                                            }
                                                        }
                                                    },
                                                    ControlType::NumericUpDown => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; display: flex; border: 1px solid #adb5bd;",
                                                            input {
                                                                r#type: "number",
                                                                style: "flex: 1; border: none; padding: 2px 4px; font-size: 12px; outline: none;",
                                                                value: "0",
                                                            }
                                                        }
                                                    },
                                                    ControlType::MenuStrip | ControlType::ContextMenuStrip => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; background: #f0f0f0; border-bottom: 1px solid #ccc; display: flex; align-items: center; padding: 0 4px; font-size: 12px;",
                                                            "Menu"
                                                        }
                                                    },
                                                    ControlType::StatusStrip => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; background: #007acc; border-top: 1px solid #005a9e; display: flex; align-items: center; padding: 0 8px; font-size: 12px; color: white;",
                                                            "{text}"
                                                        }
                                                    },
                                                    ControlType::DateTimePicker => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; display: flex; align-items: center; border: 1px solid #adb5bd; background: white; padding: 2px 4px; font-size: 12px;",
                                                            span { style: "flex: 1;", "{text}" }
                                                            span { style: "padding: 0 4px; border-left: 1px solid #ccc; cursor: pointer;", "▼" }
                                                        }
                                                    },
                                                    ControlType::LinkLabel => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; display: flex; align-items: center; font-size: 12px; color: #0066cc; text-decoration: underline; cursor: pointer;",
                                                            "{caption}"
                                                        }
                                                    },
                                                    ControlType::ToolStrip => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; background: #f0f0f0; border-bottom: 1px solid #ccc; display: flex; align-items: center; padding: 0 4px; font-size: 12px;",
                                                            "ToolStrip"
                                                        }
                                                    },
                                                    ControlType::TrackBar => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; display: flex; align-items: center; padding: 4px;",
                                                            input {
                                                                r#type: "range",
                                                                style: "width: 100%;",
                                                                min: "0",
                                                                max: "10",
                                                                value: "0",
                                                            }
                                                        }
                                                    },
                                                    ControlType::MaskedTextBox => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; display: flex; align-items: center;",
                                                            input {
                                                                r#type: "text",
                                                                style: "width: 100%; height: 100%; border: 1px solid #adb5bd; padding: 2px 4px; font-size: 12px; outline: none; {style_back} {style_font} {style_fore};",
                                                                value: "{text}",
                                                                placeholder: "___-__-____",
                                                            }
                                                        }
                                                    },
                                                    ControlType::SplitContainer => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; display: flex; border: 1px solid #adb5bd;",
                                                            div { style: "flex: 1; background: #f8f8f8; border-right: 3px solid #ccc;" }
                                                            div { style: "flex: 1; background: #f8f8f8;" }
                                                        }
                                                    },
                                                    ControlType::FlowLayoutPanel => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px dashed #adb5bd; display: flex; flex-wrap: wrap; align-content: flex-start; padding: 2px;",
                                                        }
                                                    },
                                                    ControlType::TableLayoutPanel => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px dashed #adb5bd; display: grid; grid-template-columns: 1fr 1fr; grid-template-rows: 1fr 1fr;",
                                                            div { style: "border: 1px dotted #ccc;" }
                                                            div { style: "border: 1px dotted #ccc;" }
                                                            div { style: "border: 1px dotted #ccc;" }
                                                            div { style: "border: 1px dotted #ccc;" }
                                                        }
                                                    },
                                                    ControlType::MonthCalendar => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px solid #adb5bd; background: white; display: flex; flex-direction: column; font-size: 11px;",
                                                            div {
                                                                style: "background: #0078d4; color: white; text-align: center; padding: 4px; font-weight: bold;",
                                                                "January 2026"
                                                            }
                                                            div {
                                                                style: "flex: 1; display: grid; grid-template-columns: repeat(7, 1fr); text-align: center; padding: 2px; gap: 1px;",
                                                                span { style: "font-weight: bold;", "Su" }
                                                                span { style: "font-weight: bold;", "Mo" }
                                                                span { style: "font-weight: bold;", "Tu" }
                                                                span { style: "font-weight: bold;", "We" }
                                                                span { style: "font-weight: bold;", "Th" }
                                                                span { style: "font-weight: bold;", "Fr" }
                                                                span { style: "font-weight: bold;", "Sa" }
                                                            }
                                                        }
                                                    },
                                                    ControlType::HScrollBar => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; display: flex; align-items: center;",
                                                            input {
                                                                r#type: "range",
                                                                style: "width: 100%; height: 100%;",
                                                                min: "0",
                                                                max: "100",
                                                                value: "0",
                                                            }
                                                        }
                                                    },
                                                    ControlType::VScrollBar => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; display: flex; align-items: center;",
                                                            input {
                                                                r#type: "range",
                                                                style: "width: 17px; height: 100%; writing-mode: vertical-lr; direction: rtl;",
                                                                min: "0",
                                                                max: "100",
                                                                value: "0",
                                                            }
                                                        }
                                                    },
                                                    _ => rsx! {
                                                        div { style: "border: 1px dotted red;", "Unknown Control" }
                                                    }
                                                }}
                                            }
                                        }
                                    }
                                }
                            }
                            // MsgBox overlay
                            {if msgbox_visible {
                                rsx! {
                                    div {
                                        style: "
                                            position: absolute;
                                            top: 0; left: 0; right: 0; bottom: 0;
                                            background: rgba(0,0,0,0.5);
                                            display: flex;
                                            align-items: center;
                                            justify-content: center;
                                            z-index: 1000;
                                        ",
                                        div {
                                            style: "
                                                background: #f0f0f0;
                                                border: 1px solid #999;
                                                box-shadow: 0 4px 16px rgba(0,0,0,0.5);
                                                min-width: 200px;
                                                display: flex;
                                                flex-direction: column;
                                            ",
                                            div {
                                                style: "background: #0078d4; color: white; padding: 4px 8px; font-weight: bold;",
                                                "Project1"
                                            }
                                            div {
                                                style: "padding: 20px; text-align: center; color: black;",
                                                "{msgbox_text}"
                                            }
                                            div {
                                                style: "padding: 10px; display: flex; justify-content: center;",
                                                button {
                                                    style: "padding: 4px 20px;",
                                                    onclick: move |_| msgbox_content.set(None),
                                                    "OK"
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                rsx! {}
                            }}
                        }
                    }
                } else {
                    rsx! { div { "Loading..." } }
                }
            }
        }
    }
}
