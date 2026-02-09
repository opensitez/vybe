use dioxus::prelude::*;
use crate::app_state::AppState;
use irys_forms::{ControlType, Form, EventType};
use irys_runtime::{Interpreter, RuntimeSideEffect, Value, ObjectData};
use std::collections::HashMap;
use std::cell::RefCell;
use std::rc::Rc;
use irys_parser::parse_program;

fn event_type_from_name(name: &str) -> Option<EventType> {
    match name.to_lowercase().as_str() {
        "click" => Some(EventType::Click),
        "dblclick" => Some(EventType::DblClick),
        "load" => Some(EventType::Load),
        "unload" => Some(EventType::Unload),
        "change" => Some(EventType::Change),
        "keypress" => Some(EventType::KeyPress),
        "keydown" => Some(EventType::KeyDown),
        "keyup" => Some(EventType::KeyUp),
        "mousedown" => Some(EventType::MouseDown),
        "mouseup" => Some(EventType::MouseUp),
        "mousemove" => Some(EventType::MouseMove),
        "gotfocus" => Some(EventType::GotFocus),
        "lostfocus" => Some(EventType::LostFocus),
        _ => None,
    }
}

#[component]
pub fn RuntimePanel() -> Element {
    let state = use_context::<AppState>();
    
    // We need to hold the interpreter instance
    let mut interpreter = use_signal(|| None::<Interpreter>);
    
    // We need a local copy of the form to act as the runtime state
    let mut runtime_form = use_signal(|| None::<Form>);

    let mut msgbox_content = use_signal(|| None::<String>);
    let mut parse_error = use_signal(|| None::<String>);

    // Helper to process side effects (defined before use)
    // We can't easily capture signals in a closure and reuse it because of move semantics and multiple mutable borrows of signals
    // So we'll define a macro or just a function that takes the signals as arguments?
    // A local function is fine if we pass everything.
    
    fn process_side_effects(
        interp: &mut Interpreter, 
        state: AppState, 
        runtime_form: &mut Signal<Option<Form>>, 
        msgbox_content: &mut Signal<Option<String>>
    ) {
         while let Some(effect) = interp.side_effects.pop_front() {
             match effect {
                RuntimeSideEffect::MsgBox(msg) => {
                     msgbox_content.set(Some(msg));
                }
                RuntimeSideEffect::PropertyChange { object, property, value } => {
                    let mut switched = false;
                    
                    let project_read = state.project.read();
                    if let Some(proj) = project_read.as_ref() {
                        if let Some(other_form_module) = proj.forms.iter().find(|f| f.form.name.eq_ignore_ascii_case(&object)) {
                             let current_is_it = if let Some(cf) = &*runtime_form.peek() {
                                 cf.name.eq_ignore_ascii_case(&object)
                             } else { false };

                             if !current_is_it {
                                 if property.eq_ignore_ascii_case("Visible") && (value.as_bool().unwrap_or(false) || value.as_string() == "True") {
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
                                                  // Try to call Load event (best effort)
                                                  // We intentionally ignore error if not found
                                                  let _ = interp.call_event_handler(&format!("{}_Load", other_form_module.form.name), &[]);
                                                  
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
                            // Handle qualified names (e.g. Form1.Label1 or just Label1)
                            // In VB6, you can reverence controls on the current form directly or fully qualified
                            
                            let (form_part, control_part) = if let Some(idx) = object.find('.') {
                                (&object[..idx], &object[idx+1..])
                            } else {
                                ("", object.as_str())
                            };

                            // Only update if it's for this form or unqualified
                            let is_for_this_form = form_part.is_empty() || form_part.eq_ignore_ascii_case(&frm.name);
                            
                            if is_for_this_form {
                                // check if it is the form itself
                                if control_part.eq_ignore_ascii_case(&frm.name) || (form_part.is_empty() && object.eq_ignore_ascii_case(&frm.name)) {
                                    if property.eq_ignore_ascii_case("Caption") {
                                        frm.caption = value.as_string();
                                    }
                                } else {
                                    if let Some(ctrl) = frm.get_control_by_name_mut(control_part) {
                                        match property.to_lowercase().as_str() {
                                            "text" => {
                                                let text_val = value.as_string();
                                                ctrl.set_text(text_val.clone());
                                                // If this is a ListBox, also update its List items from the text (pipe-delimited)
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
                                            },
                                            "caption" => ctrl.set_caption(value.as_string()),
                                            "backcolor" => ctrl.set_back_color(value.as_string()),
                                            "forecolor" => ctrl.set_fore_color(value.as_string()),
                                            "font" => ctrl.set_font(value.as_string()),
                                            "url" => {
                                                // Update the URL property in the control
                                                ctrl.properties.set("URL", value.as_string());
                                                // Update the iframe src via JavaScript
                                                let url = value.as_string();
                                                let _ = eval(&format!(
                                                    r#"
                                                    const iframe = document.getElementById('{}');
                                                    if (iframe) {{
                                                        iframe.src = '{}';
                                                    }}
                                                    "#,
                                                    control_part, url
                                                ));
                                            },
                                            "html" => {
                                                // Update the HTML property in the control
                                                ctrl.properties.set("HTML", value.as_string());
                                                // Update the contenteditable div innerHTML via JavaScript
                                                let html = value.as_string();
                                                let rtb_id = format!("rtb_{}", control_part);
                                                let _ = eval(&format!(
                                                    r#"
                                                    const editor = document.getElementById('{}');
                                                    if (editor) {{
                                                        editor.innerHTML = '{}';
                                                    }}
                                                    "#,
                                                    rtb_id, html.replace("'", "\\'").replace("\n", "\\n")
                                                ));
                                            },
                                            _ => {
                                                // Generic property setter
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
                    println!("[Runtime Console] {}", msg);
                }
                RuntimeSideEffect::ConsoleClear => {
                    println!("[Runtime Console] Cleared");
                }
            }
            }

        // Sync Step: Poll interpreter for control state changes to reflect code updates in UI
        if let Some(frm) = runtime_form.write().as_mut() {
            let instance_name = format!("{}Instance", frm.name);
            
            // Try to find the form instance in the environment
            if let Ok(irys_runtime::Value::Object(form_obj)) = interp.env.get(&instance_name) {
                 let form_borrow = form_obj.borrow();
                 
                 for control in frm.controls.iter_mut() {
                     // Look for control object in form instance fields (case-insensitive lookup ideally, but fields are usually lowercase in runtime)
                     if let Some(irys_runtime::Value::Object(ctrl_obj)) = form_borrow.fields.get(&control.name.to_lowercase()) {
                          let ctrl_fields = &ctrl_obj.borrow().fields;
                          
                          // Sync Caption -> Text/Caption
                          if let Some(irys_runtime::Value::String(s)) = ctrl_fields.get("caption") {
                              control.properties.set_raw("Caption", irys_forms::PropertyValue::String(s.clone()));
                              // Also sync to Text for compatibility if Text is not explicitly set differently? 
                              // Actually specific controls use one or the other.
                              control.properties.set_raw("Text", irys_forms::PropertyValue::String(s.clone()));
                          }
                          
                          // Sync Text
                          if let Some(irys_runtime::Value::String(s)) = ctrl_fields.get("text") {
                              control.properties.set_raw("Text", irys_forms::PropertyValue::String(s.clone()));
                          }
                          
                          // Sync Enabled
                          if let Some(val) = ctrl_fields.get("enabled") {
                               match val {
                                   irys_runtime::Value::Boolean(b) => { control.properties.set_raw("Enabled", irys_forms::PropertyValue::Boolean(*b)); },
                                   irys_runtime::Value::Integer(i) => { control.properties.set_raw("Enabled", irys_forms::PropertyValue::Boolean(*i != 0)); },
                                   _ => {}
                               }
                          }

                           // Sync Visible
                           if let Some(val) = ctrl_fields.get("visible") {
                               match val {
                                   irys_runtime::Value::Boolean(b) => { control.properties.set_raw("Visible", irys_forms::PropertyValue::Boolean(*b)); },
                                   irys_runtime::Value::Integer(i) => { control.properties.set_raw("Visible", irys_forms::PropertyValue::Boolean(*i != 0)); },
                                   _ => {}
                               }
                          }
                     }
                 }
            }
        }
    }

    // Initialize Runtime
    use_effect(move || {
        if interpreter.read().is_none() {
            // Get the startup form from the project, not the currently selected item
            let project_read = state.project.read();
            if let Some(proj) = project_read.as_ref() {
                if let Some(startup_form_module) = proj.get_startup_form() {
                    let form = startup_form_module.form.clone();
                    // For VB.NET forms, combine designer + user code
                    let form_code = if startup_form_module.is_vbnet() {
                        format!("{}\n{}", startup_form_module.get_designer_code(), startup_form_module.get_user_code())
                    } else {
                        startup_form_module.get_user_code().to_string()
                    };
                    let is_vbnet_form = startup_form_module.is_vbnet();
                    drop(project_read); // Release lock before using form
                    
                    runtime_form.set(Some(form.clone()));
                    
                    let mut interp = Interpreter::new();
                    
                    // Register resources
                    let mut res_map = HashMap::new();
                    if let Some(proj) = state.project.read().as_ref() {
                         for item in &proj.resources.resources {
                             res_map.insert(item.name.clone(), item.value.clone());
                         }
                    }
                    interp.register_resources(res_map);
                    
                    // Load all code files (global scope + class definitions)
                    let project_read = state.project.read();
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
                                // Register form controls as variables for global scope access
                                if let Some(runtime_form_data) = runtime_form.read().as_ref() {
                                    for control in &runtime_form_data.controls {
                                        let control_var = irys_runtime::Value::String(control.name.clone());
                                        interp.env.define(&control.name, control_var);
                                    }
                                }

                                // For VB.NET, create an instance of the form
                                if is_vbnet_form {
                                    // Generate RuntimeControls module with generic Control class
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

                                    let global_module = format!(r#"
                                        Module RuntimeGlobals
                                            Public {}Instance As New {}
                                            
                                            Public Sub InitControls()
                                                ' Auto-generated control initialization
                                                {}
                                            End Sub
                                        End Module
                                    "#, form.name, form.name, 
                                        if let Some(runtime_form_data) = runtime_form.read().as_ref() {
                                            runtime_form_data.controls.iter().map(|c| {
                                                let cap = c.get_caption().map(|s| s.to_string()).unwrap_or_else(|| c.name.clone());
                                                let en = c.is_enabled();
                                                format!(
                                                    "RuntimeGlobals.{0}Instance.{1} = New RuntimeControls.Control\nRuntimeGlobals.{0}Instance.{1}.Name = \"{1}\"\nRuntimeGlobals.{0}Instance.{1}.Caption = \"{2}\"\nRuntimeGlobals.{0}Instance.{1}.Enabled = {3}\n", 
                                                    form.name, c.name, cap.replace("\"", "\"\""), if en {"True"} else {"False"}
                                                )
                                            }).collect::<Vec<_>>().join("\n")
                                        } else { String::new() }
                                    );
                                    if let Ok(prog) = parse_program(&global_module) {
                                        if let Err(e) = interp.load_module("RuntimeGlobals", &prog) {
                                             println!("Error creating runtime globals: {:?}", e);
                                        } else {
                                            // Initialize controls immediately
                                           let _ = interp.call_procedure(&irys_parser::ast::Identifier::new("InitControls"), &[]);
                                        }
                                    }

                                    // Use Handles clause to find the Load handler
                                    let instance_name = format!("RuntimeGlobals.{}Instance", form.name);
                                    let form_name_lower = form.name.to_lowercase();

                                    // Auto-call InitializeComponent on the instance to wire controls
                                    match interp.call_instance_method(&instance_name, "InitializeComponent", &[]) {
                                        Ok(_) => {
                                            if let Some(Value::Object(form_obj)) = interp.env.get_global(&instance_name) {
                                                let mut form_ref = form_obj.borrow_mut();
                                                if let Some(runtime_form_data) = runtime_form.read().as_ref() {
                                                    for ctrl in &runtime_form_data.controls {
                                                        let key = ctrl.name.to_lowercase();
                                                        let needs_create = matches!(form_ref.fields.get(&key), None | Some(Value::Nothing) | Some(Value::String(_)));
                                                        if needs_create {
                                                            let mut ctrl_fields = HashMap::new();
                                                            ctrl_fields.insert("name".to_string(), Value::String(ctrl.name.clone()));
                                                            let caption = ctrl.get_caption().map(|s| s.to_string()).unwrap_or_else(|| ctrl.name.clone());
                                                            ctrl_fields.insert("caption".to_string(), Value::String(caption));
                                                            let text = ctrl.get_text().map(|s| s.to_string()).unwrap_or_default();
                                                            ctrl_fields.insert("text".to_string(), Value::String(text));
                                                            ctrl_fields.insert("enabled".to_string(), Value::Boolean(ctrl.is_enabled()));
                                                            let ctrl_obj = Value::Object(Rc::new(RefCell::new(ObjectData { class_name: "RuntimeControls.Control".to_string(), fields: ctrl_fields })));
                                                            form_ref.fields.insert(key.clone(), ctrl_obj.clone());
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        Err(e) => println!("InitializeComponent error: {:?}", e),
                                    }
                                    
                                    // Look for method with Handles Me.Load
                                    if let Some(load_handler) = interp.find_handles_method(&form_name_lower, "Me", "Load") {
                                        let _ = interp.call_instance_method(&instance_name, &load_handler, &[]);
                                    } else {
                                        // Fallback: try convention-based names
                                        let _ = interp.call_instance_method(&instance_name, &format!("{}_Load", form.name), &[]);
                                        let _ = interp.call_instance_method(&instance_name, "Form_Load", &[]);
                                    }
                                
                                    // Also try standard Form_Load as global sub for compatibility
                                    let _ = interp.call_event_handler("Form_Load", &[]);
                                } else {
                                    match interp.call_event_handler("Form_Load", &[]) {
                                        Ok(_) => {},
                                        Err(irys_runtime::RuntimeError::UndefinedFunction(_)) => {},
                                        Err(e) => println!("Form_Load Error: {:?}", e)
                                    }
                                }
                                process_side_effects(&mut interp, state, &mut runtime_form, &mut msgbox_content);
                            }
                        },
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

    // Helper to run events
    // Helper to run events
    let mut handle_event = move |control_name: String, event_name: String| {
         if parse_error.read().is_some() { return; }
         if msgbox_content.read().is_some() { return; }

        if let Some(interp) = interpreter.write().as_mut() {
            if let Some(frm) = runtime_form.read().as_ref() {
                // Sync current form/control state into environment
                // For VB.NET, we try to sync to the instance properties if possible
                let is_vbnet = state.project.read().as_ref()
                    .and_then(|p| p.get_startup_form())
                    .map(|f| f.is_vbnet())
                    .unwrap_or(false);

                // Common global sync (legacy/compatibility)
                interp.env.set(&format!("{}.Caption", frm.name), irys_runtime::Value::String(frm.caption.clone())).ok();
                for ctrl in &frm.controls {
                    let ctrl_name = ctrl.name.clone();
                    let caption = ctrl.get_caption().map(|s| s.to_string()).unwrap_or(ctrl_name.clone());
                    let text = ctrl.get_text().map(|s| s.to_string()).unwrap_or_default();
                    let val = ctrl.properties.get_int("Value").unwrap_or(0);
                    let _enabled = ctrl.is_enabled();
                    // ... other properties

                    interp.env.set(&format!("{}.Caption", ctrl_name), irys_runtime::Value::String(caption.clone())).ok();
                    interp.env.set(&format!("{}.Text", ctrl_name), irys_runtime::Value::String(text.clone())).ok();
                    interp.env.set(&format!("{}.Value", ctrl_name), irys_runtime::Value::Integer(val)).ok();
                    
                    if is_vbnet {
                        // Sync to RuntimeGlobals.FormInstance.Control.Prop
                         let instance_name = format!("RuntimeGlobals.{}Instance", frm.name);
                         let sync_script = format!(r#"
                            {}.{}.Caption = "{}"
                            {}.{}.Text = "{}"
                         "#, 
                            instance_name, ctrl_name, caption.replace("\"", "\"\""),
                            instance_name, ctrl_name, text.replace("\"", "\"\"")
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
                        // Form-level event
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
            
            // 1. If VB.NET, first try finding handler via Handles clause or name convention ON THE INSTANCE
            let is_vbnet = state.project.read().as_ref()
               .and_then(|p| p.get_startup_form())
               .map(|f| f.is_vbnet())
               .unwrap_or(false);
               
            if is_vbnet {
                if let Some(frm) = runtime_form.read().as_ref() {
                    let instance_name = format!("RuntimeGlobals.{}Instance", frm.name);
                    let form_class = frm.name.to_lowercase();
                    
                    // First: find handler via Handles clause (e.g. Handles btn0.Click)
                    if let Some(method_name) = interp.find_handles_method(&form_class, &control_name, &event_name) {
                        match interp.call_instance_method(&instance_name, &method_name, &[]) {
                            Ok(_) => executed = true,
                            Err(_) => {},
                        }
                    }
                    
                    // Fallback: try convention-based handler name as instance method
                    if !executed {
                        match interp.call_instance_method(&instance_name, &handler_name, &[]) {
                            Ok(_) => executed = true,
                            Err(_) => {},
                        }
                    }
                    // Form-level: also try Form_Load / Form_Click on instance when target is form
                    if !executed && control_name.eq_ignore_ascii_case(&frm.name) {
                        let conv_name = format!("{}_{}", frm.name, event_name);
                        match interp.call_instance_method(&instance_name, &conv_name, &[]) {
                            Ok(_) => executed = true,
                            Err(_) => {},
                        }
                    }
                }
            }

            // 2. Fallback to global sub/function (VB6 / Module style)
            if !executed {
                let handler_key = interp.subs.keys().find(|key| {
                    *key == &handler_lower || key.ends_with(&format!(".{}", handler_lower))
                }).cloned();

                if let Some(key) = handler_key {
                        let _ = interp.call_event_handler(&key, &[]);
                }
            }

            // Suppress noisy missing-handler logs for unwired events
            process_side_effects(interp, state, &mut runtime_form, &mut msgbox_content);
        }
    };

    let form_opt = runtime_form.read().clone();
    let msgbox_visible = msgbox_content.read().is_some();
    let msgbox_text = msgbox_content.read().clone().unwrap_or_default();
    let error_text = parse_error.read().clone();

    rsx! {
        div {
            class: "runtime-panel",
            style: "flex: 1; background: #e0e0e0; display: flex; align-items: center; justify-content: center; overflow: auto; position: relative;",
            
            {
                if let Some(err) = error_text {
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
                                "{caption}"
                            }
                            
                            // Controls
                            for control in form.controls {
                                {
                                    let control_type = control.control_type;
                                    let x = control.bounds.x;
                                    let y = control.bounds.y;
                                    let w = control.bounds.width;
                                    let h = control.bounds.height;
                                    let name = control.name.clone();
                                    let name_clone = name.clone(); // For closure
                                    
                                    // Helper to get text/caption
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
                                    if let Some(fc) = &fore_color { style_fore = format!("color: {};", fc); }
                                    let mut style_back = String::new();
                                    if let Some(bc) = &back_color { style_back = format!("background: {};", bc); }
                                    let button_bg = if let Some(bc) = &back_color {
                                        let color_part = if let Some(fc) = &fore_color { format!("color: {};", fc) } else { String::new() };
                                        format!("background: {}; {}; border: 1px solid #cbd5e1;", bc, color_part)
                                    } else {
                                        base_button_bg.to_string()
                                    };

                                    rsx! {
                                        if is_visible {
                                            div {
                                                style: "position: absolute; left: {x}px; top: {y}px; width: {w}px; height: {h}px;",
                                                
                                                {match control_type {
                                                    ControlType::Button => rsx! {
                                                        button {
                                                            style: "width: 100%; height: 100%; padding: 6px 10px; {button_bg} {style_font}; border-radius: 6px; box-shadow: 0 2px 4px rgba(37,99,235,0.12);",
                                                            disabled: !is_enabled,
                                                            onclick: move |_| handle_event(name_clone.clone(), "Click".to_string()),
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
                                                                 handle_event(name_clone.clone(), "Change".to_string());
                                                            }
                                                        }
                                                    },
                                                    ControlType::Label => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; padding: 4px 2px; {style_font} {style_fore} {style_back};",
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
                                                                onclick: move |_| handle_event(name_clone.clone(), "Click".to_string())
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
                                                                onclick: move |_| handle_event(name_clone.clone(), "Click".to_string())
                                                            }
                                                            span { "{caption}" }
                                                        }
                                                    },
                                                    ControlType::Frame => rsx! {
                                                        fieldset {
                                                            style: "width: 100%; height: 100%; {base_frame_border} margin: 0; padding: 0; border-radius: 8px; {style_back} {style_font} {style_fore};",
                                                            legend { "{caption}" }
                                                        }
                                                    },
                                                    ControlType::ListBox => rsx! {
                                                        select {
                                                            style: "width: 100%; height: 100%; border: 1px solid #cbd5e1; border-radius: 8px; {base_field_bg} {style_back} {style_font} {style_fore};",
                                                            multiple: true,
                                                            disabled: !is_enabled,
                                                            onchange: move |_| handle_event(name_clone.clone(), "Click".to_string()),
                                                            {
                                                                // Prefer the List property; if empty, fall back to Text split by '|'
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
                                                                handle_event(name_clone.clone(), "Change".to_string());
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
                                                            // Use HTML property if set, otherwise fall back to Text
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
                                                                    // Toolbar (conditional)
                                                                    if toolbar_visible {
                                                                        div {
                                                                            style: "display: flex; gap: 2px; padding: 4px; background: #f0f0f0; border-bottom: 1px solid #ccc;",
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; font-weight: bold;",
                                                                                title: "Bold (Ctrl+B)",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('bold', false, null); document.getElementById('{}').focus();", rtb_id_bold));
                                                                                },
                                                                                "B"
                                                                            }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; font-style: italic;",
                                                                                title: "Italic (Ctrl+I)",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('italic', false, null); document.getElementById('{}').focus();", rtb_id_italic));
                                                                                },
                                                                                "I"
                                                                            }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer; text-decoration: underline;",
                                                                                title: "Underline (Ctrl+U)",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('underline', false, null); document.getElementById('{}').focus();", rtb_id_underline));
                                                                                },
                                                                                "U"
                                                                            }
                                                                            div { style: "width: 1px; background: #ccc; margin: 0 4px;" }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                                title: "Bullet List",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('insertUnorderedList', false, null); document.getElementById('{}').focus();", rtb_id_ul));
                                                                                },
                                                                                "• List"
                                                                            }
                                                                            button {
                                                                                style: "padding: 4px 8px; border: 1px solid #999; background: white; cursor: pointer;",
                                                                                title: "Numbered List",
                                                                                onclick: move |_| {
                                                                                    let _ = eval(&format!("document.execCommand('insertOrderedList', false, null); document.getElementById('{}').focus();", rtb_id_ol));
                                                                                },
                                                                                "1. List"
                                                                            }
                                                                        }
                                                                    }
                                                                    // ContentEditable div
                                                                    div {
                                                                        id: "{rtb_id}",
                                                                        contenteditable: if is_enabled { "true" } else { "false" },
                                                                        style: "flex: 1; padding: 8px; overflow: auto; outline: none; background: white; {style_back} {style_font} {style_fore};",
                                                                        dangerous_inner_html: "{html}",
                                                                        onmounted: move |_| {
                                                                            // Add keyboard shortcuts
                                                                            let _ = eval(&format!(r#"
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
                                                            // Placeholder content - in real impl, we'd iterate over Columns and Items
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
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px solid #999; background: #808080; padding: 1px; overflow: auto;",
                                                            table {
                                                                style: "width: 100%; background: white; border-collapse: separate; border-spacing: 0;",
                                                                thead {
                                                                    tr {
                                                                        th { style: "background: #e0e0e0; border-right: 1px solid #999; border-bottom: 1px solid #999; padding: 2px; width: 20px;", "" }
                                                                        th { style: "background: #e0e0e0; border-right: 1px solid #999; border-bottom: 1px solid #999; padding: 2px;", "A" }
                                                                        th { style: "background: #e0e0e0; border-right: 1px solid #999; border-bottom: 1px solid #999; padding: 2px;", "B" }
                                                                    }
                                                                }
                                                                tbody {
                                                                    tr {
                                                                        td { style: "background: #e0e0e0; border-right: 1px solid #999; border-bottom: 1px solid #ddd; text-align: center;", "1" }
                                                                        td { style: "border-right: 1px solid #ddd; border-bottom: 1px solid #ddd; padding: 2px;", "" }
                                                                        td { style: "border-right: 1px solid #ddd; border-bottom: 1px solid #ddd; padding: 2px;", "" }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    },
                                                    ControlType::Panel => rsx! {
                                                        div {
                                                            style: "width: 100%; height: 100%; border: 1px solid #ccc; overflow: hidden; {style_back} {style_font} {style_fore};",
                                                            // We rely on child controls having correct coordinates relative to form (flattened) 
                                                            // OR if we support nesting, we'd need to render children here.
                                                            // Current architecture flattens controls onto the form, so this is just a background.
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
                            // ... MsgBox ...
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
                                            // Title
                                            div {
                                                style: "background: #0078d4; color: white; padding: 4px 8px; font-weight: bold;",
                                                "Project1" // Default title 
                                            }
                                            // Content
                                            div {
                                                style: "padding: 20px; text-align: center; color: black;",
                                                "{msgbox_text}"
                                            }
                                            // Button
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
