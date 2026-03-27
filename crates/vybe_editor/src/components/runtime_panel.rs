use dioxus::prelude::*;
use crate::app_state::AppState;
use std::sync::{Arc, Mutex};

/// Shared state between the VM thread and the UI.
#[derive(Clone)]
struct RunState {
    output: Arc<Mutex<Vec<String>>>,
    done: Arc<Mutex<bool>>,
    error: Arc<Mutex<Option<String>>>,
}

/// Run the project using the bytecode VM directly (no subprocess).
#[component]
pub fn RuntimePanel() -> Element {
    let mut state = use_context::<AppState>();
    let mut display_output = use_signal(String::new);
    let mut is_running = use_signal(|| true);

    let run_state = use_hook(|| {
        let rs = RunState {
            output: Arc::new(Mutex::new(Vec::new())),
            done: Arc::new(Mutex::new(false)),
            error: Arc::new(Mutex::new(None)),
        };

        let project = state.project.read();
        if let Some(proj) = project.as_ref() {
            let has_forms = !proj.forms.is_empty();
            let has_sub_main = proj.code_files.iter()
                .any(|cf| cf.code.to_uppercase().contains("SUB MAIN"));

            // Collect all source code
            let mut all_code = String::new();
            for cf in &proj.code_files {
                all_code.push_str(&cf.code);
                all_code.push('\n');
            }
            for fm in &proj.forms {
                if fm.is_vbnet() {
                    all_code.push_str(&fm.get_designer_code());
                    all_code.push('\n');
                }
                all_code.push_str(&fm.get_user_code());
                all_code.push('\n');
            }

            // For form projects without Sub Main, inject a synthetic entry point
            if has_forms && !has_sub_main {
                let startup_form = proj.get_startup_form()
                    .map(|f| f.form.name.clone())
                    .or_else(|| proj.forms.first().map(|f| f.form.name.clone()))
                    .unwrap_or_else(|| "Form1".to_string());
                all_code.push_str(&format!(
                    "\nModule __EntryPoint\n    Sub Main()\n        Application.Run(\"{}\")\n    End Sub\nEnd Module\n",
                    startup_form
                ));
            }

            // For empty projects (no code, no forms), show a message
            if all_code.trim().is_empty() {
                *rs.error.lock().unwrap() = Some("No code to run. Add a Module with Sub Main() or create a Form.".to_string());
                *rs.done.lock().unwrap() = true;
                return rs;
            }

            let out = rs.output.clone();
            let done = rs.done.clone();
            let err = rs.error.clone();

            std::thread::spawn(move || {
                // Parse
                let program = match vybe_parser_basic::parse_program(&all_code) {
                    Ok(p) => p,
                    Err(e) => {
                        *err.lock().unwrap() = Some(format!("Parse error: {:?}", e));
                        *done.lock().unwrap() = true;
                        return;
                    }
                };

                // Compile
                let chunks = match vybe_compiler_vb::Compiler::new().compile(&program) {
                    Ok(c) => c,
                    Err(e) => {
                        *err.lock().unwrap() = Some(format!("Compile error: {}", e));
                        *done.lock().unwrap() = true;
                        return;
                    }
                };

                // Set up VM with console output capture
                let mut vm = vybe_bytecode::VM::new();
                let queue = std::rc::Rc::new(std::cell::RefCell::new(vybe_host::SideEffectQueue::new()));
                vybe_host::register_all_with_gui(&mut vm, queue.clone());
                vybe_host::setup_namespaces(&mut vm);

                // Override console.log to capture output
                let captured = out.clone();
                vm.register_host_fn("wasi:cli", "log", Box::new(move |args: &[vybe_bytecode::Value]| {
                    let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
                    let line = parts.join(" ");
                    captured.lock().unwrap().push(line);
                    vybe_bytecode::Value::Null
                }));

                // Run
                match vm.run(chunks) {
                    Ok(_) => {}
                    Err(e) => {
                        let msg = format!("{}", e);
                        if !msg.starts_with("__") { // ignore internal signals like __await__
                            *err.lock().unwrap() = Some(format!("Runtime error: {}", msg));
                        }
                    }
                }

                // Collect any side effect console output
                let effects = queue.borrow_mut().drain();
                for effect in effects {
                    if let vybe_host::SideEffect::ConsoleOutput(msg) = effect {
                        out.lock().unwrap().push(msg);
                    }
                }

                *done.lock().unwrap() = true;
            });
        } else {
            *rs.done.lock().unwrap() = true;
            *rs.error.lock().unwrap() = Some("No project loaded".to_string());
        }
        rs
    });

    // Check if VM finished
    if *run_state.done.lock().unwrap() && *is_running.read() {
        let lines = run_state.output.lock().unwrap().clone();
        let err = run_state.error.lock().unwrap().clone();
        let mut text = lines.join("\n");
        if let Some(e) = err {
            if !text.is_empty() { text.push('\n'); }
            text.push_str(&e);
        }
        if text.is_empty() { text = "Program finished.".to_string(); }
        display_output.set(text);
        is_running.set(false);
    }

    let out_text = display_output.read().clone();
    let running = *is_running.read();

    rsx! {
        div {
            style: "display: flex; flex-direction: column; height: 100%; background: #1e1e1e; color: #d4d4d4; font-family: 'Cascadia Mono', 'Consolas', monospace; font-size: 13px;",
            div {
                style: "padding: 8px 12px; background: #252526; border-bottom: 1px solid #3c3c3c; display: flex; justify-content: space-between; align-items: center;",
                span { "Output" }
                if running {
                    span { style: "color: #569cd6;", "Running..." }
                } else {
                    button {
                        style: "background: #3c3c3c; color: #d4d4d4; border: 1px solid #555; padding: 2px 12px; cursor: pointer;",
                        onclick: move |_| state.run_mode.set(false),
                        "Close"
                    }
                }
            }
            pre {
                style: "flex: 1; margin: 0; padding: 12px; overflow: auto; white-space: pre-wrap;",
                "{out_text}"
            }
        }
    }
}
