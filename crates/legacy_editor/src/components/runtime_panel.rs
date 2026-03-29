use dioxus::prelude::*;
use crate::app_state::AppState;
use std::sync::{Arc, Mutex};
use std::process::Child;

#[derive(Clone)]
struct RunState {
    output: Arc<Mutex<Vec<String>>>,
    done: Arc<Mutex<bool>>,
    error: Arc<Mutex<Option<String>>>,
    child: Arc<Mutex<Option<Child>>>,
}

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
            child: Arc::new(Mutex::new(None)),
        };

        let project_path = state.current_project_path.read().clone();
        let project = state.project.read();

        if let Some(proj) = project.as_ref() {
            let has_forms = !proj.forms.is_empty();
            let is_sub_main = proj.code_files.iter()
                .any(|cf| cf.code.to_uppercase().contains("SUB MAIN"));
            let is_form_project = has_forms && !is_sub_main;

            if is_form_project {
                let out = rs.output.clone();
                let done = rs.done.clone();
                let err = rs.error.clone();
                let child_handle = rs.child.clone();

                if let Some(path) = project_path {
                    std::thread::spawn(move || {
                        let vybec = std::env::current_exe()
                            .ok()
                            .and_then(|p| p.parent().map(|d| d.join("vybec")))
                            .unwrap_or_else(|| std::path::PathBuf::from("vybec"));

                        match std::process::Command::new(&vybec)
                            .arg(&path)
                            // No piped stdio — let it use its own console/window
                            .spawn()
                        {
                            Ok(mut child) => {
                                // Bring to foreground on macOS
                                #[cfg(target_os = "macos")]
                                bring_to_front(child.id());

                                // Store handle so Stop button can kill it
                                let pid = child.id();
                                eprintln!("[EDITOR] spawned pid={}", pid);
                                *child_handle.lock().unwrap() = Some(child);

                                // Poll — release lock between iterations so Stop can acquire it
                                loop {
                                    std::thread::sleep(std::time::Duration::from_millis(200));
                                    let finished = {
                                        let mut guard = child_handle.lock().unwrap();
                                        match guard.as_mut() {
                                            Some(c) => match c.try_wait() {
                                                Ok(Some(_)) => { *guard = None; true }
                                                Ok(None) => false,
                                                Err(_) => { *guard = None; true }
                                            },
                                            None => true, // killed by Stop
                                        }
                                    };
                                    if finished { break; }
                                }
                                eprintln!("[EDITOR] pid={} finished", pid);
                            }
                            Err(e) => {
                                *err.lock().unwrap() = Some(format!(
                                    "Could not launch vybec ({}): {}",
                                    vybec.display(), e
                                ));
                            }
                        }
                        *done.lock().unwrap() = true;
                        let _ = out; // keep alive
                    });
                } else {
                    *rs.error.lock().unwrap() = Some("Project must be saved before running.".to_string());
                    *rs.done.lock().unwrap() = true;
                }
            } else {
                // Console project — capture output in-process
                let (all_code, _) = vybe_cli::runner::build_project_code(proj);

                if all_code.trim().is_empty() {
                    *rs.error.lock().unwrap() = Some("No code to run.".to_string());
                    *rs.done.lock().unwrap() = true;
                    return rs;
                }

                let out = rs.output.clone();
                let done = rs.done.clone();
                let err = rs.error.clone();

                std::thread::spawn(move || {
                    let program = match vybe_parser_basic::parse_program(&all_code) {
                        Ok(p) => p,
                        Err(e) => {
                            *err.lock().unwrap() = Some(format!("Parse error: {:?}", e));
                            *done.lock().unwrap() = true;
                            return;
                        }
                    };
                    let chunks = match vybe_compiler_vb::Compiler::new().compile(&program) {
                        Ok(c) => c,
                        Err(e) => {
                            *err.lock().unwrap() = Some(format!("Compile error: {}", e));
                            *done.lock().unwrap() = true;
                            return;
                        }
                    };

                    let mut vm = vybe_bytecode::VM::new();
                    let queue = std::rc::Rc::new(std::cell::RefCell::new(vybe_host::SideEffectQueue::new()));
                    vybe_host::register_all_with_gui(&mut vm, queue.clone());
                    vybe_host::setup_namespaces(&mut vm);

                    let captured = out.clone();
                    vm.register_host_fn("wasi:cli", "log", Box::new(move |args: &[vybe_bytecode::Value]| {
                        let line = args.iter().map(|v| format!("{v}")).collect::<Vec<_>>().join(" ");
                        captured.lock().unwrap().push(line);
                        vybe_bytecode::Value::Null
                    }));

                    match vm.run(chunks) {
                        Ok(_) => {}
                        Err(e) => {
                            let msg = format!("{}", e);
                            if !msg.starts_with("__") {
                                *err.lock().unwrap() = Some(format!("Runtime error: {}", msg));
                            }
                        }
                    }

                    for effect in queue.borrow_mut().drain() {
                        if let vybe_host::SideEffect::ConsoleOutput(msg) = effect {
                            out.lock().unwrap().push(msg);
                        }
                    }

                    *done.lock().unwrap() = true;
                });
            }
        } else {
            *rs.done.lock().unwrap() = true;
            *rs.error.lock().unwrap() = Some("No project loaded".to_string());
        }
        rs
    });

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
    let child_ref = run_state.child.clone();

    rsx! {
        div {
            style: "display: flex; flex-direction: column; height: 100%; background: #1e1e1e; color: #d4d4d4; font-family: 'Cascadia Mono', 'Consolas', monospace; font-size: 13px;",
            div {
                style: "padding: 8px 12px; background: #252526; border-bottom: 1px solid #3c3c3c; display: flex; justify-content: space-between; align-items: center;",
                span { "Output" }
                div {
                    style: "display: flex; gap: 8px; align-items: center;",
                    if running {
                        span { style: "color: #569cd6;", "Running..." }
                        button {
                            style: "background: #8b0000; color: #fff; border: none; padding: 2px 12px; cursor: pointer; border-radius: 2px;",
                            onclick: move |_| {
                                // Kill the child process if it's a form project
                                if let Some(mut child) = child_ref.lock().unwrap().take() {
                                    let _ = child.kill();
                                }
                                state.run_mode.set(false);
                            },
                            "Stop"
                        }
                    } else {
                        button {
                            style: "background: #3c3c3c; color: #d4d4d4; border: 1px solid #555; padding: 2px 12px; cursor: pointer;",
                            onclick: move |_| state.run_mode.set(false),
                            "Close"
                        }
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

/// On macOS, use `open -a` trick or NSRunningApplication to bring the process to front.
#[cfg(target_os = "macos")]
fn bring_to_front(pid: u32) {
    // Give the process a moment to create its window
    std::thread::sleep(std::time::Duration::from_millis(300));
    // Use osascript to activate the process by PID
    let _ = std::process::Command::new("osascript")
        .arg("-e")
        .arg(format!(
            "tell application \"System Events\" to set frontmost of (first process whose unix id is {}) to true",
            pid
        ))
        .spawn();
}
