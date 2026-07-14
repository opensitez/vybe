use std::sync::{Arc, Mutex};
use vybe_bytecode::value::ObjectKind;
use vybe_bytecode::{HostContext, VM, Value};
use vybe_host::gui_state::GuiState;

/// Run Pascal source through vybex pipeline: pest grammar -> walker -> common AST -> compiler -> VM
pub fn run_pascal(src: &str) -> Vec<String> {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_pascal::register);
    }
    let module = vybe_language_pascal::parse(src).expect("Pascal parse failed");

    let profile = load_pascal_profile();

    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Pascal compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let stdout_buffer: Arc<Mutex<String>> = Arc::new(Mutex::new(String::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    let out = output.clone();
    let stdout = stdout_buffer.clone();
    vm.register_host_fn(
        "wasi:io/streams",
        "[method]output-stream.blocking-write-and-flush",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let text = match args.get(1).unwrap_or(&Value::Null) {
                Value::String(s) => s.to_string(),
                Value::Object(obj) => {
                    let obj = obj.lock().unwrap();
                    if let ObjectKind::Array(items) = &obj.kind {
                        let bytes: Vec<u8> = items
                            .iter()
                            .map(|item| item.as_i32().clamp(0, 255) as u8)
                            .collect();
                        String::from_utf8_lossy(&bytes).to_string()
                    } else {
                        String::new()
                    }
                }
                _ => String::new(),
            };

            let mut pending = stdout.lock().unwrap();
            pending.push_str(&text);
            while let Some(pos) = pending.find('\n') {
                let line: String = pending.drain(..pos).collect();
                pending.drain(..1);
                out.lock().unwrap().push(line);
            }
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("Pascal run failed");
    let residual = stdout_buffer.lock().unwrap().clone();
    if !residual.is_empty() {
        output.lock().unwrap().push(residual);
    }
    let result = output.lock().unwrap().clone();
    result
}

/// Run Pascal source with GUI host functions, return (VM, GuiState, output).
/// Uses register_all_with_gui which creates widgets directly (no side effects).
pub fn run_pascal_gui(src: &str) -> (VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<String>>>) {
    let module = vybe_language_pascal::parse(src).expect("Pascal parse failed");
    let profile = load_pascal_profile();
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Pascal compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vm.register_host_fn(
        "wasi:logging/logging",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("Pascal run failed");
    (vm, gui, output)
}

/// Run Pascal source and capture every `ShowMessage(...)` invocation as a
/// `(text, title)` tuple. Returns `(VM, GuiState, msgbox_log)`.
pub fn run_pascal_gui_capture_msgbox(
    src: &str,
) -> (VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<(String, String)>>>) {
    let module = vybe_language_pascal::parse(src).expect("Pascal parse failed");
    let profile = load_pascal_profile();
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Pascal compile failed");

    let mut vm = VM::new();
    let _output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let gui = vybe_host::register_all_with_gui(&mut vm);

    // Capture msgbox calls. Must be registered AFTER
    // `register_all_with_gui` (which installs the production msgBox)
    // so this override wins.
    let msgboxes: Arc<Mutex<Vec<(String, String)>>> = Arc::new(Mutex::new(Vec::new()));
    let mb_clone = msgboxes.clone();
    vm.register_host_fn(
        "vybe:gui",
        "msgBox",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let text = args.first().map(|v| format!("{}", v)).unwrap_or_default();
            let title = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            mb_clone.lock().unwrap().push((text, title));
            Value::Null
        }),
    );

    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("Pascal run failed");
    (vm, gui, msgboxes)
}

pub fn load_pascal_profile() -> vybe_compiler::profile::LanguageProfile {
    vybe_compiler::profile::parse_profile(vybe_language_pascal::profile_source())
        .expect("Failed to parse Pascal profile")
}
