use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};
use vybe_host::gui_state::GuiState;

/// Run VB source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_vb(src: &str) -> Vec<String> {
    { static R: std::sync::Once = std::sync::Once::new(); R.call_once(vybe_language_vb::register); }
    let module = vybe_language_vb::parse(src).expect("VB parse failed");

    let profile = load_vb_profile();

    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
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
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("VB run failed");
    let result = output.lock().unwrap().clone();
    result
}

/// Run VB source, return (VM, output) for post-run inspection of globals etc.
pub fn run_vb_vm(src: &str) -> (VM, Arc<Mutex<Vec<String>>>) {
    let module = vybe_language_vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
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
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("VB run failed");
    (vm, output)
}

/// Run VB source with GUI host functions, return (VM, GuiState, output).
/// Uses register_all_with_gui which creates widgets directly (no side effects).
pub fn run_vb_gui(src: &str) -> (VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<String>>>) {
    let module = vybe_language_vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");

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
    vm.run(chunks).expect("VB run failed");
    (vm, gui, output)
}

pub fn load_vb_profile() -> vybe_compiler::profile::LanguageProfile {
    vybe_compiler::profile::parse_profile(vybe_language_vb::profile_source())
        .expect("Failed to parse VB profile")
}

pub fn dotnet_expected_one(expected: &str) -> String {
    match expected {
        "true" => "True".to_string(),
        "false" => "False".to_string(),
        other => other.to_string(),
    }
}

pub fn dotnet_expected_lines(expected: &[&str]) -> Vec<String> {
    expected
        .iter()
        .map(|item| dotnet_expected_one(item))
        .collect()
}

/// Run VB source and capture every `MsgBox(...)` invocation as a
/// `(text, title)` tuple. Returns `(VM, GuiState, msgbox_log)`.
///
/// Production msgbox shows a native dialog inline via
/// `vybe_widgets::dialogs::MessageBox::info` — there's no queue to
/// inspect after the fact. To assert on msgbox calls in headless
/// tests, we override the `vybe:gui::msgBox` host fn AFTER all the
/// real registrations are done, swapping the native-dialog impl for
/// one that pushes onto a captured `Vec`. The override is per-test
/// (not global) so production behaviour is unaffected.
pub fn run_vb_gui_capture_msgbox(
    src: &str,
) -> (VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<(String, String)>>>) {
    let module = vybe_language_vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybe_compiler::compiler::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");

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
    vm.run(chunks).expect("VB run failed");
    (vm, gui, msgboxes)
}
