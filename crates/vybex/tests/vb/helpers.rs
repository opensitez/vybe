use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_host::gui_state::GuiState;

/// Run VB source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_vb(src: &str) -> Vec<String> {
    let module = vybex::languages::vb::parse(src).expect("VB parse failed");

    let profile = load_vb_profile();

    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("VB compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("VB run failed");
    let result = output.lock().unwrap().clone();
    result
}

/// Run VB source, return (VM, output) for post-run inspection of globals etc.
pub fn run_vb_vm(src: &str) -> (VM, Arc<Mutex<Vec<String>>>) {
    let module = vybex::languages::vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("VB compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vybe_host::register_all(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("VB run failed");
    (vm, output)
}

/// Run VB source with GUI host functions, return (VM, GuiState, output).
/// Uses register_all_with_gui which creates widgets directly (no side effects).
pub fn run_vb_gui(src: &str) -> (VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<String>>>) {
    let module = vybex::languages::vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybex::compiler::Compiler::with_profile(profile)
        .compile(&module).expect("VB compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    let gui = vybe_host::register_all_with_gui(&mut vm);
    vm.register_host_fn("wasi:cli", "log", Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
        out.lock().unwrap().push(parts.join(" "));
        Value::Null
    }));
    vybe_host::setup_namespaces(&mut vm);
    vm.run(chunks).expect("VB run failed");
    (vm, gui, output)
}

pub fn load_vb_profile() -> vybex::profile::LanguageProfile {
    vybex::profile::parse_profile(vybex::languages::vb::profile_source())
        .expect("Failed to parse VB profile")
}
