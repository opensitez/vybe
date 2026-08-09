use std::sync::{Arc, Mutex, MutexGuard, OnceLock};
use vybe_platform_vybe::gui_state::GuiState;
use vybe_runtime::{HostContext, VM, Value};

fn vb_runtime_test_guard() -> MutexGuard<'static, ()> {
    static LOCK: OnceLock<Mutex<()>> = OnceLock::new();
    LOCK.get_or_init(|| Mutex::new(()))
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
}

/// Capture VB console output on BOTH surfaces, mirroring the C#/libc harness:
/// - `wasi:io/streams.blocking-write-and-flush` — the byte-faithful
///   `Console.Write`/`WriteLine` path (text is arg[1]; `WriteLine`'s newline is
///   part of that text, `Write` has none).
/// - `wasi:logging/logging.log` — line-oriented; still used by any residual log
///   call, newline implied.
///
/// Fragments are accumulated raw; [`finalize_lines`] splits them into lines.
fn register_output_capture(vm: &mut VM) -> Arc<Mutex<Vec<String>>> {
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    vm.register_host_fn(
        "web:console",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            let mut joined = parts.join(" ");
            joined.push('\n');
            out.lock().unwrap().push(joined);
            Value::Null
        }),
    );
    let out = output.clone();
    vm.register_host_fn(
        "wasi:io/streams",
        "[method]output-stream.blocking-write-and-flush",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            if let Some(text) = args.get(1) {
                let s = format!("{text}");
                if !s.is_empty() {
                    out.lock().unwrap().push(s);
                }
            }
            Value::Null
        }),
    );
    output
}

/// Concatenate captured fragments and split into lines — one line per captured
/// entry. Strips only the final empty artifact of a trailing newline; interior
/// empties are real content (`Console.WriteLine("")`).
fn finalize_lines(output: &Arc<Mutex<Vec<String>>>) -> Vec<String> {
    let joined: String = output.lock().unwrap().concat();
    if joined.is_empty() {
        return Vec::new();
    }
    let mut lines: Vec<String> = joined
        .split('\n')
        .map(|l| l.trim_end_matches('\r').to_string())
        .collect();
    if joined.ends_with('\n') {
        lines.pop();
    }
    lines
}

/// Run VB source through vybex pipeline: pest grammar → walker → common AST → compiler → VM
pub fn run_vb(src: &str) -> Vec<String> {
    let _guard = vb_runtime_test_guard();
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_vb::register);
    }
    let module = vybe_language_vb::parse(src).expect("VB parse failed");

    let profile = load_vb_profile();

    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");

    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    let output = register_output_capture(&mut vm);
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("VB run failed");
    finalize_lines(&output)
}

/// Run VB source, return (VM, output) for post-run inspection of globals etc.
pub fn run_vb_vm(src: &str) -> (VM, Arc<Mutex<Vec<String>>>) {
    let _guard = vb_runtime_test_guard();
    let module = vybe_language_vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");

    let mut vm = VM::new();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    let output = register_output_capture(&mut vm);
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("VB run failed");
    // Split accumulated fragments into lines in place so callers inspecting the
    // shared buffer see one entry per printed line, as before.
    let lines = finalize_lines(&output);
    *output.lock().unwrap() = lines;
    (vm, output)
}

/// Run VB source with GUI host functions, return (VM, GuiState, output).
/// Uses register_all_with_gui which creates widgets directly (no side effects).
pub fn run_vb_gui(src: &str) -> (VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<String>>>) {
    let _guard = vb_runtime_test_guard();
    let module = vybe_language_vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");

    let mut vm = VM::new();
    let gui = vybe_platform_vybe::init_platforms_with_gui(&mut vm);
    let output = register_output_capture(&mut vm);
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("VB run failed");
    let lines = finalize_lines(&output);
    *output.lock().unwrap() = lines;
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
    let _guard = vb_runtime_test_guard();
    let module = vybe_language_vb::parse(src).expect("VB parse failed");
    let profile = load_vb_profile();
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("VB compile failed");

    let mut vm = VM::new();
    let _output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let gui = vybe_platform_vybe::init_platforms_with_gui(&mut vm);

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

    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("VB run failed");
    (vm, gui, msgboxes)
}
