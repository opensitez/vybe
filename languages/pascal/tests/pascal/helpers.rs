use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use vybe_platform_vybe::gui_state::GuiState;
use vybe_runtime::value::ObjectKind;
use vybe_runtime::{HostContext, VM, Value};

/// Run Pascal source through vybex pipeline: pest grammar -> walker -> common AST -> compiler -> VM
pub fn run_pascal(src: &str) -> Vec<String> {
    let needs_isolated_cwd = pascal_source_needs_isolated_cwd(src);
    let _cwd_lock = if needs_isolated_cwd {
        Some(
            pascal_test_cwd_lock()
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner()),
        )
    } else {
        None
    };
    let _cwd = needs_isolated_cwd.then(PascalTestCwd::new);
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_pascal::register);
    }
    let module = vybe_language_pascal::parse(src).expect("Pascal parse failed");

    let profile = load_pascal_profile();

    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Pascal compile failed");

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let stdout_buffer: Arc<Mutex<String>> = Arc::new(Mutex::new(String::new()));
    let out = output.clone();
    vybe_compiler::primitives::platforms::init_platforms(&mut vm);
    vm.register_host_fn(
        "web:console",
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
                _ => String::new() };

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
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("Pascal run failed");
    let residual = stdout_buffer.lock().unwrap().clone();
    if !residual.is_empty() {
        output.lock().unwrap().push(residual);
    }
    let result = output.lock().unwrap().clone();
    result
}

fn pascal_source_needs_isolated_cwd(src: &str) -> bool {
    let lower = src.to_ascii_lowercase();
    [
        "assignfile",
        "rewrite(",
        "reset(",
        "append(",
        "closefile",
        "erase(",
        "rename(",
        "fileexists",
        "textfile",
        "typedfile",
        "file of",
    ]
    .iter()
    .any(|needle| lower.contains(needle))
}

fn pascal_test_cwd_lock() -> &'static Mutex<()> {
    static LOCK: OnceLock<Mutex<()>> = OnceLock::new();
    LOCK.get_or_init(|| Mutex::new(()))
}

struct PascalTestCwd {
    previous: std::path::PathBuf,
    current: std::path::PathBuf }

impl PascalTestCwd {
    fn new() -> Self {
        static NEXT_ID: AtomicUsize = AtomicUsize::new(0);
        let previous = std::env::current_dir().expect("failed to read current test directory");
        let current = std::env::temp_dir().join(format!(
            "vybe-pascal-test-{}-{}",
            std::process::id(),
            NEXT_ID.fetch_add(1, Ordering::Relaxed)
        ));
        std::fs::create_dir_all(&current).expect("failed to create Pascal test temp directory");
        std::env::set_current_dir(&current).expect("failed to enter Pascal test temp directory");
        Self { previous, current }
    }
}

impl Drop for PascalTestCwd {
    fn drop(&mut self) {
        let _ = std::env::set_current_dir(&self.previous);
        let _ = std::fs::remove_dir_all(&self.current);
    }
}

/// Run Pascal source with GUI host functions, return (VM, GuiState, output).
/// Uses register_all_with_gui which creates widgets directly (no side effects).
pub fn run_pascal_gui(src: &str) -> (VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<String>>>) {
    {
        static R: std::sync::Once = std::sync::Once::new();
        R.call_once(vybe_language_pascal::register);
    }
    let module = vybe_language_pascal::parse(src).expect("Pascal parse failed");
    let profile = load_pascal_profile();
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Pascal compile failed");

    // `Plugin::with_gui()` installs the `GuiState` into a process-wide static
    // and `gui_state()` reads it back, so two tests in flight can swap which
    // state each one holds. Pre-existing, and invisible while nothing wrote to
    // that state — it surfaced as a 5-test wobble between identical runs the
    // moment the document projection started filling it in.
    let _gui_turn = gui_test_lock()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());

    let mut vm = VM::new();
    let output: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let out = output.clone();
    // Each test is its own agent and must start on a blank document — the
    // thread-local one persists across tests that share a worker thread, and
    // an inherited control list makes `control_names.len()` depend on which
    // test ran before it.
    vybe_platform_web::html::clear_document_listeners(
        vybe_platform_web::html::active_document(),
    );
    vybe_platform_web::html::reset_active_document();
    let gui = vybe_platform_vybe::init_platforms_with_gui(&mut vm);
    vm.register_host_fn(
        "web:console",
        "log",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let parts: Vec<String> = args.iter().map(|v| format!("{v}")).collect();
            out.lock().unwrap().push(parts.join(" "));
            Value::Null
        }),
    );
    vybe_compiler::primitives::platforms::finalize_platforms(&mut vm);
    vm.run(chunks).expect("Pascal run failed");
    project_document_into(&gui);
    (vm, gui, output)
}

/// Serialises the GUI tests against the process-wide `GuiState` slot and the
/// agent's ambient document. Held for the whole run, not just the install.
fn gui_test_lock() -> &'static Mutex<()> {
    static LOCK: OnceLock<Mutex<()>> = OnceLock::new();
    LOCK.get_or_init(|| Mutex::new(()))
}

/// Copy what the guest built in the DOM onto the `GuiState` the assertions
/// read.
///
/// A control is now `document.createElement(tag)` and its state lives in the
/// `vybe_widgets` document, not in `GuiState` — so without this every
/// `control_names` / `get_property` assertion reads an object nothing writes
/// to any more. The assertions themselves are unchanged on purpose: they say
/// what a form is supposed to contain, and that is exactly as true over the
/// DOM as it was over the old host functions.
///
/// `GuiState`'s own vocabulary is preserved, including its lowercasing — it
/// keyed controls case-insensitively, so `Name := 'txtName'` was always
/// `"txtname"` here.
fn project_document_into(gui: &Arc<Mutex<GuiState>>) {
    let doc = vybe_platform_web::html::active_document();
    let elements = match vybe_platform_web::engine_widgets::with_document(doc, |d| {
        d.elements_with_id()
            .into_iter()
            .map(|(node, id)| {
                (
                    node,
                    id,
                    d.text_content(node),
                    d.value(node),
                    d.checked(node),
                    d.style_property(node, "width"),
                    d.style_property(node, "height"),
                )
            })
            .collect::<Vec<_>>()
    }) {
        Some(elements) => elements,
        None => return };

    // `node → id`, so a listener can be reported against the control name the
    // assertions use. The document body is the form itself.
    let node_names: std::collections::HashMap<u64, String> =
        vybe_platform_web::engine_widgets::with_document(doc, |d| {
            d.elements_with_id()
                .into_iter()
                .map(|(node, id)| (node, id.to_lowercase()))
                .collect()
        })
        .unwrap_or_default();

    let mut g = gui.lock().unwrap();
    for (node, kind, callback) in vybe_platform_web::html::document_listeners(doc) {
        let control = node_names
            .get(&node)
            .cloned()
            .unwrap_or_else(|| g.resolve_control_name("form1"));
        g.register_event(&control, &kind, callback);
    }
    for (node, id, text, value, checked, width, height) in elements {
        let name = id.to_lowercase();
        // The document body IS the form. `GuiState` never counted the form
        // among its `control_names` — it held that identity separately — so
        // projecting it as a control makes a two-control form report three.
        if node == vybe_platform_web::engine::DOCUMENT {
            g.seed_form_identity(&name, &text);
            continue;
        }
        g.track_live_control_name(&name, &name);
        // `Text` is the role every language lowers its caption to, so it is
        // the one property worth mirroring unconditionally; the rest only
        // when the control actually carries them.
        let caption = if text.is_empty() { value.clone() } else { text };
        g.set_property(&name, "Text", &caption);
        if checked {
            g.set_property(&name, "Checked", "True");
        }
        for (prop, css) in [("Width", width), ("Height", height)] {
            if let Some(px) = css.strip_suffix("px") {
                g.set_property(&name, prop, px);
            }
        }
    }
}

/// Run Pascal source and capture every `ShowMessage(...)` invocation as a
/// `(text, title)` tuple. Returns `(VM, GuiState, msgbox_log)`.
pub fn run_pascal_gui_capture_msgbox(
    src: &str,
) -> (VM, Arc<Mutex<GuiState>>, Arc<Mutex<Vec<(String, String)>>>) {
    let module = vybe_language_pascal::parse(src).expect("Pascal parse failed");
    let profile = load_pascal_profile();
    let chunks = vybe_compiler::primitives::Compiler::with_profile(profile)
        .compile(&module)
        .expect("Pascal compile failed");

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
    vm.run(chunks).expect("Pascal run failed");
    (vm, gui, msgboxes)
}

pub fn load_pascal_profile() -> vybe_compiler::profile::LanguageProfile {
    vybe_compiler::profile::parse_profile(vybe_language_pascal::profile_source())
        .expect("Failed to parse Pascal profile")
}
