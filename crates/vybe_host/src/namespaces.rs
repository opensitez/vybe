//! Sets up namespace objects as VM globals.
//!
//! Instead of hardcoding VB/JS namespace-to-host mappings in the compiler,
//! the host creates nested objects that any language can traverse via struct_get.
//!
//! Example: `Window.Forms.Button` is a property chain:
//!   global_get "window" → struct_get "forms" → struct_get "button"
//!
//! The compiler just emits struct_get chains for dotted names.
//! The host defines the API shape.

use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{VM, Value};

/// Create a namespace object and insert it at a dotted path in the globals.
/// e.g. set_namespace(vm, "Window.Forms", obj) creates:
///   globals["window"] = { forms: obj }
/// If globals["window"] already exists, it merges.
fn ensure_namespace(vm: &mut VM, path: &[&str]) -> Value {
    if path.is_empty() {
        return Value::Null;
    }

    let root_key = path[0].to_lowercase();

    // Get or create root
    let root = if let Some(existing) = vm.globals.get(&root_key) {
        existing.clone()
    } else {
        let obj = Value::Object(Rc::new(RefCell::new(Object::new())));
        vm.globals.insert(root_key.clone(), obj.clone());
        obj
    };

    // Traverse/create intermediate objects
    let mut current = root;
    for &segment in &path[1..] {
        let key = segment.to_lowercase();
        let next = if let Value::Object(ref obj) = current {
            let o = obj.borrow();
            o.properties.get(&key).cloned()
        } else {
            None
        };

        if let Some(existing) = next {
            current = existing;
        } else {
            let new_obj = Value::Object(Rc::new(RefCell::new(Object::new())));
            if let Value::Object(ref obj) = current {
                obj.borrow_mut().properties.insert(key, new_obj.clone());
            }
            current = new_obj;
        }
    }

    current
}

/// Set a property on a namespace object.
fn set_prop(ns: &Value, name: &str, value: Value) {
    if let Value::Object(obj) = ns {
        obj.borrow_mut().properties.insert(name.to_lowercase(), value);
    }
}

/// Create a host function wrapper that can be called from bytecode.
/// This creates a Value::Object with ObjectKind::Function that,
/// when called, invokes the host function by (module, name).
fn host_fn_ref(vm: &VM, module: &str, name: &str) -> Value {
    // Look up the host function index
    if let Some(&idx) = vm.host_registry.get(&(module.to_string(), name.to_string())) {
        // Create a callable Value that the VM can invoke
        // We store the host function index as a special property
        let mut obj = Object::new();
        obj.properties.insert("__host_module".into(), Value::String(Rc::from(module)));
        obj.properties.insert("__host_name".into(), Value::String(Rc::from(name)));
        obj.properties.insert("__host_idx".into(), Value::F64(idx as f64));
        obj.kind = ObjectKind::HostFunction(idx);
        Value::Object(Rc::new(RefCell::new(obj)))
    } else {
        Value::Null
    }
}

/// Register all namespace objects on the VM.
/// Call this AFTER registering all host functions.
pub fn setup_namespaces(vm: &mut VM) {
    setup_math(vm);
    setup_console(vm);
    setup_string(vm);
    setup_convert(vm);
    setup_array(vm);
    setup_window_forms(vm);
    setup_io(vm);
    setup_vybe_gui(vm);
    setup_application(vm);
}

fn setup_vybe_gui(vm: &mut VM) {
    let gui = ensure_namespace(vm, &["vybe", "gui"]);
    set_prop(&gui, "controlsadd", host_fn_ref(vm, "vybe:gui", "controlsAdd"));
    set_prop(&gui, "setproperty", host_fn_ref(vm, "vybe:gui", "setProperty"));
    set_prop(&gui, "onevent", host_fn_ref(vm, "vybe:gui", "onEvent"));
    set_prop(&gui, "msgbox", host_fn_ref(vm, "vybe:gui", "msgBox"));
    set_prop(&gui, "showform", host_fn_ref(vm, "vybe:gui", "showForm"));
    set_prop(&gui, "closeform", host_fn_ref(vm, "vybe:gui", "closeForm"));
}

fn setup_application(vm: &mut VM) {
    let app = ensure_namespace(vm, &["Application"]);
    set_prop(&app, "run", host_fn_ref(vm, "vybe:gui", "runApplication"));
}

fn setup_math(vm: &mut VM) {
    let math = ensure_namespace(vm, &["Math"]);
    // Methods
    for name in &[
        "floor", "ceil", "round", "abs", "sqrt", "pow", "min", "max",
        "sin", "cos", "tan", "log", "sign", "trunc", "log2", "log10",
        "cbrt", "hypot", "atan2", "asin", "acos", "atan", "exp", "clz32",
    ] {
        set_prop(&math, name, host_fn_ref(vm, "vybe:math", name));
    }
    // Constants
    set_prop(&math, "pi", Value::F64(std::f64::consts::PI));
    set_prop(&math, "e", Value::F64(std::f64::consts::E));
}

fn setup_console(vm: &mut VM) {
    let console = ensure_namespace(vm, &["Console"]);
    set_prop(&console, "writeline", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&console, "write", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&console, "readline", host_fn_ref(vm, "wasi:cli", "readLine"));

    // Also set up JS-style "console" (lowercase)
    let js_console = ensure_namespace(vm, &["console"]);
    set_prop(&js_console, "log", host_fn_ref(vm, "wasi:cli", "log"));
    set_prop(&js_console, "error", host_fn_ref(vm, "wasi:cli", "error"));
    set_prop(&js_console, "warn", host_fn_ref(vm, "wasi:cli", "warn"));
}

fn setup_string(vm: &mut VM) {
    // Microsoft.VisualBasic.Strings functions are top-level in VB,
    // but we also provide them under a "Strings" namespace object
    let strings = ensure_namespace(vm, &["Strings"]);
    for (vb_name, host_name) in &[
        ("left", "left"), ("right", "right"), ("mid", "mid"),
        ("instr", "instr"), ("ucase", "ucase"), ("lcase", "lcase"),
        ("trim", "trim"), ("ltrim", "ltrim"), ("rtrim", "rtrim"),
        ("len", "length"), ("asc", "asc"), ("chr", "chr"),
        ("space", "space"), ("replace", "replaceAll"), ("split", "split"),
    ] {
        set_prop(&strings, vb_name, host_fn_ref(vm, "vybe:string", host_name));
    }
}

fn setup_convert(vm: &mut VM) {
    let convert = ensure_namespace(vm, &["Convert"]);
    set_prop(&convert, "toint32", host_fn_ref(vm, "vybe:convert", "cint"));
    set_prop(&convert, "toint", host_fn_ref(vm, "vybe:convert", "cint"));
    set_prop(&convert, "todouble", host_fn_ref(vm, "vybe:convert", "cdbl"));
    set_prop(&convert, "tostring", host_fn_ref(vm, "vybe:convert", "toString"));
    set_prop(&convert, "toboolean", host_fn_ref(vm, "vybe:convert", "cbool"));
}

fn setup_array(vm: &mut VM) {
    let array = ensure_namespace(vm, &["Array"]);
    set_prop(&array, "isarray", host_fn_ref(vm, "vybe:array", "isArray"));
    set_prop(&array, "from", host_fn_ref(vm, "vybe:array", "from"));
}

fn setup_window_forms(vm: &mut VM) {
    let forms = ensure_namespace(vm, &["Window", "Forms"]);

    // Each control type gets a dedicated host function registered as
    // vybe:gui/new_{Type} that calls newControl with the type name.
    // This makes them callable via struct_get chains:
    //   global_get "window" → struct_get "forms" → struct_get "button" → call 0
    let control_types = [
        "Button", "Label", "TextBox", "CheckBox", "RadioButton",
        "ComboBox", "ListBox", "Panel", "GroupBox", "TabControl",
        "TabPage", "DataGridView", "ProgressBar", "TrackBar",
        "NumericUpDown", "DateTimePicker", "RichTextBox", "PictureBox",
        "MenuStrip", "ToolStrip", "StatusStrip", "SplitContainer",
        "FlowLayoutPanel", "TableLayoutPanel", "LinkLabel", "MaskedTextBox",
        "ListView", "WebBrowser", "MonthCalendar", "ContextMenuStrip",
        "Timer", "BindingSource", "ToolTip", "ImageList",
        "OpenFileDialog", "SaveFileDialog", "FolderBrowserDialog",
        "ColorDialog", "FontDialog",
    ];

    for type_name in &control_types {
        // Register a constructor host function for this control type
        let hn = format!("new_{}", type_name);
        let type_str = type_name.to_string();
        vm.register_host_fn("vybe:gui", &hn, {
            let type_str = type_str.clone();
            Box::new(move |_args: &[Value]| {
                use vybe_bytecode::value::Object;
                use std::sync::atomic::{AtomicU32, Ordering};
                static COUNTER: AtomicU32 = AtomicU32::new(1);
                let id = COUNTER.fetch_add(1, Ordering::Relaxed);
                let name = format!("{}_{}", type_str, id);
                let mut obj = Object::new();
                obj.properties.insert("__control_type".into(), Value::String(Rc::from(type_str.as_str())));
                obj.properties.insert("__control_name".into(), Value::String(Rc::from(name.as_str())));
                obj.properties.insert("name".into(), Value::String(Rc::from(name.as_str())));
                obj.properties.insert("width".into(), Value::F64(100.0));
                obj.properties.insert("height".into(), Value::F64(30.0));
                obj.properties.insert("left".into(), Value::F64(0.0));
                obj.properties.insert("top".into(), Value::F64(0.0));
                Value::Object(Rc::new(RefCell::new(obj)))
            })
        });
        set_prop(&forms, &type_name.to_lowercase(), host_fn_ref(vm, "vybe:gui", &hn));
    }

    // Form constructor
    set_prop(&forms, "form", host_fn_ref(vm, "vybe:gui", "newForm"));
}

fn setup_io(vm: &mut VM) {
    // System.IO.File
    let file = ensure_namespace(vm, &["File"]);
    set_prop(&file, "readalltext", host_fn_ref(vm, "wasi:filesystem", "readFile"));
    set_prop(&file, "writealltext", host_fn_ref(vm, "wasi:filesystem", "writeFile"));
    set_prop(&file, "exists", host_fn_ref(vm, "wasi:filesystem", "exists"));
    set_prop(&file, "delete", host_fn_ref(vm, "wasi:filesystem", "deleteFile"));

    // Also under System.IO path
    let io = ensure_namespace(vm, &["IO"]);
    set_prop(&io, "file", file);

    let directory = ensure_namespace(vm, &["Directory"]);
    set_prop(&directory, "getfiles", host_fn_ref(vm, "wasi:filesystem", "listDir"));
    set_prop(&directory, "createdirectory", host_fn_ref(vm, "wasi:filesystem", "mkdir"));
    set_prop(&directory, "exists", host_fn_ref(vm, "wasi:filesystem", "exists"));
}
