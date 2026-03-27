use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value};
use crate::side_effect::{PropValue, SideEffect, SideEffectQueue};

pub fn register(vm: &mut VM, queue: Rc<RefCell<SideEffectQueue>>) {
    let q = queue.clone();
    vm.register_host_fn("vybe:gui", "createForm", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            let title = str_arg(args, 0, "Form1");
            let name = title.clone();
            q.borrow_mut().push(SideEffect::PropertyChange {
                object: name.clone(), property: "Text".into(),
                value: PropValue::String(title),
            });
            Value::String(Rc::from(name.as_str()))
        })
    });
    vm.register_host_fn("vybe:gui", "addControl", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            let form_name = str_arg(args, 0, "Form1");
            let control_type = str_arg(args, 1, "Button");
            let control_name = str_arg(args, 2, "control1");
            let left = i32_arg(args, 3, 0);
            let top = i32_arg(args, 4, 0);
            let width = i32_arg(args, 5, 100);
            let height = i32_arg(args, 6, 30);
            q.borrow_mut().push(SideEffect::AddControl {
                form_name, control_name: control_name.clone(), control_type,
                left, top, width, height, parent_name: String::new(),
            });
            Value::String(Rc::from(control_name.as_str()))
        })
    });
    vm.register_host_fn("vybe:gui", "setProperty", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            let control = str_arg(args, 0, "");
            let property = str_arg(args, 1, "");
            let value = vm_to_prop(args.get(2).cloned().unwrap_or(Value::Null));
            q.borrow_mut().push(SideEffect::PropertyChange { object: control, property, value });
            Value::Null
        })
    });
    vm.register_host_fn("vybe:gui", "getProperty", Box::new(|_| Value::Null));
    vm.register_host_fn("vybe:gui", "onEvent", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            let control = str_arg(args, 0, "");
            let event = str_arg(args, 1, "");
            let callback = args.get(2).cloned().unwrap_or(Value::Null);
            eprintln!("[onEvent] ctrl={} event={} callback={}", control, event, callback.type_tag());
            q.borrow_mut().register_event(&control, &event, callback);
            Value::Null
        })
    });
    vm.register_host_fn("vybe:gui", "showForm", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            q.borrow_mut().push(SideEffect::FormShow { form_name: str_arg(args, 0, "Form1") });
            Value::Null
        })
    });
    vm.register_host_fn("vybe:gui", "runApplication", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            // Accept either a string name or a form object
            let form_name = match args.first() {
                Some(Value::Object(obj)) => {
                    let o = obj.borrow();
                    o.properties.get("__control_name")
                        .map(|v| format!("{}", v))
                        .unwrap_or_else(|| str_arg(args, 0, "Form1"))
                }
                _ => str_arg(args, 0, "Form1"),
            };
            q.borrow_mut().push(SideEffect::RunApplication { form_name });
            Value::Null
        })
    });
    vm.register_host_fn("vybe:gui", "msgBox", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            q.borrow_mut().push(SideEffect::MsgBox { text: str_arg(args, 0, ""), title: str_arg(args, 1, "") });
            Value::Null
        })
    });
    vm.register_host_fn("vybe:gui", "closeForm", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            q.borrow_mut().push(SideEffect::FormClose { form_name: str_arg(args, 0, "Form1") });
            Value::Null
        })
    });

    // --- WinForms-style OOP API ---
    // newControl(controlType) → creates an object representing a control
    // The object has __control_type, __control_name, and methods
    vm.register_host_fn("vybe:gui", "newControl", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            use vybe_bytecode::value::Object;
            static COUNTER: std::sync::atomic::AtomicU32 = std::sync::atomic::AtomicU32::new(1);
            let control_type = str_arg(args, 0, "Button");
            let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            let name = format!("{}_{}", control_type, id);
            let mut obj = Object::new();
            obj.properties.insert("__control_type".into(), Value::String(Rc::from(control_type.as_str())));
            obj.properties.insert("__control_name".into(), Value::String(Rc::from(name.as_str())));
            obj.properties.insert("name".into(), Value::String(Rc::from(name.as_str())));
            // Default size
            obj.properties.insert("width".into(), Value::F64(100.0));
            obj.properties.insert("height".into(), Value::F64(30.0));
            obj.properties.insert("left".into(), Value::F64(0.0));
            obj.properties.insert("top".into(), Value::F64(0.0));
            Value::Object(Rc::new(std::cell::RefCell::new(obj)))
        })
    });

    // controlSetProperty(controlObj, property, value)
    // Pushes a PropertyChange side effect using the control's __control_name
    vm.register_host_fn("vybe:gui", "controlSetProperty", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            if let Some(Value::Object(obj)) = args.first() {
                let o = obj.borrow();
                let control_name = o.properties.get("__control_name")
                    .map(|v| format!("{}", v))
                    .unwrap_or_default();
                let property = str_arg(args, 1, "");
                let value = vm_to_prop(args.get(2).cloned().unwrap_or(Value::Null));
                // Also store on the object itself for later reads
                drop(o);
                obj.borrow_mut().properties.insert(property.to_lowercase(), args.get(2).cloned().unwrap_or(Value::Null));
                q.borrow_mut().push(SideEffect::PropertyChange {
                    object: control_name, property, value,
                });
            }
            Value::Null
        })
    });

    // controlsAdd(formName, controlObj) — adds a control to a form
    // Reads __control_type, __control_name, and position from the object
    vm.register_host_fn("vybe:gui", "controlsAdd", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            // Accept either a string or a form object as first arg
            let form_name = match args.first() {
                Some(Value::Object(obj)) => {
                    let o = obj.borrow();
                    o.properties.get("__control_name")
                        .or_else(|| o.properties.get("name"))
                        .map(|v| format!("{}", v))
                        .unwrap_or_else(|| "Form1".into())
                }
                _ => str_arg(args, 0, "Form1"),
            };
            if let Some(Value::Object(obj)) = args.get(1) {
                let o = obj.borrow();
                let control_type = o.properties.get("__control_type")
                    .map(|v| format!("{}", v)).unwrap_or_else(|| "Button".into());
                let control_name = o.properties.get("name")
                    .or_else(|| o.properties.get("__control_name"))
                    .map(|v| format!("{}", v)).unwrap_or_else(|| "ctrl".into());
                // Read position from Location (Point with x,y) or direct left/top
                let (left, top) = if let Some(Value::Object(loc)) = o.properties.get("location") {
                    let loc = loc.borrow();
                    (loc.properties.get("x").map(|v| v.as_f64() as i32).unwrap_or(0),
                     loc.properties.get("y").map(|v| v.as_f64() as i32).unwrap_or(0))
                } else {
                    (o.properties.get("left").map(|v| v.as_f64() as i32).unwrap_or(0),
                     o.properties.get("top").map(|v| v.as_f64() as i32).unwrap_or(0))
                };
                // Read size from Size (with width,height) or direct width/height
                let (width, height) = if let Some(Value::Object(sz)) = o.properties.get("size") {
                    let sz = sz.borrow();
                    (sz.properties.get("width").map(|v| v.as_f64() as i32).unwrap_or(100),
                     sz.properties.get("height").map(|v| v.as_f64() as i32).unwrap_or(30))
                } else {
                    (o.properties.get("width").map(|v| v.as_f64() as i32).unwrap_or(100),
                     o.properties.get("height").map(|v| v.as_f64() as i32).unwrap_or(30))
                };
                q.borrow_mut().push(SideEffect::AddControl {
                    form_name, control_name: control_name.clone(), control_type,
                    left, top, width, height, parent_name: String::new(),
                });
                // Emit any properties that were set before adding
                // Capitalize first letter to match WinForms convention (text → Text)
                for (key, val) in &o.properties {
                    if key.starts_with("__") || key == "name" || key == "left" || key == "top"
                        || key == "width" || key == "height" { continue; }
                    let prop_name = capitalize_first(key);
                    q.borrow_mut().push(SideEffect::PropertyChange {
                        object: control_name.clone(),
                        property: prop_name,
                        value: vm_to_prop(val.clone()),
                    });
                }
            }
            Value::Null
        })
    });

    // newForm(title?) → creates a form object
    vm.register_host_fn("vybe:gui", "newForm", {
        let q = q.clone();
        Box::new(move |args: &[Value]| {
            use vybe_bytecode::value::Object;
            let title = str_arg(args, 0, "Form1");
            let name = title.clone();
            let mut obj = Object::new();
            obj.properties.insert("__control_type".into(), Value::String(Rc::from("Form")));
            obj.properties.insert("__control_name".into(), Value::String(Rc::from(name.as_str())));
            obj.properties.insert("name".into(), Value::String(Rc::from(name.as_str())));
            obj.properties.insert("text".into(), Value::String(Rc::from(title.as_str())));
            obj.properties.insert("width".into(), Value::F64(800.0));
            obj.properties.insert("height".into(), Value::F64(600.0));
            q.borrow_mut().push(SideEffect::PropertyChange {
                object: name.clone(), property: "Text".into(),
                value: PropValue::String(title),
            });
            Value::Object(Rc::new(std::cell::RefCell::new(obj)))
        })
    });

    // No-op function for layout methods (SuspendLayout, ResumeLayout, etc.)
    vm.register_host_fn("vybe:gui", "noop", Box::new(|_args: &[Value]| Value::Null));

    // Generic WinForms control constructor: new_Button(), new_Label(), etc.
    // Creates a simple object with __control_type set.
    let control_types = [
        "Button", "Label", "TextBox", "CheckBox", "RadioButton", "ComboBox",
        "ListBox", "Panel", "GroupBox", "TabControl", "TabPage", "DataGridView",
        "ProgressBar", "TrackBar", "NumericUpDown", "DateTimePicker", "RichTextBox",
        "PictureBox", "MenuStrip", "ToolStrip", "StatusStrip", "SplitContainer",
        "FlowLayoutPanel", "TableLayoutPanel", "LinkLabel", "MaskedTextBox",
        "ListView", "WebBrowser", "MonthCalendar", "ContextMenuStrip",
        "Timer", "BindingSource", "DataSet", "ImageList", "ToolTip",
        "NotifyIcon", "ErrorProvider", "HelpProvider", "BackgroundWorker",
        "Form", "TreeView",
    ];
    for ct in control_types {
        let type_name = ct.to_string();
        vm.register_host_fn("vybe:gui", &format!("new_{}", ct), Box::new(move |_args: &[Value]| {
            let mut obj = vybe_bytecode::value::Object::new();
            obj.properties.insert("__control_type".into(), Value::String(Rc::from(type_name.as_str())));
            obj.properties.insert("__type".into(), Value::String(Rc::from(type_name.as_str())));
            Value::Object(Rc::new(RefCell::new(obj)))
        }));
    }

    // addHandler(controlName, eventName, callback) → registers an event handler
    vm.register_host_fn("vybe:gui", "addHandler", {
        let q = queue.clone();
        Box::new(move |args: &[Value]| {
            let ctrl = str_arg(args, 0, "");
            let event = str_arg(args, 1, "");
            if let Some(callback) = args.get(2) {
                q.borrow_mut().register_event(&ctrl, &event, callback.clone());
            }
            Value::Null
        })
    });

    // removeHandler(eventTarget, handlerName) → removes an event handler
    vm.register_host_fn("vybe:gui", "removeHandler", Box::new(|args: &[Value]| {
        let _target = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let _handler = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        // In the side-effect model, event handlers are managed by the UI layer.
        // Removal would be a side-effect, but for now we accept the call silently.
        Value::Null
    }));
}

fn str_arg(args: &[Value], idx: usize, default: &str) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_else(|| default.into())
}
fn i32_arg(args: &[Value], idx: usize, default: i32) -> i32 {
    args.get(idx).map(|v| v.as_f64() as i32).unwrap_or(default)
}
fn vm_to_prop(v: Value) -> PropValue {
    match v {
        Value::Null | Value::Undefined => PropValue::Null,
        Value::Bool(b) => PropValue::Bool(b),
        Value::I32(n) => PropValue::Int(n as i64),
        Value::I64(n) => PropValue::Int(n),
        Value::F64(n) => PropValue::Float(n),
        Value::String(s) => PropValue::String(s.to_string()),
        Value::Object(_) | Value::V128(_) => PropValue::String(format!("{}", v)),
    }
}

fn capitalize_first(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        None => String::new(),
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
    }
}
