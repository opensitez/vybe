use std::cell::RefCell;
use std::rc::Rc;

use vybe_bytecode::{VM, Value};
use crate::side_effect::{PropValue, SideEffect, SideEffectQueue};

/// Host function indices for GUI operations.
#[derive(Clone)]
pub struct GuiHostFns {
    pub create_form: u16,
    pub add_control: u16,
    pub set_property: u16,
    pub get_property: u16,
    pub on_event: u16,
    pub show_form: u16,
    pub run_application: u16,
    pub msg_box: u16,
    pub close_form: u16,
}

/// Register GUI host functions on the VM. All functions push to the shared queue.
pub fn register_gui_runtime(
    vm: &mut VM,
    queue: Rc<RefCell<SideEffectQueue>>,
) -> GuiHostFns {
    // gui.createForm(title) → returns the form name as a string
    let create_form = vm.register_host_fn("gui.createForm", {
        let q = queue.clone();
        Box::new(move |args: &[Value]| {
            let title = arg_str(args, 0, "Form1");
            let form_name = title.clone();
            // Set form text property
            q.borrow_mut().push(SideEffect::PropertyChange {
                object: form_name.clone(),
                property: "Text".into(),
                value: PropValue::String(title),
            });
            Value::String(Rc::from(form_name.as_str()))
        })
    });

    // gui.addControl(form_name, control_type, control_name, left, top, width, height)
    // → returns control_name
    let add_control = vm.register_host_fn("gui.addControl", {
        let q = queue.clone();
        Box::new(move |args: &[Value]| {
            let form_name = arg_str(args, 0, "Form1");
            let control_type = arg_str(args, 1, "Button");
            let control_name = arg_str(args, 2, "control1");
            let left = arg_i32(args, 3, 0);
            let top = arg_i32(args, 4, 0);
            let width = arg_i32(args, 5, 100);
            let height = arg_i32(args, 6, 30);
            q.borrow_mut().push(SideEffect::AddControl {
                form_name,
                control_name: control_name.clone(),
                control_type,
                left, top, width, height,
                parent_name: String::new(),
            });
            Value::String(Rc::from(control_name.as_str()))
        })
    });

    // gui.setProperty(control_name, property_name, value)
    let set_property = vm.register_host_fn("gui.setProperty", {
        let q = queue.clone();
        Box::new(move |args: &[Value]| {
            let control = arg_str(args, 0, "");
            let property = arg_str(args, 1, "");
            let value = vm_value_to_prop(args.get(2).cloned().unwrap_or(Value::Null));
            q.borrow_mut().push(SideEffect::PropertyChange {
                object: control,
                property,
                value,
            });
            Value::Null
        })
    });

    // gui.getProperty(control_name, property_name) → value
    // Note: this requires reading back from the form state. For now returns Null.
    let get_property = vm.register_host_fn("gui.getProperty", {
        Box::new(move |_args: &[Value]| {
            // TODO: wire to form state for reading back property values
            Value::Null
        })
    });

    // gui.onEvent(control_name, event_name, callback_fn)
    // Stores the callback in the shared queue for the UI to invoke later.
    let on_event = vm.register_host_fn("gui.onEvent", {
        let q = queue.clone();
        Box::new(move |args: &[Value]| {
            let control = arg_str(args, 0, "");
            let event = arg_str(args, 1, "");
            let callback = args.get(2).cloned().unwrap_or(Value::Null);
            q.borrow_mut().register_event(&control, &event, callback);
            Value::Null
        })
    });

    // gui.showForm(form_name)
    let show_form = vm.register_host_fn("gui.showForm", {
        let q = queue.clone();
        Box::new(move |args: &[Value]| {
            let form_name = arg_str(args, 0, "Form1");
            q.borrow_mut().push(SideEffect::FormShow { form_name });
            Value::Null
        })
    });

    // gui.runApplication(form_name) — start the form event loop
    let run_application = vm.register_host_fn("gui.runApplication", {
        let q = queue.clone();
        Box::new(move |args: &[Value]| {
            let form_name = arg_str(args, 0, "Form1");
            q.borrow_mut().push(SideEffect::RunApplication { form_name });
            Value::Null
        })
    });

    // gui.msgBox(text, title)
    let msg_box = vm.register_host_fn("gui.msgBox", {
        let q = queue.clone();
        Box::new(move |args: &[Value]| {
            let text = arg_str(args, 0, "");
            let title = arg_str(args, 1, "");
            q.borrow_mut().push(SideEffect::MsgBox { text, title });
            Value::Null
        })
    });

    // gui.closeForm(form_name)
    let close_form = vm.register_host_fn("gui.closeForm", {
        let q = queue.clone();
        Box::new(move |args: &[Value]| {
            let form_name = arg_str(args, 0, "Form1");
            q.borrow_mut().push(SideEffect::FormClose { form_name });
            Value::Null
        })
    });

    GuiHostFns {
        create_form, add_control, set_property, get_property,
        on_event, show_form, run_application, msg_box, close_form,
    }
}

// -- Helpers to extract args --

fn arg_str(args: &[Value], idx: usize, default: &str) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_else(|| default.into())
}

fn arg_i32(args: &[Value], idx: usize, default: i32) -> i32 {
    args.get(idx).map(|v| v.as_f64() as i32).unwrap_or(default)
}

/// Convert a VM Value to a PropValue for the side effect queue.
fn vm_value_to_prop(v: Value) -> PropValue {
    match v {
        Value::Null => PropValue::Null,
        Value::Bool(b) => PropValue::Bool(b),
        Value::I32(n) => PropValue::Int(n as i64),
        Value::I64(n) => PropValue::Int(n),
        Value::F64(n) => PropValue::Float(n),
        Value::String(s) => PropValue::String(s.to_string()),
        Value::Object(_) => PropValue::String(format!("{}", v)),
    }
}
