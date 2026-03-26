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
            q.borrow_mut().push(SideEffect::RunApplication { form_name: str_arg(args, 0, "Form1") });
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
}

fn str_arg(args: &[Value], idx: usize, default: &str) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_else(|| default.into())
}
fn i32_arg(args: &[Value], idx: usize, default: i32) -> i32 {
    args.get(idx).map(|v| v.as_f64() as i32).unwrap_or(default)
}
fn vm_to_prop(v: Value) -> PropValue {
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
