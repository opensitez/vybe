//! System.Drawing — Point, Size, Font constructors

use std::cell::RefCell;
use std::rc::Rc;
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::Object;

pub fn register(vm: &mut VM) {
    // New Point(x, y)
    vm.register_host_fn("vybe:drawing", "pointNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let x = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
        let y = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Point")));
        obj.properties.insert("x".into(), Value::F64(x));
        obj.properties.insert("y".into(), Value::F64(y));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // New Size(width, height)
    vm.register_host_fn("vybe:drawing", "sizeNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let w = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
        let h = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Size")));
        obj.properties.insert("width".into(), Value::F64(w));
        obj.properties.insert("height".into(), Value::F64(h));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));

    // New Font(name, size)
    vm.register_host_fn("vybe:drawing", "fontNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let name = args.first().map(|v| format!("{}", v)).unwrap_or_else(|| "Arial".into());
        let size = args.get(1).map(|v| v.as_f64()).unwrap_or(12.0);
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Rc::from("Font")));
        obj.properties.insert("name".into(), Value::String(Rc::from(name.as_str())));
        obj.properties.insert("size".into(), Value::F64(size));
        obj.properties.insert("bold".into(), Value::Bool(false));
        obj.properties.insert("italic".into(), Value::Bool(false));
        Value::Object(Rc::new(RefCell::new(obj)))
    }));
}
