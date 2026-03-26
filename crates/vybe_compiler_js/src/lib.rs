pub mod compiler;
pub mod scope;
pub mod module_loader;

pub use compiler::Compiler;
pub use module_loader::load_and_compile;
use vybe_bytecode::{VM, Value};
use std::rc::Rc;

/// Set up VM with all host functions needed by JS, then compile.
pub fn setup_and_compile(
    vm: &mut VM,
    program: &vybe_parser_js::Program,
) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    vybe_host::register_all(vm);
    register_js_coercion(vm);
    Compiler::new().compile(program)
}

/// Set up VM with all host functions + GUI, then compile.
pub fn setup_and_compile_with_gui(
    vm: &mut VM,
    program: &vybe_parser_js::Program,
    queue: std::rc::Rc<std::cell::RefCell<vybe_host::SideEffectQueue>>,
) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    vybe_host::register_all_with_gui(vm, queue);
    register_js_coercion(vm);
    Compiler::new().compile(program)
}

/// Register JS-specific coercion host functions.
pub fn register_js_coercion(vm: &mut VM) {
    vm.register_host_fn("js:coerce", "typeof", Box::new(|args: &[Value]| {
        let v = args.first().unwrap_or(&Value::Null);
        let s = match v {
            Value::Null => "undefined",
            Value::Bool(_) => "boolean",
            Value::F64(_) | Value::I32(_) | Value::I64(_) => "number",
            Value::String(_) => "string",
            Value::Object(o) => {
                let obj = o.borrow();
                match &obj.kind {
                    vybe_bytecode::value::ObjectKind::Function(_) => "function",
                    _ => "object",
                }
            }
        };
        Value::String(Rc::from(s))
    }));

    vm.register_host_fn("js:coerce", "toNumber", Box::new(|args: &[Value]| {
        Value::F64(js_to_number(args.first().unwrap_or(&Value::Null)))
    }));

    vm.register_host_fn("js:coerce", "toString", Box::new(|args: &[Value]| {
        Value::String(Rc::from(format!("{}", args.first().unwrap_or(&Value::Null)).as_str()))
    }));

    vm.register_host_fn("js:coerce", "toBoolean", Box::new(|args: &[Value]| {
        Value::Bool(js_truthy(args.first().unwrap_or(&Value::Null)))
    }));

    vm.register_host_fn("js:coerce", "looseEq", Box::new(|args: &[Value]| {
        let a = args.first().unwrap_or(&Value::Null);
        let b = args.get(1).unwrap_or(&Value::Null);
        Value::Bool(js_loose_eq(a, b))
    }));

    vm.register_host_fn("js:coerce", "add", Box::new(|args: &[Value]| {
        let a = args.first().unwrap_or(&Value::Null);
        let b = args.get(1).unwrap_or(&Value::Null);
        if matches!(a, Value::String(_)) || matches!(b, Value::String(_)) {
            Value::String(Rc::from(format!("{}{}", a, b).as_str()))
        } else {
            Value::F64(a.as_f64() + b.as_f64())
        }
    }));
}

fn js_truthy(v: &Value) -> bool {
    match v {
        Value::Null => false,
        Value::Bool(b) => *b,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::String(s) => !s.is_empty(),
        Value::Object(_) => true,
    }
}

fn js_to_number(v: &Value) -> f64 {
    match v {
        Value::Null => 0.0,
        Value::Bool(true) => 1.0,
        Value::Bool(false) => 0.0,
        Value::F64(n) => *n,
        Value::I32(n) => *n as f64,
        Value::I64(n) => *n as f64,
        Value::String(s) => s.trim().parse().unwrap_or(f64::NAN),
        Value::Object(_) => f64::NAN,
    }
}

fn js_loose_eq(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Null, Value::Null) => true,
        (Value::F64(x), Value::String(s)) | (Value::String(s), Value::F64(x)) => {
            *x == s.trim().parse::<f64>().unwrap_or(f64::NAN)
        }
        (Value::Bool(_), _) => js_loose_eq(&Value::F64(js_to_number(a)), b),
        (_, Value::Bool(_)) => js_loose_eq(a, &Value::F64(js_to_number(b))),
        _ => a.eq(b),
    }
}
