pub mod compiler;
pub mod scope;

pub use compiler::{Compiler, HostFnTable};
use vybe_bytecode::{VM, Value};
use std::rc::Rc;

/// Set up the VM with JS runtime host functions and compile a JS program.
/// Returns the compiled chunks. The VM is ready to run after this.
pub fn setup_js_runtime(vm: &mut VM) -> HostFnTable {
    // Register all JS-specific host functions.
    // The order must match HostFnTable field assignment below.

    let console_log = vm.register_host_fn("console.log", Box::new(|args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        println!("{}", parts.join(" "));
        Value::Null
    }));

    let console_error = vm.register_host_fn("console.error", Box::new(|args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        eprintln!("{}", parts.join(" "));
        Value::Null
    }));

    let console_warn = vm.register_host_fn("console.warn", Box::new(|args: &[Value]| {
        let parts: Vec<String> = args.iter().map(|v| format!("{}", v)).collect();
        eprintln!("[warn] {}", parts.join(" "));
        Value::Null
    }));

    let js_typeof = vm.register_host_fn("js.typeof", Box::new(|args: &[Value]| {
        let v = args.first().unwrap_or(&Value::Null);
        let s = match v {
            Value::Null => "undefined", // JS conflates null/undefined in typeof
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

    let js_to_number = vm.register_host_fn("js.toNumber", Box::new(|args: &[Value]| {
        let v = args.first().unwrap_or(&Value::Null);
        Value::F64(js_coerce_to_number(v))
    }));

    let js_to_string = vm.register_host_fn("js.toString", Box::new(|args: &[Value]| {
        let v = args.first().unwrap_or(&Value::Null);
        Value::String(Rc::from(format!("{}", v).as_str()))
    }));

    let js_to_boolean = vm.register_host_fn("js.toBoolean", Box::new(|args: &[Value]| {
        let v = args.first().unwrap_or(&Value::Null);
        Value::Bool(js_is_truthy(v))
    }));

    let js_loose_eq = vm.register_host_fn("js.looseEq", Box::new(|args: &[Value]| {
        let a = args.first().unwrap_or(&Value::Null);
        let b = args.get(1).unwrap_or(&Value::Null);
        Value::Bool(js_loose_equal(a, b))
    }));

    let js_add = vm.register_host_fn("js.add", Box::new(|args: &[Value]| {
        let a = args.first().unwrap_or(&Value::Null);
        let b = args.get(1).unwrap_or(&Value::Null);
        js_add_impl(a, b)
    }));

    HostFnTable {
        console_log, console_error, console_warn,
        js_typeof, js_to_number, js_to_string, js_to_boolean,
        js_loose_eq, js_add,
    }
}

/// Compile a JS program using a pre-configured host function table.
pub fn compile(program: &vybe_parser_js::Program, host: HostFnTable) -> Result<Vec<vybe_bytecode::Chunk>, String> {
    let compiler = Compiler::new(host);
    compiler.compile(program)
}

// -- JS semantics implemented as plain Rust functions --
// These are the "JS runtime" — called by the VM via host function dispatch.

fn js_is_truthy(v: &Value) -> bool {
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

fn js_coerce_to_number(v: &Value) -> f64 {
    match v {
        Value::Null => 0.0,
        Value::Bool(true) => 1.0,
        Value::Bool(false) => 0.0,
        Value::F64(n) => *n,
        Value::I32(n) => *n as f64,
        Value::I64(n) => *n as f64,
        Value::String(s) => s.trim().parse::<f64>().unwrap_or(f64::NAN),
        Value::Object(_) => f64::NAN,
    }
}

fn js_loose_equal(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Null, Value::Null) => true,
        (Value::F64(x), Value::String(s)) | (Value::String(s), Value::F64(x)) => {
            let n = s.trim().parse::<f64>().unwrap_or(f64::NAN);
            *x == n
        }
        (Value::Bool(_), _) => {
            let an = js_coerce_to_number(a);
            js_loose_equal(&Value::F64(an), b)
        }
        (_, Value::Bool(_)) => {
            let bn = js_coerce_to_number(b);
            js_loose_equal(a, &Value::F64(bn))
        }
        _ => a.eq(b),
    }
}

fn js_add_impl(a: &Value, b: &Value) -> Value {
    // If either is a string, concatenate
    if matches!(a, Value::String(_)) || matches!(b, Value::String(_)) {
        return Value::String(Rc::from(format!("{}{}", a, b).as_str()));
    }
    // Otherwise numeric
    Value::F64(a.as_f64() + b.as_f64())
}
