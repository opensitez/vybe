//! vybe:rt — The dynamic language runtime module.
//!
//! These are the core operations that dynamic languages need but WASM
//! doesn't have natively. In .wasm output, these become standard imports:
//!
//!   (import "vybe:rt" "dyn_add" (func (param anyref anyref) (result anyref)))
//!
//! This module registers them as host functions so the VM can resolve them.
//! When compiling to real .wasm, the compiler emits `call_import` to these
//! instead of using custom opcodes.

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

pub fn register(vm: &mut VM) {
    // Dynamic arithmetic
    vm.register_host_fn("vybe:rt", "dyn_add", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let a = args.first().unwrap_or(&Value::Null);
        let b = args.get(1).unwrap_or(&Value::Null);
        match (a, b) {
            (Value::String(s1), _) => Value::String(Arc::from(format!("{}{}", s1, b).as_str())),
            (_, Value::String(s2)) => Value::String(Arc::from(format!("{}{}", a, s2).as_str())),
            _ => Value::F64(a.as_f64() + b.as_f64()),
        }
    }));

    vm.register_host_fn("vybe:rt", "dyn_sub", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64()).unwrap_or(0.0) - args.get(1).map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    vm.register_host_fn("vybe:rt", "dyn_mul", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64()).unwrap_or(0.0) * args.get(1).map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    vm.register_host_fn("vybe:rt", "dyn_div", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64()).unwrap_or(0.0) / args.get(1).map(|v| v.as_f64()).unwrap_or(1.0))
    }));

    vm.register_host_fn("vybe:rt", "dyn_mod", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64()).unwrap_or(0.0) % args.get(1).map(|v| v.as_f64()).unwrap_or(1.0))
    }));

    vm.register_host_fn("vybe:rt", "dyn_neg", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(-args.first().map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    // Dynamic comparison
    vm.register_host_fn("vybe:rt", "dyn_eq", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let a = args.first().unwrap_or(&Value::Null);
        let b = args.get(1).unwrap_or(&Value::Null);
        Value::Bool(a.eq(b))
    }));

    vm.register_host_fn("vybe:rt", "dyn_ne", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let a = args.first().unwrap_or(&Value::Null);
        let b = args.get(1).unwrap_or(&Value::Null);
        Value::Bool(!a.eq(b))
    }));

    vm.register_host_fn("vybe:rt", "dyn_lt", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::Bool(args.first().map(|v| v.as_f64()).unwrap_or(0.0) < args.get(1).map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    vm.register_host_fn("vybe:rt", "dyn_gt", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::Bool(args.first().map(|v| v.as_f64()).unwrap_or(0.0) > args.get(1).map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    vm.register_host_fn("vybe:rt", "dyn_le", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::Bool(args.first().map(|v| v.as_f64()).unwrap_or(0.0) <= args.get(1).map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    vm.register_host_fn("vybe:rt", "dyn_ge", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::Bool(args.first().map(|v| v.as_f64()).unwrap_or(0.0) >= args.get(1).map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    // Dynamic type operations
    vm.register_host_fn("vybe:rt", "dyn_not", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::Bool(!truthy(args.first().unwrap_or(&Value::Null)))
    }));

    vm.register_host_fn("vybe:rt", "dyn_to_bool", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::Bool(truthy(args.first().unwrap_or(&Value::Null)))
    }));

    // String operations
    vm.register_host_fn("vybe:rt", "str_concat", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let a = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let b = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        Value::String(Arc::from(format!("{}{}", a, b).as_str()))
    }));

    // Property access (dynamic, string-based)
    vm.register_host_fn("vybe:rt", "get_prop", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let name = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            return obj.lock().unwrap().get(&name);
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:rt", "set_prop", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let name = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let val = args.get(2).cloned().unwrap_or(Value::Null);
            obj.lock().unwrap().set(name, val.clone());
            return val;
        }
        Value::Null
    }));

    // Object/array construction
    vm.register_host_fn("vybe:rt", "new_object", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Object(Arc::new(Mutex::new(Object::new())))
    }));

    vm.register_host_fn("vybe:rt", "new_array", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let elems: Vec<Value> = args.to_vec();
        Value::Object(Arc::new(Mutex::new(Object::new_array(elems))))
    }));

    vm.register_host_fn("vybe:rt", "array_get", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            return obj.lock().unwrap().get(&key);
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:rt", "array_set", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let key = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let val = args.get(2).cloned().unwrap_or(Value::Null);
            obj.lock().unwrap().set(key, val.clone());
            return val;
        }
        Value::Null
    }));

    // Type checking
    vm.register_host_fn("vybe:rt", "typeof", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let tag = match args.first().unwrap_or(&Value::Null) {
            Value::Null => "null",
            Value::Undefined => "undefined",
            Value::Bool(_) => "boolean",
            Value::I32(_) | Value::I64(_) | Value::F64(_) => "number",
            Value::String(_) => "string",
            Value::Object(o) => {
                let ob = o.lock().unwrap();
                match &ob.kind {
                    ObjectKind::Function(_) | ObjectKind::HostFunction(_) => "function",
                    ObjectKind::Array(_) => "array",
                    _ => "object",
                }
            }
            Value::V128(_) => "v128",
            Value::WeakRef(_) => "weakref",
            Value::Symbol(_) => "symbol",
            Value::BigInt(_) => "bigint",
        };
        Value::String(Arc::from(tag))
    }));

    // Global variable access (dynamic, string-keyed)
    // In standard WASM these would be declared globals, but dynamic languages
    // need string-keyed globals which go through the runtime.
    vm.register_host_fn("vybe:rt", "global_get", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        // This is resolved by the VM at runtime through vm.globals
        // The host function is a placeholder — the VM intercepts it
        Value::Null
    }));
    vm.register_host_fn("vybe:rt", "global_set", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        Value::Null
    }));

    // Indexed struct field access (WASM GC style — by field index, not string)
    vm.register_host_fn("vybe:rt", "struct_get_idx", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let o = obj.lock().unwrap();
            // Indexed access: use properties in insertion order (approximate)
            if let Some(val) = o.properties.values().nth(idx) {
                return val.clone();
            }
        }
        Value::Null
    }));

    vm.register_host_fn("vybe:rt", "struct_set_idx", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Object(obj)) = args.first() {
            let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
            let val = args.get(2).cloned().unwrap_or(Value::Null);
            let mut o = obj.lock().unwrap();
            if let Some(key) = o.properties.keys().nth(idx).cloned() {
                o.properties.insert(key, val.clone());
                return val;
            }
        }
        Value::Null
    }));
}

fn truthy(v: &Value) -> bool {
    match v {
        Value::Null | Value::Undefined => false,
        Value::Bool(b) => *b,
        Value::F64(n) => *n != 0.0 && !n.is_nan(),
        Value::I32(n) => *n != 0,
        Value::I64(n) => *n != 0,
        Value::String(s) => !s.is_empty(),
        Value::Object(_) => true,
        Value::WeakRef(w) => w.upgrade().is_some(),
        Value::V128(b) => b.iter().any(|&x| x != 0),
        Value::Symbol(_) => true,
        Value::BigInt(n) => *n != 0,
    }
}
