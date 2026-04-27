use std::sync::Arc;
use vybe_bytecode::{VM, Value, HostContext};

pub fn register(vm: &mut VM) {
    // ECMA-262 coercions retired — moved to `ecma:*`:
    //   toBool         → ecma:boolean.Boolean        (§7.1.2 ToBoolean)
    //   toNumber       → ecma:number.Number          (§7.1.4 ToNumber)
    //   toStr/toString → ecma:string.String          (§7.1.17 ToString)
    //   toFixed        → ecma:number.toFixed         (§21.1.3.3)
    //   parseInt       → ecma:number.parseInt        (§19.2.5)
    //   parseFloat     → ecma:number.parseFloat      (§19.2.4)
    //   isNaN          → ecma:number.isNaN           (§19.2.3 / §21.1.2.4)
    //   isFinite       → ecma:number.isFinite        (§19.2.2 / §21.1.2.2)
    //   isInteger      → ecma:number.isInteger       (§21.1.2.3)
    //   btoa/atob/{en,de}codeURI{,Component} → ecma:string (\u{00a7}19.2.6 + WHATWG \u{00a7}8.3)

    // --- VB-compatible conversion functions ---

    // val(str) → parse as number, 0 on failure
    vm.register_host_fn("vybe:convert", "val", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        Value::F64(s.trim().parse::<f64>().unwrap_or(0.0))
    }));

    // isNumeric(value) → true if can be parsed as number
    vm.register_host_fn("vybe:convert", "isNumeric", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first().unwrap_or(&Value::Null) {
            Value::F64(n) => Value::Bool(!n.is_nan()),
            Value::I32(_) | Value::I64(_) => Value::Bool(true),
            Value::String(s) => Value::Bool(s.trim().parse::<f64>().is_ok()),
            _ => Value::Bool(false),
        }
    }));

    // cint(value) → floor to integer
    vm.register_host_fn("vybe:convert", "cint", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64().floor()).unwrap_or(0.0))
    }));

    // cdbl(value) → to double (identity for numbers, parse for strings)
    vm.register_host_fn("vybe:convert", "cdbl", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first().unwrap_or(&Value::Null) {
            Value::String(s) => Value::F64(s.trim().parse::<f64>().unwrap_or(0.0)),
            v => Value::F64(v.as_f64()),
        }
    }));

    // clng(value) → Long (same as cint for our VM)
    vm.register_host_fn("vybe:convert", "clng", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64().floor()).unwrap_or(0.0))
    }));

    // csng(value) → Single (just convert to float)
    vm.register_host_fn("vybe:convert", "csng", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    // cchar(value) → first character
    vm.register_host_fn("vybe:convert", "cchar", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        match s.chars().next() {
            Some(c) => Value::String(Arc::from(c.to_string().as_str())),
            None => Value::String(Arc::from("")),
        }
    }));

    // hex(value) → hex string
    vm.register_host_fn("vybe:convert", "hex", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let n = args.first().map(|v| v.as_f64() as i64).unwrap_or(0);
        Value::String(Arc::from(format!("{:X}", n).as_str()))
    }));

    // oct(value) → octal string
    vm.register_host_fn("vybe:convert", "oct", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let n = args.first().map(|v| v.as_f64() as i64).unwrap_or(0);
        Value::String(Arc::from(format!("{:o}", n).as_str()))
    }));

    // str(value) → string with leading space for positive numbers (VB6 compat)
    vm.register_host_fn("vybe:convert", "str", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let n = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
        let s = if n >= 0.0 { format!(" {}", n) } else { format!("{}", n) };
        Value::String(Arc::from(s.as_str()))
    }));

    // vartype(value) → VB VarType constant
    vm.register_host_fn("vybe:convert", "varType", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let vt = match args.first().unwrap_or(&Value::Null) {
            Value::Null | Value::Undefined => 1.0,     // vbNull
            Value::Bool(_) => 11.0, // vbBoolean
            Value::I32(_) => 2.0,   // vbInteger
            Value::I64(_) | Value::BigInt(_) => 3.0,   // vbLong
            Value::F64(_) => 5.0,   // vbDouble
            Value::String(_) | Value::Symbol(_) => 8.0, // vbString
            Value::Object(_) | Value::V128(_) | Value::WeakRef(_) => 9.0,// vbObject
        };
        Value::F64(vt)
    }));

    // isdate(value) → always false for now (we don't have a Date type in the VM)
    vm.register_host_fn("vybe:convert", "isDate", Box::new(|_ctx: &mut HostContext, _args: &[Value]| Value::Bool(false)));

    // isempty(value) → null or empty string
    vm.register_host_fn("vybe:convert", "isEmpty", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first().unwrap_or(&Value::Null) {
            Value::Null => Value::Bool(true),
            Value::String(s) => Value::Bool(s.is_empty()),
            _ => Value::Bool(false),
        }
    }));

    // isobject(value) → true if Object
    vm.register_host_fn("vybe:convert", "isObject", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::Bool(matches!(args.first().unwrap_or(&Value::Null), Value::Object(_)))
    }));

    // isTypeOf(value, typeName) → check if value's type matches or inherits from typeName
    // Checks __type property and __class_name for user classes
    vm.register_host_fn("vybe:convert", "isTypeOf", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let target = args.get(1).map(|v| format!("{}", v).to_lowercase()).unwrap_or_default();
        match args.first() {
            Some(Value::Object(obj)) => {
                let o = obj.lock().unwrap();
                // Check __type property
                if let Some(t) = o.properties.get("__type") {
                    if format!("{}", t).to_lowercase() == target { return Value::Bool(true); }
                }
                // Check __control_type
                if let Some(t) = o.properties.get("__control_type") {
                    if format!("{}", t).to_lowercase() == target { return Value::Bool(true); }
                }
                // Check __class_name (for user classes that store class info)
                if let Some(t) = o.properties.get("__class") {
                    if format!("{}", t).to_lowercase() == target { return Value::Bool(true); }
                }
                // For arrays, check "array" or "list"
                if matches!(o.kind, vybe_bytecode::value::ObjectKind::Array(_)) {
                    if target == "array" || target == "list" { return Value::Bool(true); }
                }
                // Everything is an Object
                if target == "object" { return Value::Bool(true); }
                Value::Bool(false)
            }
            Some(Value::String(_)) => Value::Bool(target == "string" || target == "object"),
            Some(Value::F64(_)) | Some(Value::I32(_)) | Some(Value::I64(_)) => {
                Value::Bool(target == "integer" || target == "double" || target == "number" || target == "object")
            }
            Some(Value::Bool(_)) => Value::Bool(target == "boolean" || target == "object"),
            _ => Value::Bool(target == "object"),
        }
    }));

    // iif(condition, trueValue, falseValue) → ternary
    vm.register_host_fn("vybe:convert", "iif", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let cond = match args.first().unwrap_or(&Value::Null) {
            Value::Bool(b) => *b,
            Value::F64(n) => *n != 0.0,
            Value::Null => false,
            _ => true,
        };
        if cond {
            args.get(1).cloned().unwrap_or(Value::Null)
        } else {
            args.get(2).cloned().unwrap_or(Value::Null)
        }
    }));

    // choose(index, val1, val2, ...) → value at index (1-based)
    vm.register_host_fn("vybe:convert", "choose", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let idx = args.first().map(|v| v.as_f64() as usize).unwrap_or(0);
        if idx >= 1 && idx < args.len() {
            args[idx].clone()
        } else {
            Value::Null
        }
    }));

    // rgb(r, g, b) → color as integer
    vm.register_host_fn("vybe:convert", "rgb", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let r = args.first().map(|v| v.as_f64() as u32).unwrap_or(0) & 0xFF;
        let g = args.get(1).map(|v| v.as_f64() as u32).unwrap_or(0) & 0xFF;
        let b = args.get(2).map(|v| v.as_f64() as u32).unwrap_or(0) & 0xFF;
        Value::F64((r | (g << 8) | (b << 16)) as f64)
    }));
}

// base64_encode/base64_decode helpers retired with `vybe:convert.btoa`/`atob`.
// The new home is `crates/vybe_host/src/ecma/string.rs` which has its own
// copies. Same semantics; co-located with the URI fns + btoa/atob there.
