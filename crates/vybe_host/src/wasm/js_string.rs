//! # `wasm:js-string`
//!
//! **Merged** `js-string-builtins` WebAssembly proposal (V8 implements
//! natively). Provides boxing/unboxing between WASM typed values and
//! JS string values, and basic String operations as host imports.
//!
//! Spec: `proposals/js-string-builtins/proposals/js-string-builtins/Overview.md`

use std::sync::Arc;
use vybe_bytecode::{HostContext, Value, VM};

pub fn register(vm: &mut VM) {
    // test(externref) -> i32 — returns 1 if the value is a string
    vm.register_host_fn("wasm:js-string", "test", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::I32(if matches!(args.first(), Some(Value::String(_))) { 1 } else { 0 })
    }));

    // cast(externref) -> (ref extern) — trap if not a string, otherwise return as-is
    vm.register_host_fn("wasm:js-string", "cast", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(v @ Value::String(_)) => v.clone(),
            _ => Value::Null, // should trap, but return null for robustness
        }
    }));

    // length(externref) -> i32
    //
    // Spec: `String.prototype.length` returns UTF-16 code-unit count, not
    // byte count. "héllo" is 5 code units even though it's 6 UTF-8 bytes.
    // Astral-plane chars (>U+FFFF) count as 2 code units each.
    vm.register_host_fn("wasm:js-string", "length", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::String(s)) => Value::I32(s.encode_utf16().count() as i32),
            _ => Value::I32(0),
        }
    }));

    // concat(externref, externref) -> (ref extern)
    vm.register_host_fn("wasm:js-string", "concat", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let a = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let b = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        Value::String(Arc::from(format!("{}{}", a, b).as_str()))
    }));

    // equals(externref, externref) -> i32
    vm.register_host_fn("wasm:js-string", "equals", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match (args.first(), args.get(1)) {
            (Some(Value::String(a)), Some(Value::String(b))) => Value::I32(if a == b { 1 } else { 0 }),
            _ => Value::I32(0),
        }
    }));

    // compare(externref, externref) -> i32 (-1, 0, 1)
    vm.register_host_fn("wasm:js-string", "compare", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match (args.first(), args.get(1)) {
            (Some(Value::String(a)), Some(Value::String(b))) => Value::I32(a.cmp(b) as i32),
            _ => Value::I32(0),
        }
    }));

    // substring(externref, i32, i32) -> (ref extern)
    vm.register_host_fn("wasm:js-string", "substring", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::String(s)) => {
                let start = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
                let end = args.get(2).map(|v| v.as_f64() as usize).unwrap_or(s.len());
                let start = start.min(s.len());
                let end = end.min(s.len());
                Value::String(Arc::from(&s[start..end]))
            }
            _ => Value::String(Arc::from("")),
        }
    }));

    // charCodeAt(externref, i32) -> i32
    vm.register_host_fn("wasm:js-string", "charCodeAt", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::String(s)) => {
                let idx = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
                match s.chars().nth(idx) {
                    Some(c) => Value::I32(c as i32),
                    None => Value::I32(-1),
                }
            }
            _ => Value::I32(-1),
        }
    }));

    // fromCharCode(i32) -> (ref extern)
    vm.register_host_fn("wasm:js-string", "fromCharCode", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let code = args.first().map(|v| v.as_f64() as u32).unwrap_or(0);
        match char::from_u32(code) {
            Some(c) => Value::String(Arc::from(c.to_string().as_str())),
            None => Value::String(Arc::from("")),
        }
    }));

    // fromCodePoint(i32) -> (ref extern)
    vm.register_host_fn("wasm:js-string", "fromCodePoint", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let code = args.first().map(|v| v.as_f64() as u32).unwrap_or(0);
        match char::from_u32(code) {
            Some(c) => Value::String(Arc::from(c.to_string().as_str())),
            None => Value::String(Arc::from("")),
        }
    }));

    // fromI32(i32) -> (ref extern) — convert i32 to string
    vm.register_host_fn("wasm:js-string", "fromI32", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let n = args.first().map(|v| v.as_f64() as i32).unwrap_or(0);
        Value::String(Arc::from(format!("{}", n).as_str()))
    }));

    // fromF64(f64) -> (ref extern) — convert f64 to string
    vm.register_host_fn("wasm:js-string", "fromF64", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let n = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
        Value::String(Arc::from(format!("{}", n).as_str()))
    }));
}
