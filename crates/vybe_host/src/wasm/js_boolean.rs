//! # `wasm:js-boolean`
//!
//! Stage-1 `js-primitive-builtins` WebAssembly proposal — **not merged**.
//! Box/unbox between WASM typed values and JS Boolean values.
//!
//! Spec: `proposals/js-primitive-builtins/proposals/js-primitive-builtins/Overview.md`

use vybe_bytecode::{HostContext, Value, VM};

pub fn register(vm: &mut VM) {
    // test(externref) -> i32 — returns 1 if the value is a boolean
    vm.register_host_fn("wasm:js-boolean", "test", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::I32(if matches!(args.first(), Some(Value::Bool(_))) { 1 } else { 0 })
    }));

    // cast(externref) -> i32 — extract boolean as i32 (0 or 1), trap if not bool
    vm.register_host_fn("wasm:js-boolean", "cast", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::Bool(b)) => Value::I32(if *b { 1 } else { 0 }),
            _ => Value::I32(0),
        }
    }));
}
