//! # `wasm:js-undefined`
//!
//! Stage-1 `js-primitive-builtins` WebAssembly proposal — **not merged**.
//! Provides the "undefined" singleton predicate.
//!
//! Spec: `proposals/js-primitive-builtins/proposals/js-primitive-builtins/Overview.md`

use vybe_bytecode::{HostContext, Value, VM};

pub fn register(vm: &mut VM) {
    // test(externref) -> i32 — returns 1 if the value is undefined
    vm.register_host_fn("wasm:js-undefined", "test", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::I32(if matches!(args.first(), Some(Value::Undefined)) { 1 } else { 0 })
    }));
}
