//! # `wasm:js-number`
//!
//! Stage-1 `js-primitive-builtins` WebAssembly proposal — **not merged**,
//! no V8 native implementation. Provides boxing/unboxing between WASM
//! typed values and JS Number values.
//!
//! Spec: `proposals/js-primitive-builtins/proposals/js-primitive-builtins/Overview.md`

use vybe_bytecode::{HostContext, Value, VM};

pub fn register(vm: &mut VM) {
    // test(externref) -> i32 — returns 1 if the value is a number
    vm.register_host_fn("wasm:js-number", "test", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::I32(match args.first() {
            Some(Value::F64(_)) | Some(Value::I32(_)) | Some(Value::I64(_)) => 1,
            _ => 0,
        })
    }));

    // testI32(externref) -> i32 — returns 1 if the value fits in i32
    vm.register_host_fn("wasm:js-number", "testI32", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::I32(match args.first() {
            Some(Value::I32(_)) => 1,
            Some(Value::F64(n)) => if *n == (*n as i32) as f64 { 1 } else { 0 },
            _ => 0,
        })
    }));

    // fromF64(f64) -> externref — box an f64 into the universal representation
    vm.register_host_fn("wasm:js-number", "fromF64", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    // fromI32(i32) -> externref — box an i32
    vm.register_host_fn("wasm:js-number", "fromI32", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::I32(args.first().map(|v| v.as_f64() as i32).unwrap_or(0))
    }));

    // toF64(externref) -> f64 — unbox to f64
    vm.register_host_fn("wasm:js-number", "toF64", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::F64(args.first().map(|v| v.as_f64()).unwrap_or(0.0))
    }));

    // toI32(externref) -> i32 — unbox to i32 (truncate)
    vm.register_host_fn("wasm:js-number", "toI32", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        Value::I32(args.first().map(|v| v.as_f64() as i32).unwrap_or(0))
    }));
}
