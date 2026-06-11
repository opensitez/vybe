//! ECMA-262 §19.3 — globalThis.
//!
//! `globalThis` is the universal name for the global object. Per spec
//! §19.3.1 it must always resolve to the same object regardless of
//! context (browser → window, Node → global, etc.). In Vybe we expose
//! a fresh plain object via `ecma:globalThis.get` so user code can
//! detect its existence and bind properties on it.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::Object;
use vybe_bytecode::{HostContext, VM, Value};

/// Process-global singleton — initialised on first request and
/// returned identically forever after. Matches §19.3.1 "the same
/// global object across realms" expectation.
static GLOBAL_THIS: std::sync::OnceLock<Value> = std::sync::OnceLock::new();

fn global_this() -> Value {
    GLOBAL_THIS
        .get_or_init(|| Value::Object(Arc::new(Mutex::new(Object::new()))))
        .clone()
}

pub fn register(vm: &mut VM) {
    // 0-arg getter — see the existing constants pattern in
    // `ecma:number.MAX_SAFE_INTEGER` etc.
    vm.register_host_fn(
        "ecma:globalThis",
        "get",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| global_this()),
    );
}

/// Returned to namespaces wiring so the existing `vm.globals.insert("globalThis", ...)`
/// can use the same singleton instead of constructing its own.
pub fn shared_singleton() -> Value {
    global_this()
}
