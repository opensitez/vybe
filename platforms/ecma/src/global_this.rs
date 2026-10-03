//! ECMA-262 §19.3 — globalThis.
//!
//! `globalThis` is the universal name for the global object. Per spec
//! §19.3.1 it must always resolve to the same object regardless of
//! context (browser → window, Node → global, etc.). In Vybe we expose
//! a fresh plain object via `ecma:globalThis.get` so user code can
//! detect its existence and bind properties on it.

use vybe_runtime::value::Object;
use vybe_runtime::{HostContext, VM, Value};

pub fn register(vm: &mut VM) {
    // Keep identity within this VM, while isolating requests served by
    // separate VMs. PHP also uses this object for its request globals.
    let global_this = Value::Object(vybe_runtime::heap::alloc(Object::new()));
    vm.set_global_owned("globalThis", global_this.clone());
    vm.register_host_fn(
        "ecma:globalThis",
        "get",
        Box::new(move |_ctx: &mut HostContext, _args: &[Value]| global_this.clone()),
    );
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn global_this_is_isolated_between_vms() {
        let mut first = VM::new();
        register(&mut first);
        let mut second = VM::new();
        register(&mut second);

        let Some(Value::Object(first_global)) = first.global("globalThis") else {
            panic!("first VM has no globalThis object");
        };
        let Some(Value::Object(second_global)) = second.global("globalThis") else {
            panic!("second VM has no globalThis object");
        };
        assert!(!std::sync::Arc::ptr_eq(first_global, second_global));
    }
}
