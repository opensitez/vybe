//! The `ecma:*` host platform as a universal [`Plugin`].
//!
//! ecma is a pure-computation runtime (ECMA-262 host functions), so it is
//! always-on — no sandbox capability gates it. This is the first platform to
//! join the same plugin framework the language frontends already use; the
//! host-fn registration that `vybe_host::register_all` did by hand becomes a
//! `Plugin::init` driven by `vybe_plugin::init_all_on_vm_with_caps`.
//!
//! Note: [`crate::register_globals`] (constructor↔prototype wiring) is NOT run
//! here — it must run once, after EVERY plugin's host fns are registered, so it
//! stays a separate post-registration pass owned by the assembler.

use vybe_plugin::framework::{Framework, Plugin};

/// The `ecma:*` runtime platform plugin.
pub struct EcmaPlugin;

impl Plugin for EcmaPlugin {
    fn name(&self) -> &'static str {
        "ecma"
    }

    fn init(&self, fw: &mut Framework<'_>) {
        // VM-scoped: register the ECMA-262 host functions. In the global
        // (compile-time) pass there is no VM and nothing to do.
        if let Some(vm) = fw.vm.as_deref_mut() {
            crate::register(vm);
        }
    }

    // required_capability defaults to None — always-on pure runtime.
}
