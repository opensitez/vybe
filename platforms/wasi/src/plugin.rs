//! The wasi platform as a `vybe_runtime::Plugin` — one plugin, same type as all
//! the others. `init` registers the `wasi:*` host functions, gating each
//! sub-module by capability via [`vybe_runtime::Framework::granted`] (a faithful
//! port of the old hand-written `register_with_capabilities`).
//!
//! Only wasi-crate modules are registered here. `node:fs` and `ecma:date`
//! belong to the node / ecma plugins.

/// The wasi platform plugin.
pub struct Plugin;

impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "wasi"
    }

    /// Every resource table this platform owns on the running program's
    /// behalf. Runs regardless of which capabilities the CURRENT policy grants:
    /// the state belongs to whichever program created it, under whatever
    /// policy was in force then, and a narrower policy now is no reason to
    /// leave it behind.
    fn reset(&self) {
        crate::reset_host_globals();
    }

    fn init(&self, fw: &mut vybe_runtime::Framework<'_>) {
        use vybe_runtime::capabilities::Capability;

        // Decide every gate up front (immutable borrow) before taking &mut VM.
        let console = fw.granted(Capability::Console);
        let clock = fw.granted(Capability::Clock);
        let random = fw.granted(Capability::Random);
        let files = fw.granted(Capability::FileRead) || fw.granted(Capability::FileWrite);
        let environment = fw.granted(Capability::Environment);
        let http = fw.granted(Capability::Http);
        let sockets = fw.granted(Capability::Sockets);
        let database = fw.granted(Capability::Database);
        let crypto = fw.granted(Capability::Crypto);

        let Some(vm) = fw.vm.as_deref_mut() else {
            return;
        };

        if console {
            crate::console::register(vm);
        }
        if clock {
            crate::clock::register(vm);
        }
        if random {
            crate::random::register(vm);
        }
        if files {
            crate::fs::register(vm); // flat wasi:filesystem convenience surface
            crate::filesystem::register(vm); // canonical wasi:filesystem/types
            // io after filesystem so its stream handlers take precedence.
            crate::io::register(vm);
        }
        if environment {
            crate::env::register(vm);
        }
        if http {
            crate::http::register(vm);
        }
        if sockets {
            crate::sockets::register(vm);
        }
        if database {
            crate::sql::register(vm);
        }
        if crypto {
            crate::crypto::register(vm);
        }
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
