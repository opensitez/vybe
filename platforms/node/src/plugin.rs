//! The node platform as a `vybe_runtime::Plugin` — one plugin, same type as all
//! the others. `init` registers the `node:*` host functions, gating each
//! surface by capability via [`vybe_runtime::Framework::granted`]:
//! - `node:os`/`path`/`process` — always-on (read-only / pure).
//! - `node:fs` — FileRead ∨ FileWrite.
//! - `node:child_process` — Process.
//! - `node:http`/`https` server surface — HttpServer.

/// The node platform plugin.
pub struct Plugin;

impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "node"
    }

    /// Descriptors opened by the program that just ran. The request context is
    /// a thread-local guard cleared on drop, so it needs nothing here.
    // No `reset`: the open-descriptor table is VM-owned storage
    // (`vybe_runtime::resources`), dropped by `reset_to`.

    fn init(&self, fw: &mut vybe_runtime::Framework<'_>) {
        use vybe_runtime::capabilities::Capability;

        let files = fw.granted(Capability::FileRead) || fw.granted(Capability::FileWrite);
        let process = fw.granted(Capability::Process);
        let http_server = fw.granted(Capability::HttpServer);

        let Some(vm) = fw.vm.as_deref_mut() else {
            return;
        };

        crate::register_always_on(vm);
        if files {
            crate::fs::register(vm);
        }
        if process {
            crate::child_process::register(vm);
        }
        if http_server {
            crate::http::register(vm);
            crate::https::register(vm);
        }
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
