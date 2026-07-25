//! The `wasi:*` host platform as a universal [`Plugin`].
//!
//! Unlike ecma (always-on), wasi spans several sandbox capabilities, so it
//! returns `None` from [`Plugin::required_capability`] and gates each
//! sub-module internally via [`Framework::granted`] — a faithful port of the
//! per-module gating that `vybe_host::register_with_capabilities` did by hand.
//!
//! Only the wasi-crate modules are registered here. `node:fs` (registered
//! alongside filesystem under the same file capability) and `ecma:date` (gated
//! under Clock) belong to the node / ecma plugins respectively.

use vybe_bytecode::capabilities::Capability;
use vybe_plugin::framework::{Framework, Plugin};

/// The `wasi:*` host platform plugin.
pub struct WasiPlugin;

impl Plugin for WasiPlugin {
    fn name(&self) -> &'static str {
        "wasi"
    }

    fn init(&self, fw: &mut Framework<'_>) {
        // Decide every gate up front (immutable borrow of `fw`) before taking
        // the `&mut VM` (mutable borrow) — the two can't overlap.
        let console = fw.granted(Capability::Console);
        let clock = fw.granted(Capability::Clock);
        let random = fw.granted(Capability::Random);
        // Filesystem surface is gated by either file capability, matching
        // `register_with_capabilities`.
        let files = fw.granted(Capability::FileRead) || fw.granted(Capability::FileWrite);
        let environment = fw.granted(Capability::Environment);
        let http = fw.granted(Capability::Http);
        let sockets = fw.granted(Capability::Sockets);
        let database = fw.granted(Capability::Database);
        let crypto = fw.granted(Capability::Crypto);

        let Some(vm) = fw.vm.as_deref_mut() else {
            return; // global (compile-time) pass — no host fns to register.
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
            // io::register runs AFTER filesystem so its [method]input-stream /
            // output-stream handlers take precedence for the standard fd streams.
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

    // required_capability defaults to None — wasi gates per-module above.
}
