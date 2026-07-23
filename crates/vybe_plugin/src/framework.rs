//! The unified plugin framework.
//!
//! Every capability provider in Vybe — a source language, a target platform,
//! the compiler itself, the host runtime, an LSP — is a [`Plugin`]. A plugin's
//! [`Plugin::init`] is the single place it wires itself up: it registers
//! whatever it provides (a language, hooks, host functions, …) into the
//! [`Framework`] it is handed. A plugin registers *as many or as few*
//! capabilities as it has.
//!
//! Registration comes in two lifetimes, and the [`Framework`] spans both:
//!
//! * **Global / compile-time** — language descriptors and hooks live in the
//!   process-wide [`crate::registry`] tables. The compiler needs them without
//!   any running VM (to look up a parser/profile for a source file), so these
//!   register even when the framework carries no VM. Drive them with
//!   [`init_all`].
//! * **VM-scoped** — host functions (and, later, a compile hook) attach to a
//!   specific [`VM`]. Drive them with [`init_all_on_vm`], which hands each
//!   plugin a framework carrying that VM.
//!
//! This is static today (`vybex` builds the plugin list and calls [`init_all`])
//! and is the natural dylib entry point tomorrow: a loadable module exports a
//! `Plugin` factory, the host `dlopen`s it, and calls the same `init`.

use crate::registry::{LanguageHooks, LanguagePlugin};
use vybe_bytecode::{HostContext, VM, Value};

/// A capability provider. Its [`init`](Plugin::init) registers everything the
/// plugin offers into the [`Framework`].
pub trait Plugin {
    /// Stable identifier (`"php"`, `"dotnet"`, `"compiler"`, `"host"`…).
    fn name(&self) -> &'static str;

    /// Register this plugin's capabilities. Called once, with the shared
    /// [`Framework`]. Only the registrations a plugin actually needs are made.
    fn init(&self, fw: &mut Framework<'_>);
}

/// The single registration surface handed to every plugin's `init`.
///
/// It bridges the two registration domains: the process-global registries (the
/// language/hook tables in [`crate::registry`]) — always available — and the
/// per-VM registries (host functions) on an *optional* borrowed [`VM`]. A
/// language plugin ignores the VM; a host plugin requires it.
pub struct Framework<'a> {
    /// The VM being provisioned, when this is a VM-scoped registration pass.
    /// `None` during the global (compile-time) pass, where only language
    /// descriptors and hooks are registered.
    pub vm: Option<&'a mut VM>,
}

impl<'a> Framework<'a> {
    /// A framework for the global (compile-time) registration pass — no VM.
    pub fn global() -> Framework<'a> {
        Framework { vm: None }
    }

    /// A framework scoped to `vm`, for VM-scoped registration (host functions).
    pub fn with_vm(vm: &'a mut VM) -> Framework<'a> {
        Framework { vm: Some(vm) }
    }

    /// Register a source language (parser + profile + optional emit/normalize/tree).
    /// Global — works with or without a VM.
    pub fn register_language(&mut self, plugin: LanguagePlugin) {
        crate::registry::register_language(plugin);
    }

    /// Register a language's core hooks (value_eq, relational_compare, proxy_*…).
    /// Global — works with or without a VM.
    pub fn register_hooks(&mut self, name: &'static str, hooks: LanguageHooks) {
        crate::registry::register_hooks(name, hooks);
    }

    /// Register a host function on the VM (`module`::`name`). Requires a
    /// VM-scoped framework (see [`init_all_on_vm`]); panics otherwise, which is
    /// a plugin-wiring bug (a host plugin was run in the global pass).
    pub fn register_host_fn(
        &mut self,
        module: &str,
        name: &str,
        f: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
    ) {
        match self.vm.as_deref_mut() {
            Some(vm) => vm.register_host_fn(module, name, f),
            None => panic!(
                "register_host_fn(\"{module}\", \"{name}\") requires a VM-scoped Framework; \
                 run this plugin via init_all_on_vm, not the global init_all"
            ),
        }
    }
}

/// Run every plugin's `init` against a **global** (VM-less) [`Framework`].
///
/// This is the registration path for languages: each plugin's `init` registers
/// its language descriptor + hooks into the process-wide registry. Call it once
/// at startup. It is the replacement for hand-calling each `register()`.
pub fn init_all(plugins: &[&dyn Plugin]) {
    let mut fw = Framework::global();
    for p in plugins {
        p.init(&mut fw);
    }
}

/// Run every plugin's `init` against a [`Framework`] scoped to `vm`.
///
/// Use for plugins that register VM-scoped capabilities (host functions). A
/// language plugin is harmless here too (it just ignores the VM).
pub fn init_all_on_vm(vm: &mut VM, plugins: &[&dyn Plugin]) {
    let mut fw = Framework::with_vm(vm);
    for p in plugins {
        p.init(&mut fw);
    }
}
