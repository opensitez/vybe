//! The unified plugin framework.
//!
//! Every capability provider in Vybe — a source language, a target platform,
//! the compiler itself, the host runtime, an LSP — is a [`Plugin`]. A plugin's
//! [`Plugin::init`] is the single place it wires itself up: it registers
//! whatever it provides (a language, hooks, host functions, a compile hook…)
//! into the [`Framework`] it is handed. A plugin registers *as many or as few*
//! capabilities as it has.
//!
//! This is static today (`vybex` builds the plugin list and calls
//! [`init_all`]) and is the natural dylib entry point tomorrow: a loadable
//! module exports a `Plugin` factory, the host `dlopen`s it, and calls the
//! same `init`.

use crate::registry::{LanguageHooks, LanguagePlugin};
use vybe_bytecode::{HostContext, Value, VM};

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
/// It bridges the two registration domains: process-global registries (the
/// language/hook tables in [`crate::registry`]) and the per-VM registries
/// (host functions, and — later — the compile hook) on the borrowed [`VM`].
pub struct Framework<'a> {
    /// The VM being provisioned. Host functions and the compile hook attach here.
    pub vm: &'a mut VM,
}

impl<'a> Framework<'a> {
    /// Create a framework over `vm`.
    pub fn new(vm: &'a mut VM) -> Self {
        Self { vm }
    }

    /// Register a source language (parser + profile + optional emit/normalize/tree).
    pub fn register_language(&mut self, plugin: LanguagePlugin) {
        crate::registry::register_language(plugin);
    }

    /// Register a language's core hooks (value_eq, relational_compare, proxy_*…).
    pub fn register_hooks(&mut self, name: &'static str, hooks: LanguageHooks) {
        crate::registry::register_hooks(name, hooks);
    }

    /// Register a host function on the VM (`module`::`name`).
    pub fn register_host_fn(
        &mut self,
        module: &str,
        name: &str,
        f: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
    ) {
        self.vm.register_host_fn(module, name, f);
    }
}

/// Run every plugin's `init` against a fresh [`Framework`] over `vm`.
pub fn init_all(vm: &mut VM, plugins: &[&dyn Plugin]) {
    let mut fw = Framework::new(vm);
    for p in plugins {
        p.init(&mut fw);
    }
}
