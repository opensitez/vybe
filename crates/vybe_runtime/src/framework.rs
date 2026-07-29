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

use crate::registry::{LanguageHooks, LanguageDef};
use crate::capabilities::{Capabilities, Capability};
use crate::{HostContext, TypeDef, VM, Value};

/// A capability provider. Its [`init`](Plugin::init) registers everything the
/// plugin offers into the [`Framework`].
pub trait Plugin: Sync {
    /// Stable identifier (`"php"`, `"dotnet"`, `"compiler"`, `"host"`…).
    fn name(&self) -> &'static str;

    /// Register this plugin's capabilities. Called once, with the shared
    /// [`Framework`]. Only the registrations a plugin actually needs are made.
    fn init(&self, fw: &mut Framework<'_>);

    /// The sandbox capability this plugin's host functions require, if any.
    ///
    /// During a VM-scoped pass ([`init_all_on_vm`]) a plugin whose required
    /// capability is not granted is skipped entirely — this is how the plugin
    /// framework replaces the old hand-gated `register_with_capabilities`.
    /// `None` (the default) means always-on (e.g. `ecma`/`web` pure runtime).
    /// Plugins that span multiple capabilities (wasi, node) return `None` here
    /// and gate their sub-registrations internally via [`Framework::granted`].
    fn required_capability(&self) -> Option<Capability> {
        None
    }

    /// Second-phase registration, run AFTER every plugin's [`init`] has
    /// completed (see [`init_all_on_vm_with_caps`]). For work that must observe
    /// the fully-populated VM — e.g. the ecma plugin wires `globalThis` +
    /// constructor↔prototype links here, since it resolves host fns other
    /// plugins registered by registry index. Default: no-op.
    fn finalize(&self, _fw: &mut Framework<'_>) {}
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
    /// The sandbox policy for this pass, when VM-scoped. Plugins that span
    /// several capabilities gate their sub-registrations with [`granted`].
    /// `None` during the global pass (no gating) or when a VM pass runs with
    /// full trust.
    caps: Option<&'a Capabilities>,
}

impl<'a> Framework<'a> {
    /// A framework for the global (compile-time) registration pass — no VM.
    pub fn global() -> Framework<'a> {
        Framework {
            vm: None,
            caps: None,
        }
    }

    /// A framework scoped to `vm`, for VM-scoped registration (host functions).
    /// No capability gating (full trust) — see [`with_vm_and_caps`].
    pub fn with_vm(vm: &'a mut VM) -> Framework<'a> {
        Framework {
            vm: Some(vm),
            caps: None,
        }
    }

    /// A VM-scoped framework carrying a sandbox policy. Plugins gate their
    /// sub-registrations via [`granted`]; [`init_all_on_vm`] gates whole
    /// plugins via [`Plugin::required_capability`].
    pub fn with_vm_and_caps(vm: &'a mut VM, caps: &'a Capabilities) -> Framework<'a> {
        Framework {
            vm: Some(vm),
            caps: Some(caps),
        }
    }

    /// Whether `cap` is granted by this pass's sandbox policy. Returns `true`
    /// when there is no policy (global pass or full-trust VM pass), so an
    /// ungated plugin always registers.
    pub fn granted(&self, cap: Capability) -> bool {
        self.caps.is_none_or(|c| c.has(cap))
    }

    /// Register a source language (parser + profile + optional emit/normalize/tree).
    /// Global — works with or without a VM.
    pub fn register_language(&mut self, def: LanguageDef) {
        crate::registry::register_language(def);
    }

    /// Register a language's core hooks (value_eq, relational_compare, proxy_*…).
    /// Global — works with or without a VM.
    pub fn register_hooks(&mut self, name: &'static str, hooks: LanguageHooks) {
        crate::registry::register_hooks(name, hooks);
    }

    /// Register a named runtime type into the VM's `TypeRegistry`, returning
    /// its type id (pass it as a `parent` for subtypes). VM-scoped: a type's
    /// method vtable resolves host fns by registry index, so register types
    /// from a plugin's [`Plugin::finalize`] — after every plugin's host fns
    /// exist. This is the `register_type` counterpart to [`register_host_fn`].
    pub fn register_type(&mut self, def: TypeDef) -> usize {
        match self.vm.as_deref_mut() {
            Some(vm) => vm.type_registry.register(def),
            None => panic!(
                "register_type requires a VM-scoped Framework; run this plugin \
                 via init_all_on_vm, not the global init_all"
            ),
        }
    }

    /// The type id already registered under `name`, if any — for parent links
    /// or for attaching methods to a pre-existing type (e.g. `Object` = id 0).
    pub fn type_id(&self, name: &str) -> Option<usize> {
        self.vm.as_deref().and_then(|vm| vm.type_registry.get_id(name))
    }

    /// The host-fn registry index for `module`::`name`, for building a type's
    /// method vtable. `None` when that host fn is absent (e.g. its capability
    /// was withheld) — the caller simply skips that method.
    pub fn host_fn_index(&self, module: &str, name: &str) -> Option<usize> {
        self.vm.as_deref().and_then(|vm| {
            vm.host_registry
                .get(&(module.to_string(), name.to_string()))
                .copied()
        })
    }

    /// Attach host fn `module`::`name` as method `method` on an existing type
    /// id. No-op when the host fn is absent. The registrar-level form of the
    /// hand-rolled `h(vm, …)` + `add_host_method` pattern.
    pub fn add_host_method(&mut self, type_id: usize, method: &str, module: &str, name: &str) {
        if let Some(idx) = self.host_fn_index(module, name) {
            if let Some(vm) = self.vm.as_deref_mut() {
                vm.type_registry.add_host_method(type_id, method, idx);
            }
        }
    }

    /// Set the constructor of the type named `type_name` to host fn
    /// `module`::`name`. No-op when either the type or the host fn is absent.
    /// The registrar-level form of `get_id` + `set_constructor`.
    pub fn set_constructor(&mut self, type_name: &str, module: &str, name: &str) {
        if let (Some(tid), Some(idx)) = (self.type_id(type_name), self.host_fn_index(module, name)) {
            if let Some(vm) = self.vm.as_deref_mut() {
                vm.type_registry
                    .set_constructor(tid, crate::Method::HostFn(idx));
            }
        }
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

/// A plugin's link-time registration. Every plugin crate submits one of these;
/// the registration loops iterate whatever the final binary linked.
///
/// This is THE registry. There is no plugin list in code — not for languages,
/// not for platforms, not for host-function providers. A binary chooses its
/// plugin set by depending on the crates it wants, and the loop below picks
/// them up. Adding a plugin is a Cargo edit, never a code edit, and no crate
/// has to know another crate's name to run it.
pub struct PluginEntry(pub &'static (dyn Plugin + Sync));
inventory::collect!(PluginEntry);

/// Submit this crate's plugin to the registry. Call once per plugin crate at
/// module scope: `vybe_runtime::register_plugin!(Plugin);`
#[macro_export]
macro_rules! register_plugin {
    ($ty:path) => {
        $crate::inventory::submit! {
            $crate::PluginEntry(&$ty)
        }
    };
}

/// Every plugin linked into this binary, in link order.
pub fn plugins() -> impl Iterator<Item = &'static dyn Plugin> {
    inventory::iter::<PluginEntry>
        .into_iter()
        .map(|e| e.0 as &'static dyn Plugin)
}

/// Run every REGISTERED plugin's `init` against a global (VM-less)
/// [`Framework`] — the compile-time pass that populates language descriptors,
/// hooks and namespace trees.
pub fn init_registered() {
    let mut fw = Framework::global();
    for p in plugins() {
        p.init(&mut fw);
    }
}

/// Both phases for every REGISTERED plugin, scoped to `vm` and gated by
/// `caps`. This is the one registration loop.
pub fn init_all_registered(vm: &mut VM, caps: &Capabilities) {
    init_registered_plugins(vm, caps);
    finalize_registered_plugins(vm, caps);
}

/// Phase 1 for every REGISTERED plugin. Split out so a caller can override
/// host functions between the phases (see [`init_plugins`]).
pub fn init_registered_plugins(vm: &mut VM, caps: &Capabilities) {
    for p in plugins() {
        run_phase(vm, caps, p, false);
    }
}

/// Phase 2 for every REGISTERED plugin.
pub fn finalize_registered_plugins(vm: &mut VM, caps: &Capabilities) {
    for p in plugins() {
        run_phase(vm, caps, p, true);
    }
}

fn run_phase(vm: &mut VM, caps: &Capabilities, p: &dyn Plugin, finalize: bool) {
    if let Some(cap) = p.required_capability() {
        if !caps.has(cap) {
            return;
        }
    }
    let mut fw = Framework::with_vm_and_caps(vm, caps);
    if finalize {
        p.finalize(&mut fw);
    } else {
        p.init(&mut fw);
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

/// Run every plugin's `init` against a full-trust [`Framework`] scoped to
/// `vm` (no capability gating). See [`init_all_on_vm_with_caps`] for the
/// sandboxed path.
///
/// Use for plugins that register VM-scoped capabilities (host functions). A
/// language plugin is harmless here too (it just ignores the VM).
pub fn init_all_on_vm(vm: &mut VM, plugins: &[&dyn Plugin]) {
    let mut fw = Framework::with_vm(vm);
    for p in plugins {
        p.init(&mut fw);
    }
}

/// Run every plugin's `init` against a [`Framework`] scoped to `vm` and
/// gated by `caps` — the capability-based replacement for
/// `vybe_host::register_with_capabilities`.
///
/// A plugin whose [`Plugin::required_capability`] is not granted is skipped
/// wholesale; plugins that span several capabilities (wasi, node) declare
/// `None` and gate their sub-registrations internally via
/// [`Framework::granted`].
pub fn init_all_on_vm_with_caps(vm: &mut VM, caps: &Capabilities, plugins: &[&dyn Plugin]) {
    init_plugins(vm, caps, plugins);
    finalize_plugins(vm, caps, plugins);
}

/// Phase 1 of [`init_all_on_vm_with_caps`]: every plugin's [`Plugin::init`]
/// (host functions, language descriptors), skipping any whose
/// [`Plugin::required_capability`] is not granted.
///
/// Exposed separately so a caller can **interleave** work between the phases —
/// notably overriding host functions before the `finalize` pass wires globals.
/// [`VM::register_host_fn`] appends a *new* index rather than replacing, and
/// `finalize` resolves host fns by index, so an override must land before it.
pub fn init_plugins(vm: &mut VM, caps: &Capabilities, plugins: &[&dyn Plugin]) {
    for p in plugins {
        if let Some(cap) = p.required_capability() {
            if !caps.has(cap) {
                continue;
            }
        }
        let mut fw = Framework::with_vm_and_caps(vm, caps);
        p.init(&mut fw);
    }
}

/// Phase 2 of [`init_all_on_vm_with_caps`]: every plugin's [`Plugin::finalize`],
/// run once all host functions exist (ECMA globals, TypeRegistry vtables).
pub fn finalize_plugins(vm: &mut VM, caps: &Capabilities, plugins: &[&dyn Plugin]) {
    for p in plugins {
        if let Some(cap) = p.required_capability() {
            if !caps.has(cap) {
                continue;
            }
        }
        let mut fw = Framework::with_vm_and_caps(vm, caps);
        p.finalize(&mut fw);
    }
}
