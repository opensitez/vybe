//! Vybe System Interface (VSI) modules.
//!
//! Each module registers host functions with (module, name) pairs on the VM,
//! following the WASI capability-based security model.
//!
//! Capabilities control which modules are available:
//! - Safe (always on): math, string, array, convert, json, object, types, rt
//! - Requires permission: fs, database, sockets, http, env, gui, threading, crypto

pub mod console;
pub mod math;
pub mod string;
pub mod array;
pub mod convert;
pub mod json;
pub mod fs;
pub mod clock;
pub mod env;
pub mod random;
pub mod http;
pub mod object;
pub mod regex;
pub mod collections;
pub mod runtime;
pub mod database;
pub mod gui;
pub mod types;
pub mod sockets;
pub mod crypto;
pub mod xml;
pub mod threading;
pub mod data;
pub mod drawing;
pub mod rt;

use vybe_bytecode::{VM, Value, HostContext};
use std::collections::HashSet;

/// Capability flags for host module access.
/// Follows WASI's capability-based security model.
#[derive(Debug, Clone)]
pub struct Capabilities {
    granted: HashSet<Capability>,
}

/// Individual capabilities that can be granted or denied.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Capability {
    /// Console I/O (stdout, stderr). Safe for most contexts.
    Console,
    /// Filesystem read access.
    FileRead,
    /// Filesystem write access (implies FileRead).
    FileWrite,
    /// Network: outbound HTTP requests.
    Http,
    /// Network: TCP/UDP sockets (server + client).
    Sockets,
    /// Database connections (SQLite, MySQL, etc.).
    Database,
    /// Environment variables and process args.
    Environment,
    /// GUI / window creation.
    Gui,
    /// Threading / background tasks.
    Threading,
    /// Cryptographic operations.
    Crypto,
    /// System clock access (time, sleep).
    Clock,
    /// Random number generation.
    Random,
    /// XML parsing.
    Xml,
}

impl Capabilities {
    /// Full access — all capabilities granted. For trusted CLI usage.
    pub fn all() -> Self {
        use Capability::*;
        let mut granted = HashSet::new();
        for cap in [Console, FileRead, FileWrite, Http, Sockets, Database,
                    Environment, Gui, Threading, Crypto, Clock, Random, Xml] {
            granted.insert(cap);
        }
        Capabilities { granted }
    }

    /// Safe subset — only pure computation, no I/O or side effects.
    /// Suitable for untrusted code (web playground, sandboxed eval).
    pub fn safe() -> Self {
        use Capability::*;
        let mut granted = HashSet::new();
        for cap in [Console, Clock, Random] {
            granted.insert(cap);
        }
        Capabilities { granted }
    }

    /// No capabilities — pure computation only.
    pub fn none() -> Self {
        Capabilities { granted: HashSet::new() }
    }

    /// Custom: start with none, add specific capabilities.
    pub fn with(caps: &[Capability]) -> Self {
        Capabilities { granted: caps.iter().copied().collect() }
    }

    pub fn has(&self, cap: Capability) -> bool {
        self.granted.contains(&cap)
    }

    pub fn grant(&mut self, cap: Capability) {
        self.granted.insert(cap);
    }

    pub fn revoke(&mut self, cap: Capability) {
        self.granted.remove(&cap);
    }
}

/// Register all standard VSI modules on a VM (no GUI).
/// All capabilities granted.
pub fn register_all(vm: &mut VM) {
    register_with_capabilities(vm, &Capabilities::all());
    // Register no-op GUI stubs so compiled code that emits controlSetProperty/showForm/closeForm
    // doesn't fail with "Unresolved import" in non-GUI contexts.
    if vm.host_registry.get(&("vybe:gui".to_string(), "controlSetProperty".to_string())).is_none() {
        vm.register_host_fn("vybe:gui", "controlSetProperty", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "showForm", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "closeForm", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "noop", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "runApplication", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "onEvent", Box::new(|_ctx, _| Value::Null));
        vm.register_host_fn("vybe:gui", "controlsAdd", Box::new(|_ctx, _| Value::Null));
    }
    // DO NOT call setup_namespaces here — tests override host fns after register_all.
    // setup_namespaces must be called AFTER all host fn registrations.
}

/// Register modules based on granted capabilities.
pub fn register_with_capabilities(vm: &mut VM, caps: &Capabilities) {
    // Always registered — pure computation, no security risk
    math::register(vm);
    string::register(vm);
    array::register(vm);
    convert::register(vm);
    json::register(vm);
    object::register(vm);
    regex::register(vm);
    collections::register(vm);
    runtime::register(vm);
    types::register(vm);
    data::register(vm);
    drawing::register(vm);
    rt::register(vm);

    // Capability-gated modules
    if caps.has(Capability::Console) {
        console::register(vm);
    }
    if caps.has(Capability::Clock) {
        clock::register(vm);
    }
    if caps.has(Capability::Random) {
        random::register(vm);
    }
    if caps.has(Capability::FileRead) || caps.has(Capability::FileWrite) {
        fs::register(vm);
    }
    if caps.has(Capability::Environment) {
        env::register(vm);
    }
    if caps.has(Capability::Http) {
        http::register(vm);
    }
    if caps.has(Capability::Sockets) {
        sockets::register(vm);
    }
    if caps.has(Capability::Database) {
        database::register(vm);
    }
    if caps.has(Capability::Threading) {
        threading::register(vm);
    }
    if caps.has(Capability::Crypto) {
        crypto::register(vm);
    }
    if caps.has(Capability::Xml) {
        xml::register(vm);
    }

    // Set up namespace objects, type registry
    crate::namespaces::setup_namespaces(vm);
    crate::builtin_types::register_all(vm);
}

/// Register all standard VSI modules + GUI module.
/// Returns the shared GuiState — pass it to the form launcher.
#[cfg(feature = "gui")]
pub fn register_all_with_gui(
    vm: &mut VM,
    queue: std::rc::Rc<std::cell::RefCell<crate::SideEffectQueue>>,
) -> std::rc::Rc<std::cell::RefCell<crate::gui_state::GuiState>> {
    let gui = std::rc::Rc::new(std::cell::RefCell::new(crate::gui_state::GuiState::new()));
    register_all(vm);
    gui::register(vm, queue, gui.clone());
    // DO NOT call setup_namespaces here — callers do it after all overrides.
    gui
}

/// Register with capabilities + GUI.
/// Returns the shared GuiState.
#[cfg(feature = "gui")]
pub fn register_with_capabilities_and_gui(
    vm: &mut VM,
    caps: &Capabilities,
    queue: std::rc::Rc<std::cell::RefCell<crate::SideEffectQueue>>,
) -> std::rc::Rc<std::cell::RefCell<crate::gui_state::GuiState>> {
    let gui = std::rc::Rc::new(std::cell::RefCell::new(crate::gui_state::GuiState::new()));
    register_with_capabilities(vm, caps);
    if caps.has(Capability::Gui) {
        gui::register(vm, queue, gui.clone());
        crate::namespaces::setup_namespaces(vm);
    }
    gui
}
