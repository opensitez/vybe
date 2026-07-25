//! The bundled host **platform** plugins, and the one call that registers them
//! onto a VM.
//!
//! Every provider in Vybe is a [`vybe_bytecode::Plugin`] and they all run
//! through the single loop (`init_all_on_vm_with_caps`). The *full* plugin list
//! — languages + platforms — lives at the app layer (`vybex`), which is the only
//! crate that depends on every language. This module is the platform half, kept
//! here because `vybe_emitter` is the lowest crate that already depends on all
//! five platform crates.
//!
//! Use it wherever a VM needs host functions but not language descriptors:
//! emit-time host-fn validation, eval/mini-VMs, and test harnesses.

use vybe_bytecode::VM;
use vybe_bytecode::capabilities::Capabilities;

/// Register every bundled host platform (`ecma`, `web`, `wasi`, `node`,
/// `vybe`) onto `vm` via the one plugin loop, gated by `caps`.
///
/// Each platform's `Plugin::init` registers its host functions and its
/// `finalize` registers its runtime types (ECMA globals + TypeRegistry
/// vtables), in the framework's two-phase pass.
pub fn register_platforms(vm: &mut VM, caps: &Capabilities) {
    let vybe = vybe_platform_vybe::Plugin::new();
    let plugins: [&dyn vybe_bytecode::Plugin; 5] = [
        &vybe_platform_ecma::Plugin,
        &vybe_platform_web::Plugin,
        &vybe_platform_wasi::Plugin,
        &vybe_platform_node::Plugin,
        &vybe,
    ];
    vybe_bytecode::init_all_on_vm_with_caps(vm, caps, &plugins);
}

/// [`register_platforms`] with every capability granted, plus the `vybe:gui`
/// no-op stubs so compiled control/form code links in non-GUI contexts.
pub fn register_platforms_all(vm: &mut VM) {
    register_platforms(vm, &Capabilities::all());
    gui_stubs(vm);
}

/// Phase 1 only — run every platform's `Plugin::init` (host functions), and
/// the `vybe:gui` stubs, but NOT the `finalize` pass.
///
/// Use this when host functions must be **overridden before** globals are
/// wired: `register_host_fn` appends a *new* index rather than replacing, and
/// the ecma `finalize` stamps `globalThis`/prototypes by index — so an override
/// registered after `finalize` would not be seen by those globals. Test
/// harnesses that capture output do:
/// [`init_platforms`] → override host fns → [`finalize_platforms`].
pub fn init_platforms(vm: &mut VM) {
    let caps = Capabilities::all();
    with_platforms(|plugins| vybe_bytecode::init_plugins(vm, &caps, plugins));
    gui_stubs(vm);
}

/// Phase 2 only — run every platform's `Plugin::finalize` (ECMA globals +
/// TypeRegistry vtables). Pair with [`init_platforms`].
pub fn finalize_platforms(vm: &mut VM) {
    let caps = Capabilities::all();
    with_platforms(|plugins| vybe_bytecode::finalize_plugins(vm, &caps, plugins));
}

/// [`init_platforms`] with the **widget-backed** `vybe:gui` surface: the `vybe`
/// plugin owns a fresh `GuiState`, and the shared handle is returned for the
/// form launcher / assertions. Phase 1 only — call [`finalize_platforms`] after
/// any host-fn overrides.
#[cfg(feature = "gui")]
pub fn init_platforms_with_gui(
    vm: &mut VM,
) -> std::sync::Arc<std::sync::Mutex<vybe_platform_vybe::gui_state::GuiState>> {
    let caps = Capabilities::all();
    let vybe = vybe_platform_vybe::Plugin::with_gui();
    let plugins: [&dyn vybe_bytecode::Plugin; 5] = [
        &vybe_platform_ecma::Plugin,
        &vybe_platform_web::Plugin,
        &vybe_platform_wasi::Plugin,
        &vybe_platform_node::Plugin,
        &vybe,
    ];
    vybe_bytecode::init_plugins(vm, &caps, &plugins);
    vybe.gui_state().expect("with_gui() always creates a GuiState")
}

/// The `vybe:gui` no-op stubs, installed only when the real widget-backed
/// surface isn't present.
fn gui_stubs(vm: &mut VM) {
    if vm
        .host_registry
        .get(&("vybe:gui".to_string(), "controlSetProperty".to_string()))
        .is_none()
    {
        vybe_platform_vybe::register_gui_stubs(vm);
    }
}

/// Build the platform plugin list and hand it to `f` (the `vybe` plugin is a
/// value, so it must outlive the borrow).
fn with_platforms<R>(f: impl FnOnce(&[&dyn vybe_bytecode::Plugin]) -> R) -> R {
    let vybe = vybe_platform_vybe::Plugin::new();
    let plugins: [&dyn vybe_bytecode::Plugin; 5] = [
        &vybe_platform_ecma::Plugin,
        &vybe_platform_web::Plugin,
        &vybe_platform_wasi::Plugin,
        &vybe_platform_node::Plugin,
        &vybe,
    ];
    f(&plugins)
}
