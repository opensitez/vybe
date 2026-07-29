//! Driving the ONE plugin registration loop for a VM.
//!
//! There is no platform list here, and no plugin list anywhere in the tree.
//! Every plugin crate — language, platform, host-function provider, they are
//! all the same `vybe_runtime::Plugin` — submits itself at link time via
//! `register_plugin!`, and these calls run whatever the final binary linked.
//! Which plugins a binary has is a Cargo question; running them is one loop.
//!
//! Use it wherever a VM needs host functions: emit-time host-fn validation,
//! eval/mini-VMs, and test harnesses.

use vybe_runtime::VM;
use vybe_runtime::capabilities::Capabilities;

/// Register every linked plugin onto `vm` via the one loop, gated by `caps`.
///
/// Each plugin's `Plugin::init` registers its host functions and its
/// `finalize` registers its runtime types (ECMA globals + TypeRegistry
/// vtables), in the framework's two-phase pass.
pub fn register_platforms(vm: &mut VM, caps: &Capabilities) {
    vybe_runtime::init_all_registered(vm, caps);
}

/// [`register_platforms`] with every capability granted, plus the `vybe:gui`
/// no-op stubs so compiled control/form code links in non-GUI contexts.
pub fn register_platforms_all(vm: &mut VM) {
    register_platforms(vm, &Capabilities::all());
}

/// Phase 1 only — every plugin's `Plugin::init` (host functions), and the
/// `vybe:gui` stubs, but NOT the `finalize` pass.
///
/// Use this when host functions must be **overridden before** globals are
/// wired: `register_host_fn` appends a *new* index rather than replacing, and
/// the ecma `finalize` stamps `globalThis`/prototypes by index — so an override
/// registered after `finalize` would not be seen by those globals. Test
/// harnesses that capture output do:
/// [`init_platforms`] → override host fns → [`finalize_platforms`].
pub fn init_platforms(vm: &mut VM) {
    vybe_runtime::init_registered_plugins(vm, &Capabilities::all());
}

/// Phase 2 only — every plugin's `Plugin::finalize` (ECMA globals +
/// TypeRegistry vtables). Pair with [`init_platforms`].
pub fn finalize_platforms(vm: &mut VM) {
    vybe_runtime::finalize_registered_plugins(vm, &Capabilities::all());
}

// The `vybe:gui` no-op stubs are installed by the `vybe` plugin itself when no
// widget surface is present — see its `Plugin::init`. Nothing here names a
// platform crate, which is what lets the emit layer live in the compiler.
