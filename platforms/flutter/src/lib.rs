//! Flutter platform package.
//!
//! A Flutter-shaped adapter over the DOM. To Dart source we *are* Flutter
//! (`Scaffold`, `Column`, `Checkbox`, `Text`, …); under the hood a widget IS an
//! element, built at its construction site through `web:*` — the same surface
//! the dotnet (WinForms) and plib (VCL) adapters render through. No
//! Flutter-specific host functions, no parallel widget runtime.
//!
//! The compiler-side code generation surface lives under [`emitter`].
//!
//! The adapter also owns its Dart *runtime* — `runApp`, `setState`, and the
//! composite inflation (`build`/`createState`) that only the guest can run. It
//! is provided as source ([`runtime_source`]) and compiled into a program ONLY
//! when that program renders (references `runApp`), so widget-only code
//! (construction, `is`-checks, the TDD suite) carries none of it — mirroring
//! how the dotnet adapter emits per-class ctor chunks only for the classes a
//! program uses.

pub mod emitter;

/// The Flutter adapter's Dart runtime: `runApp`, `setState`, composite
/// inflation, and the minimal `EdgeInsets`/`Alignment` value types. Pure Dart
/// over `web:*` — no Flutter-specific host functions. The Dart frontend appends
/// it only when a module references `runApp`.
///
pub fn runtime_source() -> &'static str {
    include_str!("runtime.dart")
}

/// Register this platform with the shared registry.
///
/// The plugin entry point: the host calls this after linking — or, once
/// platforms ship as dylibs, after `dlopen`. `vybe_compiler` must NOT depend on
/// this crate; it reaches the platform through `vybe_runtime::registry`
/// function pointers, exactly as it already does for languages.
pub fn register() {
    // Mount the namespace tree HERE, at plugin-registration time. `vybe_compiler`
    // must never gain a Cargo dependency on this crate — that is what would stop
    // it ever shipping as a dylib.
    emitter::tree_register::register_namespace_tree();
    vybe_runtime::registry::register_platform(vybe_runtime::registry::PlatformDef {
        name: "flutter",
        emit_dispatch: None,
        register_tree: Some(crate::emitter::tree_register::register_namespace_tree),
        namespace_constants: None,
        component_descriptor: None,
        is_descriptor_class: None,
        numeric_format_helper: None,
        read_binary_module: None,
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — the SAME single plugin trait
/// languages use. `init` registers the platform's compiler-facing surface, and
/// this is the dylib entry point. Mirrors `vybe_language_*::Plugin` exactly, so
/// there is one plugin concept in the tree, not two.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "flutter"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
