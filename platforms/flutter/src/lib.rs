//! Flutter platform package.
//!
//! A Flutter-shaped adapter over the existing `vybe_widgets`/`vybe:gui`
//! runtime. To Dart source we *are* Flutter (`Scaffold`, `Column`,
//! `Checkbox`, `Text`, …); under the hood every widget instantiates and
//! drives the same `vybe_widgets` controls that the dotnet (WinForms) and
//! plib (VCL) adapters already use — no Flutter-specific host functions,
//! no parallel widget runtime.
//!
//! The compiler-side code generation surface lives under [`emitter`].
//!
//! The adapter also owns its Dart *runtime* — the `runApp`/widget-tree
//! realizer that walks the constructed widget config objects and drives
//! `vybe:gui`. It is provided as source ([`runtime_source`]) and compiled into
//! a program ONLY when that program renders (references `runApp`), so
//! widget-only code (construction, `is`-checks, the TDD suite) carries none of
//! it — mirroring how the dotnet adapter emits per-class ctor chunks only for
//! the classes a program uses.

pub mod emitter;

/// The Flutter adapter's Dart runtime: `runApp`, the widget-tree realizer, and
/// the minimal `EdgeInsets`/`Alignment` value types. Pure Dart over the
/// existing `vybe:gui` host — no Flutter-specific host functions. The Dart
/// frontend appends this only when a module references `runApp`.
///
/// The realizer needs two pieces of catalog knowledge at runtime — which widget
/// types are transparent wrappers, and which property keys the backing controls
/// actually act on. Both are GENERATED here from the catalog rather than
/// hand-copied into the Dart source, so the widget modules stay the single
/// source of truth.
pub fn runtime_source() -> &'static str {
    static SOURCE: std::sync::OnceLock<String> = std::sync::OnceLock::new();
    SOURCE.get_or_init(|| {
        let dart_list = |values: &[&str]| {
            values
                .iter()
                .map(|v| format!("\"{v}\""))
                .collect::<Vec<_>>()
                .join(", ")
        };
        format!(
            "{}\n\n\
             // ── Generated from the widget catalog (platforms/flutter/src/emitter/widgets) ──\n\
             var _vfTransparentTypes = [{}];\n\
             var _vfLiveProperties = [{}];\n",
            include_str!("runtime.dart"),
            dart_list(&emitter::catalog::transparent_types()),
            dart_list(emitter::catalog::LIVE_PROPERTIES),
        )
    })
}

/// Register this platform with the shared registry.
///
/// The plugin entry point: the host calls this after linking — or, once
/// platforms ship as dylibs, after `dlopen`. `vybe_compiler` must NOT depend on
/// this crate; it reaches the platform through `vybe_runtime::registry`
/// function pointers, exactly as it already does for languages.
pub fn register() {
    // Mount the namespace tree HERE, at plugin-registration time. Previously
    // `vybe_compiler` called this from `resolver.rs` via a hard Cargo
    // dependency, which is what prevented this crate from ever being a dylib.
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
