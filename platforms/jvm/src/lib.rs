//! JVM-shaped platform package — the JDK class libraries.
//!
//! `java.util.List`, `java.time.Instant`, `java.io.File` are the JDK's own
//! names: the identical token in Java, Kotlin, Scala and Groovy. They are
//! declared ONCE here and reached through the COMMON RESOLVER, so a language
//! that wants them declares no `java.*` entries at all — it declares tree data
//! (`type_scopes` + `kind = "tree-ambient"`) in its profile, exactly as
//! csharp/vb reach `dotnet.*` with zero `System.*` entries of their own.
//!
//! Before this crate, every `java.*` name lived in `languages/java`, which made
//! the JDK the property of one frontend and meant a second JVM language would
//! have to declare the whole surface again.

pub mod emitter;

/// Embedded `java.*` profile fragment — the DATA the namespace tree is built
/// from. A platform ships its own declarations for the same reason a language
/// does: the registrar and the declarations are one artifact and must not be
/// separated.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this platform with the shared registry.
///
/// The plugin entry point: the host calls this after linking — or, once
/// platforms ship as dylibs, after `dlopen`. `vybe_compiler` must NOT depend on
/// this crate; it reaches the platform through `vybe_runtime::registry`
/// function pointers, exactly as it already does for languages.
pub fn register() {
    // Mount the namespace tree HERE, at plugin-registration time, so the root
    // exists as soon as the platform is linked — no language has to ask for it.
    emitter::tree_register::register_namespace_tree();
    vybe_runtime::registry::register_platform(vybe_runtime::registry::PlatformDef {
        name: "jvm",
        emit_dispatch: Some(crate::emitter::dispatch::dispatch),
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
/// this is the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "jvm"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
