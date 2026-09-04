//! Common AST targeting WASM bytecode.
//!
//! Every language parser (pest-based) produces this AST. One compiler consumes it.
//! The AST represents semantic operations, not syntax — language-specific syntax
//! is resolved by each language's tree walker before reaching this level.
//!
//! ## Modules
//! * [`emitter`] — the WASM bytecode emission layer (dict, collections, strings,
//!   loops, classes, stdlib, invoke, dotnet). Was the `vybe_compiler_common`
//!   crate; moved in-tree now that it has a single consumer.
//! * [`primitives`] — shared AST-driven primitives and bytecode emission.
//! * [`languages`] — per-language walkers + profiles.
//! * [`platforms`] — reusable runtime/framework surfaces such as .NET.

// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub use vybe_ast as ast;
pub mod adapters;
pub mod bundle;
pub mod component_classes;
pub mod dynamic;
pub mod host_imports;
pub mod primitives;
pub use primitives::Compiler;
pub mod languages;
pub mod lsp;
pub mod platforms {
    //! Facade over the platform crates so `crate::platforms::…` paths keep
    //! resolving.
    //!
    //! Platform REGISTRATION no longer happens here — a platform registers
    //! itself from `Plugin::init`, driven by the language that needs it
    //! (`vybe_language_dart::register()` -> `vybe_platform_flutter::register()`),
    //! so the compiler never mounts a platform tree and a platform can become a
    //! dylib.
    //!
    //! What remains are direct DATA references (dotnet `numeric_format`,
    //! `winforms`, `component_descriptor`, `namespace_constant_mappings`; plib
    //! `gcl`). Route those through the registry and these dependencies — and
    //! this facade — go away entirely.
}
pub use vybe_runtime::profile;
pub mod projects;
pub mod registry; // platform class-library registration (dotnet BCL, pascal plib)

/// Register the built-in (statically-linked) languages into the `vybe_plugin`
/// plugin registry. When a language moves to its own crate it registers itself
/// (via the aggregator) and is removed from this list — the compiler keeps
/// dispatching through the registry either way.
fn register_builtin_languages() {
    // `java` is an out-of-crate language (`vybe_lang_java`); it self-registers.
    // Built-in registrations are now handled by the language crates' own registrars.
}

/// Idempotent, `Once`-guarded language registration — called before any
/// registry dispatch so built-ins are present regardless of entry point.
pub(crate) fn ensure_languages_registered() {
    static ONCE: std::sync::Once = std::sync::Once::new();
    ONCE.call_once(register_builtin_languages);
}
