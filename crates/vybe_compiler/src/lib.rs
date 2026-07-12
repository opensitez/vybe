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
//! * [`compiler`] — AST-driven dispatch that feeds the emitter.
//! * [`languages`] — per-language walkers + profiles.
//! * [`platforms`] — reusable runtime/framework surfaces such as .NET.

pub use vybe_ast as ast;
pub mod adapters;
pub mod bundle;
pub mod compiler;
pub mod dynamic;
pub mod host_imports;
pub use compiler::Compiler; // avoid the `vybe_compiler::compiler::Compiler` stutter
pub mod emitter;
pub mod languages;
pub mod lsp;
pub mod platforms {
    //! Facade over the platform crates so `crate::platforms::…` paths
    //! keep resolving; the packages live at `platforms/*` in the workspace.
    pub use vybe_platform_dotnet as dotnet;
    pub use vybe_platform_libc as libc;
    pub use vybe_platform_plib as plib;
}
pub use vybe_plugin::profile;
pub mod projects;
pub mod registry; // platform class-library registration (dotnet BCL, pascal plib)

/// Register the built-in (statically-linked) languages into the `vybe_plugin`
/// plugin registry. When a language moves to its own crate it registers itself
/// (via the aggregator) and is removed from this list — the compiler keeps
/// dispatching through the registry either way.
fn register_builtin_languages() {
    use languages::*;
    use vybe_plugin::registry::{LanguagePlugin, register_language};
    macro_rules! reg {
        ($name:literal, $m:ident, emit: $emit:expr, norm: $norm:expr, tree: $tree:expr) => {
            register_language(LanguagePlugin {
                name: $name,
                parse: $m::parse,
                profile_source: $m::profile_source,
                emit_dispatch: $emit,
                normalize_class: $norm,
                register_tree: $tree,
            });
        };
    }
    // `java` is an out-of-crate language (`vybe_lang_java`); it self-registers.
}

/// Idempotent, `Once`-guarded language registration — called before any
/// registry dispatch so built-ins are present regardless of entry point.
pub(crate) fn ensure_languages_registered() {
    static ONCE: std::sync::Once = std::sync::Once::new();
    ONCE.call_once(register_builtin_languages);
}
