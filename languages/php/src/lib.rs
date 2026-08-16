//! PHP language support.
//!
//! Pest grammar in `grammar.pest` parses PHP 8 source. The walker in
//! `walker.rs` converts the parse tree into the common
//! `vybe_compiler::ast::Module`. From there everything goes through the
//! shared compiler — no PHP-specific code in `compiler.rs`.

// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod emitter;
pub mod normalize_class;
pub mod protocol;
pub mod tree_register;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct PhpParser;

pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

pub(crate) fn normalize_source_for_parser(source: &str) -> String {
    walker::normalize_source_for_parser(source)
}

/// Embedded TOML profile source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

// ── PHP's variable namespace ────────────────────────────────────────────────
//
// PHP keeps variables in a namespace separate from functions and classes: `$x`
// and `x()` are unrelated bindings that can coexist in one program. The `$` is
// PHP's spelling of that distinction, so all knowledge of it lives here and
// none of it reaches shared code. See `vybe_runtime::registry::VariableNamespace`.

fn variable_body(name: &str) -> &str {
    name.strip_prefix('$').unwrap_or(name)
}

fn spell_variable(body: &str) -> String {
    format!("${body}")
}

fn variable_global_key(name: &str, canon: &str) -> Option<String> {
    if !name.starts_with('$') {
        return None;
    }
    // Superglobals (`$_GET`, `$_SERVER`, `$_POST`, …) are registered by the
    // host under their literal source spelling, so mangling them would make
    // them unreachable. Everything else gets a prefix that keeps a global
    // `$foo` from colliding with a function named `foo`.
    if name.starts_with("$_") {
        Some(name.to_string())
    } else {
        Some(format!("__php_var_{canon}"))
    }
}

static VARIABLE_NAMESPACE: vybe_runtime::registry::VariableNamespace =
    vybe_runtime::registry::VariableNamespace {
        body: variable_body,
        spell: spell_variable,
        global_key: variable_global_key,
    };

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "php",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: Some(tree_register::register_namespace_tree),
        expand_source: None,
    });
    vybe_runtime::registry::register_hooks(
        "php",
        vybe_runtime::registry::LanguageHooks {
            variable_namespace: Some(&VARIABLE_NAMESPACE),
            constructor_ref_autoload: Some(
                emitter::autoload_adapter::emit_constructor_ref_with_autoload,
            ),
            dynamic_constructor_ref_autoload: Some(
                emitter::autoload_adapter::emit_dynamic_constructor_ref_with_autoload,
            ),
            normalize_source: Some(normalize_source_for_parser),
            ..Default::default()
        },
    );
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "php"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
    /// Drop the walker's per-program registries at the TENANT BOUNDARY.
    ///
    /// `VM::reset_to` rolls back everything the VM owns, but the front end runs
    /// in this process before any VM state exists, so its `thread_local!`
    /// registries were never in that reset's reach — `CLASS_REGISTRY` among
    /// them, which answers every inheritance, abstractness and interface
    /// question the walker asks. In a reused VM one program's classes were
    /// still visible to the next.
    ///
    /// This hook is where the framework already expects that to be handled:
    /// `Plugin::reset` is documented as "process-global state a plugin owns
    /// directly, which is exactly why the VM's own reset could never reach it."
    fn reset(&self) {
        crate::walker::reset_program_state();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
