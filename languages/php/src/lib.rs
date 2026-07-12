//! PHP language support.
//!
//! Pest grammar in `grammar.pest` parses PHP 8 source. The walker in
//! `walker.rs` converts the parse tree into the common
//! `vybe_compiler::ast::Module`. From there everything goes through the
//! shared compiler — no PHP-specific code in `compiler.rs`.

pub mod emitter;
pub mod normalize_class;
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

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_plugin::registry::register_language(vybe_plugin::registry::LanguagePlugin {
        name: "php",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: Some(tree_register::register_namespace_tree),
    });
    vybe_plugin::registry::register_hooks("php", vybe_plugin::registry::LanguageHooks {
        relational_compare: Some(emitter::relational_adapter::emit_relational_compare),
        constructor_ref_autoload: Some(emitter::autoload_adapter::emit_constructor_ref_with_autoload),
        dynamic_constructor_ref_autoload: Some(emitter::autoload_adapter::emit_dynamic_constructor_ref_with_autoload),
        normalize_source: Some(normalize_source_for_parser),
        str_getcsv: Some(emitter::string_adapter::emit_str_getcsv),
        ..Default::default()
    });
}

/// This crate as a [`vybe_plugin::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_plugin::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "php"
    }
    fn init(&self, _fw: &mut vybe_plugin::Framework<'_>) {
        register();
    }
}
