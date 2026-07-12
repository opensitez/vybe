//! Fortran language support (Modern Fortran 90/95/2003/2008/2018).
//!
//! Pest grammar in `grammar.pest` parses free-form Fortran source.
//! The walker in `walker.rs` converts the parse tree into the common
//! `vybe_compiler::ast::Module`. From there everything goes through the
//! shared compiler — no Fortran-specific code in `compiler.rs`.

pub mod emitter;
pub mod normalize_class;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct FortranParser;

pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded TOML profile source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_plugin::registry::register_language(vybe_plugin::registry::LanguagePlugin {
        name: "fortran",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: None,
    });
}

/// This crate as a [`vybe_plugin::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_plugin::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "fortran"
    }
    fn init(&self, _fw: &mut vybe_plugin::Framework<'_>) {
        register();
    }
}
