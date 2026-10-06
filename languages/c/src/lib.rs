// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod emitter;
mod normalize_entry;
mod normalize_type_guards;
pub mod tree_register;
pub mod walker;

#[allow(dead_code)]
mod native_grammar {
    include!(concat!(env!("OUT_DIR"), "/native_parser.rs"));
}
pub(crate) use native_grammar::Rule;
pub(crate) struct CParser;
impl CParser {
    pub(crate) fn parse(rule: Rule, source: &str) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        let mut options = grameo::MatchOptions::default();
        // Preprocessed translation units can contain millions of source bytes.
        // Keep a finite linear work budget rather than a fixed small-file ceiling.
        options.max_steps = options.max_steps.max(source.len().saturating_mul(1024));
        grameo::engine::parse_program(&native_grammar::Parser, rule.name(), source, options)
            .map(|tree| tree.into_typed_pairs())
    }
}

/// Parse C source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    // Platforms this language needs. A language already links them (see
    // Cargo.toml), so registering them here means the compiler never has to —
    // and it works from ANY host, including language test binaries that never
    // construct vybex. See flexclassplan.md.
    vybe_platform_libc::register();

    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "c",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: None,
        register_tree: Some(tree_register::register_namespace_tree),
        expand_source: Some(walker::expand_source),
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "c"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
