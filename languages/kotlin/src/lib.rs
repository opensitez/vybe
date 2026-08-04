// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));

pub mod emitter;
pub mod normalize_class;
pub mod protocol;
pub mod tree_register;
mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub(crate) struct KotlinParser;

/// Parse Kotlin source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    // Parse on a thread with a LARGE stack. The expression grammar is ~17
    // levels deep and `walk_expr`'s frames are big in debug builds, so a
    // moderately nested expression (`(xs.zip(listOf(1))).toString()`) walks a
    // pair tree deep enough to blow the default 8 MiB main stack — measured
    // as a hard overflow, not a hang. The documented alternative (flattening
    // the precedence chain) caps how faithful the grammar can be; a worker
    // thread with headroom does not.
    let source = source.to_string();
    std::thread::Builder::new()
        .name("kotlin-parse".into())
        .stack_size(256 * 1024 * 1024)
        .spawn(move || walker::parse(&source))
        .map_err(|e| format!("parse thread: {e}"))?
        .join()
        .map_err(|_| "Kotlin parse thread panicked".to_string())?
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry.
pub fn register() {
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "kotlin",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: Some(tree_register::register_namespace_tree),
        expand_source: None,
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language with the shared framework.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "kotlin"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the registry.
vybe_runtime::register_plugin!(Plugin);
