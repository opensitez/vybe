// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod emitter;
mod normalize_entry;
mod normalize_exact_word_regions;
mod normalize_shutdown;
pub mod normalize_affine_bytes;
mod normalize_linear_bytes;
mod normalize_storage_access;
mod normalize_loop_backing;
mod normalize_word_provenance;
mod normalize_static_helpers;
mod normalize_type_guards;
pub mod tree_register;
pub mod walker;
pub mod frontend;

#[cfg(test)]
mod streaming_unit_tests;

#[cfg(test)]
mod sdl_header_tests;

mod native_grammar {
    include!(concat!(env!("OUT_DIR"), "/native_parser.rs"));
}
pub(crate) use native_grammar::Rule;
pub(crate) struct CParser;
impl CParser {
    fn options(source: &str) -> grameo::MatchOptions {
        let mut options = grameo::MatchOptions::default();
        options.max_steps = options.max_steps.max(source.len().saturating_mul(1024));
        options
    }

    pub(crate) fn parse(rule: Rule, source: &str) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        // Retain an uncapped Grameo default; if its default becomes finite,
        // scale that work budget for multi-megabyte preprocessed C units.
        frontend::parse_captures(rule, source, Self::options(source))
    }

    pub(crate) fn parse_top_level(source: &str) -> Result<Vec<grameo::compat::Pair<'_, Rule>>, grameo::ParseError> {
        frontend::parse_top_level_units(source, Self::options(source))
    }
}

/// Parse C source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    frontend::parse_registered(source)
}

/// Explicit PEG reference, including the existing ordered contextual lowering.
pub fn parse_with_peg(source: &str) -> Result<vybe_ast::Module, String> {
    frontend::with_backend(frontend::Backend::Peg, || walker::parse(source))
}

/// Native matching; C contextual lowering remains unchanged.
pub fn parse_optimized(source: &str) -> Result<vybe_ast::Module, String> {
    frontend::with_backend(frontend::Backend::Native, || walker::parse(source))
}

/// Opt-in owned scalar Pratt construction with contextual recapture.
/// This does not provide streaming compilation or per-function commits.
pub fn parse_owned_ast(source: &str) -> Result<vybe_ast::Module, String> {
    frontend::with_backend(frontend::Backend::Owned, || walker::parse(source))
}

/// Explicit mixed binary construction with scope-sensitive unary lowering.
/// This remains experimental and does not provide per-unit streaming commits.
pub fn parse_owned_contextual(source: &str) -> Result<vybe_ast::Module, String> {
    frontend::with_backend(frontend::Backend::OwnedContextual, || walker::parse(source))
}

/// Construct binary ASTs with Grameo Pratt at the existing typed C scope boundary.
/// Matching still retains the whole capture tree; streaming is not implemented.
pub fn parse_contextual_pratt(source: &str) -> Result<vybe_ast::Module, String> {
    frontend::with_backend(frontend::Backend::ContextualPratt, || walker::parse(source))
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

fn selected_parser(selection: Option<&str>) -> vybe_runtime::registry::ParseFn {
    match frontend::selected_backend(selection) {
        frontend::Backend::Native => parse_optimized,
        frontend::Backend::Owned => parse_owned_ast,
        frontend::Backend::OwnedContextual => parse_owned_contextual,
        frontend::Backend::ContextualPratt => parse_contextual_pratt,
        frontend::Backend::Peg => parse_with_peg,
    }
}

fn compiler_optimizations(selected: vybe_runtime::registry::ParseFn) -> vybe_runtime::registry::CompilerOptimizations {
    let optimized = std::ptr::fn_addr_eq(selected, parse_owned_ast as vybe_runtime::registry::ParseFn)
        || std::ptr::fn_addr_eq(selected, parse_owned_contextual as vybe_runtime::registry::ParseFn);
    if optimized || std::ptr::fn_addr_eq(selected, parse_optimized as vybe_runtime::registry::ParseFn)
        || std::ptr::fn_addr_eq(selected, parse_with_peg as vybe_runtime::registry::ParseFn)
        || std::ptr::fn_addr_eq(selected, parse_contextual_pratt as vybe_runtime::registry::ParseFn) {
        vybe_runtime::registry::CompilerOptimizations {
            symbol_membership_index: true,
            user_namespace_cache: true,
            completion_worker_lowering: optimized,
            ..Default::default()
        }
    } else {
        Default::default()
    }
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
        parse: selected_parser(std::env::var("VYBE_C_FRONTEND").ok().as_deref()),
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: None,
        register_tree: Some(tree_register::register_namespace_tree),
        expand_source: Some(walker::expand_source),
    });
    vybe_runtime::registry::register_hooks(
        "c",
        vybe_runtime::registry::LanguageHooks {
            compiler_optimizations: Some(compiler_optimizations),
            ..Default::default()
        },
    );
}

#[cfg(test)]
mod immutable_binding_tests;

#[cfg(test)]
mod wide_read_tests;

#[cfg(test)]
mod float_ratio_tests;

mod goto_lowering;

mod goto_prefix_policy;

#[cfg(test)]
mod goto_prefix_integration_tests;

#[cfg(test)]
mod goto_lowering_c_tests;

#[cfg(test)]
mod nested_goto_tests;

#[cfg(test)]
mod frontend_capability_tests {
    use super::*;

    #[test]
    fn lookup_acceleration_is_independent_of_matching_backend() {
        for selection in [None, Some("native"), Some("peg"), Some("unknown"), Some("pratt")] {
            let options = compiler_optimizations(selected_parser(selection));
            assert!(options.symbol_membership_index);
            assert!(options.user_namespace_cache);
            assert!(!options.completion_worker_lowering);
        }
        for selection in [Some("owned"), Some("mixed")] {
            let options = compiler_optimizations(selected_parser(selection));
            assert!(options.symbol_membership_index);
            assert!(options.user_namespace_cache);
            assert!(options.completion_worker_lowering);
        }
        assert_eq!(compiler_optimizations(parse), Default::default());
    }

    #[test]
    fn native_is_default_and_peg_remains_explicit() {
        assert!(std::ptr::fn_addr_eq(selected_parser(None), parse_optimized as vybe_runtime::registry::ParseFn));
        for selection in [Some("peg"), Some("unknown")] {
            assert!(std::ptr::fn_addr_eq(selected_parser(selection), parse_with_peg as vybe_runtime::registry::ParseFn));
        }
        assert_eq!(compiler_optimizations(parse_optimized), compiler_optimizations(parse_with_peg));
    }

    #[test]
    fn production_lookup_acceleration_preserves_c_emission() {
        register();
        let profile = std::sync::Arc::new(
            vybe_compiler::profile::parse_profile(profile_source()).unwrap(),
        );
        for source in [
            "int g=3; int f(int x){return x+g;} int main(void){return f(4);}",
            "struct P {int x;}; int f(struct P *p){p->x+=2; return p->x;} int main(void){struct P p={3}; return f(&p);}",
            "typedef unsigned int word; word f(word x){return (x>>3)==7;} int main(void){return f(56);}",
        ] {
            let module = parse_with_peg(source).unwrap();
            let native_module = parse_optimized(source).unwrap();
            assert_eq!(format!("{module:?}"), format!("{native_module:?}"));
            let compile = |options| {
                vybe_compiler::primitives::Compiler::with_shared_profile(profile.clone())
                    .with_optimizations(options)
                    .compile_with_imports(&module)
                    .unwrap()
            };
            let baseline = compile(Default::default());
            let accelerated = compile(compiler_optimizations(parse_with_peg));
            let native = vybe_compiler::primitives::Compiler::with_shared_profile(profile.clone())
                .with_optimizations(compiler_optimizations(selected_parser(None)))
                .compile_with_imports(&native_module)
                .unwrap();
            assert_eq!(baseline.chunks.len(), accelerated.chunks.len());
            assert_eq!(accelerated.chunks.len(), native.chunks.len());
            for (left, right) in baseline.chunks.iter().zip(&accelerated.chunks) {
                assert_eq!(left.name, right.name);
                assert_eq!(left.code, right.code, "{}", left.name);
            }
            let write = |chunks: &[vybe_runtime::Chunk]| {
                vybe_platform_wasm::write_wasm_with_options(
                    chunks,
                    vybe_platform_wasm::WasmWriteOptions {
                        include_vybe_metadata: false,
                        include_name_section: true,
                    },
                )
            };
            let bytes = write(&accelerated.chunks);
            assert_eq!(write(&baseline.chunks), bytes);
            assert_eq!(write(&native.chunks), bytes);
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap();
        }
    }
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
