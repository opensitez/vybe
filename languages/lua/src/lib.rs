// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod emitter;
mod normalize;
pub mod normalize_class;
mod protocol;
pub mod tree_register;
pub mod walker;

#[cfg(all(not(feature = "native-parser"), feature = "legacy-parser"))]
use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
#[cfg(all(not(feature = "native-parser"), feature = "legacy-parser"))]
pub(crate) struct LuaParser;

#[cfg(feature = "native-parser")]
mod native_grammar {
    include!(concat!(env!("OUT_DIR"), "/native_parser.rs"));
}
#[cfg(feature = "native-parser")]
pub(crate) use native_grammar::Rule;
#[cfg(feature = "native-parser")]
pub(crate) struct LuaParser;
#[cfg(feature = "native-parser")]
impl LuaParser {
    pub(crate) fn parse(rule: Rule, source: &str) -> Result<vybe_parser::compat::Pairs<'_, Rule>, vybe_parser::ParseError> {
        native_grammar::Parser.parse_pairs(rule, source)
    }
}
#[cfg(not(any(feature = "native-parser", feature = "legacy-parser")))]
compile_error!("Lua requires either the native-parser or legacy-parser feature");

/// Parse Lua source into the common AST.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Embedded profile TOML source.
pub fn profile_source() -> &'static str {
    include_str!("profile")
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "lua",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: Some(normalize_class::normalize_class),
        register_tree: Some(tree_register::register_namespace_tree),
        expand_source: None,
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language (and any forms) with the shared framework. Also the dylib entry point.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "lua"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);

#[cfg(all(test, feature = "native-parser", feature = "legacy-parser"))]
mod parser_conformance {
    mod legacy {
        use crate::normalize;
        #[derive(pest_derive::Parser)]
        #[grammar = "src/grammar.pest"]
        struct LuaParser;
        mod walker { include!(concat!(env!("OUT_DIR"), "/legacy_test_walker.rs")); }
        pub fn parse(source: &str) -> Result<vybe_ast::Module, String> { walker::parse(source) }
    }

    #[test]
    fn normalized_lua_modules_match_reference_backend() {
        for source in [
            "", "local x=1+2*3; return x", "local a,b=f(); return a,b",
            "function f(a, ...) return a, ... end; return f(2)",
            "for i=1,3 do print(i) end", "for i=1,3 do f=function() return i end end",
            "local t={x=1,[2]=3,4}; return t.x,t[2]",
            "function t:m(x) return self.x+x end; t:m(1)",
            "if a then f() else g() end", "while a do a=a-1 end",
            "return not a and b or c, -2^2, a .. b .. c, a ~= b",
            "::again:: goto again", "local s='é😀'; return s",
        ] {
            let native = crate::parse(source).unwrap_or_else(|e| panic!("native {source}: {e}"));
            let reference = legacy::parse(source).unwrap_or_else(|e| panic!("reference {source}: {e}"));
            // The independent conformance package compares raw AST fields
            // structurally. Here full normalized output (including all spans
            // and canonical metadata) is also compared without omitting fields.
            assert_eq!(format!("{native:#?}"), format!("{reference:#?}"), "{source}");
        }
        for source in ["local x=", "return 1+", "function(", "if a then"] {
            assert!(crate::parse(source).is_err());
            assert!(legacy::parse(source).is_err());
        }
    }

    #[test]
    fn native_and_reference_lua_modules_emit_identical_wasm_bytes() {
        crate::register();
        for source in ["return 1+2*3", "local x=2; return x^3", "local function f(a) return a+1 end; return f(4)"] {
            let emit = |module: &vybe_ast::Module| {
                let profile = vybe_compiler::profile::parse_profile(crate::profile_source()).unwrap();
                vybe_compiler::Compiler::with_profile(profile).compile(module).unwrap()
            };
            let native = emit(&crate::parse(source).unwrap());
            let reference = emit(&legacy::parse(source).unwrap());
            assert!(!native.is_empty());
            assert_eq!(native.len(), reference.len());
            for (a,b) in native.iter().zip(&reference) {
                assert_eq!(a.name,b.name);
                assert_eq!(a.code,b.code, "{source}: {}", a.name);
                assert_eq!(format!("{:?}",a.constants),format!("{:?}",b.constants));
            }
        }
    }
}
