// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
pub mod emitter;
pub mod walker;

use pest_derive::Parser;

#[derive(Parser)]
#[grammar = "src/grammar.pest"]
pub struct WashmParser;

/// Parse Bash/washm source into the common AST.
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
        name: "washm",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: None,
        register_tree: None,
        expand_source: None,
    });
    vybe_runtime::registry::register_language(vybe_runtime::registry::LanguageDef {
        name: "bash",
        parse,
        profile_source,
        emit_dispatch: Some(emitter::dispatch::dispatch),
        normalize_class: None,
        register_tree: None,
        expand_source: None,
    });
}

/// This crate as a [`vybe_runtime::Plugin`] — its `init` registers the
/// language with the shared framework.
pub struct Plugin;
impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "washm"
    }
    fn init(&self, _fw: &mut vybe_runtime::Framework<'_>) {
        register();
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
vybe_runtime::register_plugin!(Plugin);

#[cfg(test)]
mod tests {
    use super::*;
    use pest::Parser;

    #[test]
    fn test_parse_pipeline_subst() {
        let code = "res=$(printf '%s' \"routed\" | # comment\n cat)\n";
        let res = parse(code);
        assert!(res.is_ok(), "{:?}", res.err());
    }

    #[test]
    fn test_double_quote_line() {
        let li = WashmParser::parse(
            Rule::list,
            "[ \"$tag\" = \"#status\" ] || fail \"hash after equal: want '#status', got [$tag]\"\n",
        );
        assert!(li.is_ok(), "list error: {:?}", li.err());
        let prg = WashmParser::parse(
            Rule::program,
            "[ \"$tag\" = \"#status\" ] || fail \"hash after equal: want '#status', got [$tag]\"\n",
        );
        assert!(prg.is_ok(), "program error: {:?}", prg.err());
    }

    #[test]
    fn test_param_length_and_prefix() {
        let code = "len=${#str}\n[ \"$len\" -eq 7 ] || fail \"param length\"\n";
        assert!(WashmParser::parse(Rule::program, code).is_ok());

        let code2 = "trim1=${path#*/}\ntrim2=${path##*/}\n";
        assert!(WashmParser::parse(Rule::program, code2).is_ok());
    }

    #[test]
    fn test_arithmetic_in_loop() {
        let code = "acc=$((acc + n)) # accumulate\n";
        assert!(WashmParser::parse(Rule::program, code).is_ok());
    }

    #[test]
    fn test_array_with_comments() {
        let code =
            "items=(\n apple # first\n banana # second\n)\n[ \"${items[0]}\" = \"apple\" ]\n";
        assert!(WashmParser::parse(Rule::program, code).is_ok());
    }

    #[test]
    fn test_heredoc_body() {
        let code = "doc=$(cat <<'EOF'\n# not a comment\ndata line\nEOF\n)\nexpected=\"# not a comment\"$'\\n'\"data line\"\n";
        assert!(WashmParser::parse(Rule::program, code).is_ok());
    }
}
