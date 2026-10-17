//! PHP language support.
//!
//! The grammar in `grammar.pest` parses PHP 8 source. The walker in
//! `walker.rs` converts the parse tree into the common
//! `vybe_compiler::ast::Module`. From there everything goes through the
//! shared compiler — no PHP-specific code in `compiler.rs`.

// Force-link every plugin crate in `[dependencies]` so its link-time
// registration reaches the registry. Generated from Cargo.toml — see build.rs.
include!(concat!(env!("OUT_DIR"), "/linked_plugins.rs"));
mod committed_unit;
pub use committed_unit::{CommittedSourceUnit, GrammarCommit, parse_committed_source_unit};
mod array_read;
mod cookies;
pub(crate) mod core_exceptions;
mod datetime;
mod dynamic_symbols;
pub mod emitter;
mod environment;
mod filesystem;
mod globals;
mod logging;
pub mod normalize_class;
mod numeric;
mod pdo_rows;
pub mod protocol;
mod random;
mod regex;
mod serialization;
mod sodium;
mod spl_filesystem;
mod spl_classes;
mod string_bytes;
mod html_entities;
mod html_entity_names;
mod string_scan;
pub mod tree_register;
pub mod walker;

#[cfg(test)]
#[path = "../build_support/pratt.rs"]
mod pratt_binding_tests;

mod native_grammar {
    include!(concat!(env!("OUT_DIR"), "/native_parser.rs"));
}
pub(crate) use native_grammar::Rule;
pub(crate) struct PhpParser;
impl PhpParser {
    pub(crate) fn parse(rule: Rule, source: &str) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
        native_grammar::Parser.parse_pairs(rule, source)
    }
}

// Verification-only baseline selection, read once when an AST builder is
// created. No TLS or policy branch is compiled into ordinary builds.
#[cfg(feature = "service-verification")]
std::thread_local! {
    static PHP_PRATT_ADMISSION_CACHE_FOR_VALIDATION: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}
#[cfg(feature = "service-verification")]
pub fn with_pratt_admission_cache_for_validation<T>(enabled: bool, run: impl FnOnce() -> T) -> T {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) { PHP_PRATT_ADMISSION_CACHE_FOR_VALIDATION.with(|flag| flag.set(self.0)); }
    }
    let restore = Restore(PHP_PRATT_ADMISSION_CACHE_FOR_VALIDATION.with(|flag| flag.replace(enabled)));
    let result = run();
    drop(restore);
    result
}

/// Actual frontend admission requests, collected only inside an explicit
/// verification scope. Keys include the builder identity: equal offsets in
/// different source units cannot share an admission result.
#[cfg(feature = "service-verification")]
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct PrattAdmissionCensus {
    pub builders: usize,
    pub requests: usize,
    pub distinct_requests: usize,
    pub repeated_requests: usize,
    pub scans: usize,
    pub hits: usize,
}
#[cfg(feature = "service-verification")]
#[derive(Default)]
struct PrattAdmissionCensusState {
    census: PrattAdmissionCensus,
    requests: std::collections::HashSet<(usize, usize, bool)>,
}
#[cfg(feature = "service-verification")]
std::thread_local! {
    static PHP_PRATT_ADMISSION_CENSUS: std::cell::RefCell<Option<PrattAdmissionCensusState>> = const { std::cell::RefCell::new(None) };
}
/// Observe real public-parser admission traffic. Nested scopes and unwinding
/// restore the previous collector; ordinary builds contain no collector.
#[cfg(feature = "service-verification")]
pub fn with_pratt_admission_census_for_validation<T>(run: impl FnOnce() -> T) -> (T, PrattAdmissionCensus) {
    struct Restore(Option<PrattAdmissionCensusState>);
    impl Drop for Restore {
        fn drop(&mut self) {
            PHP_PRATT_ADMISSION_CENSUS.with(|state| { state.replace(self.0.take()); });
        }
    }
    let restore = Restore(PHP_PRATT_ADMISSION_CENSUS.with(|state| state.replace(Some(PrattAdmissionCensusState::default()))));
    let result = run();
    let census = PHP_PRATT_ADMISSION_CENSUS.with(|state| state.borrow().as_ref().unwrap().census);
    drop(restore);
    (result, census)
}
#[cfg(feature = "service-verification")]
fn admission_census_builder_for_validation() -> usize {
    PHP_PRATT_ADMISSION_CENSUS.with(|state| {
        let mut state = state.borrow_mut();
        let Some(state) = state.as_mut() else { return 0; };
        state.census.builders += 1;
        state.census.builders
    })
}
#[cfg(feature = "service-verification")]
fn record_pratt_admission_for_validation(builder: usize, offset: usize, mode: bool, hit: bool) {
    if builder == 0 { return; }
    PHP_PRATT_ADMISSION_CENSUS.with(|state| {
        let mut state = state.borrow_mut();
        let Some(state) = state.as_mut() else { return; };
        state.census.requests += 1;
        if state.requests.insert((builder, offset, mode)) {
            state.census.distinct_requests += 1;
        } else {
            state.census.repeated_requests += 1;
        }
        if hit { state.census.hits += 1; } else { state.census.scans += 1; }
    });
}

/// Parse with owned common-AST hooks and certified hybrid Pratt regions.
/// Context-dependent syntax retains its ordered contextual lowering.
pub fn parse(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse_direct(source)
}

/// Explicit PEG reference for compatibility checks and frontend benchmarks.
pub fn parse_with_peg(source: &str) -> Result<vybe_ast::Module, String> {
    walker::parse(source)
}

/// Opt-in preparation receipts collected during common-AST construction.
/// Does not resolve targets, execute includes, or publish source declarations.
pub fn parse_with_source_units(source: &str)
    -> Result<(vybe_ast::Module, Vec<grameo::source_units::Request<String>>), String> {
    walker::parse_direct_with_source_units(source)
}

/// Opt-in streaming preparation boundary: committed syntax requests arrive
/// before contextual AST normalization. Actual loads remain runtime operations.
pub fn parse_with_source_unit_preparation(source: &str,
    on_committed: &mut dyn FnMut(Vec<grameo::source_units::Request<String>>))
    -> Result<vybe_ast::Module,String> {
    walker::parse_direct_with_source_unit_preparation(source,on_committed)
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

fn compiler_optimizations(selected: vybe_runtime::registry::ParseFn) -> vybe_runtime::registry::CompilerOptimizations {
    if std::ptr::fn_addr_eq(selected, parse as vybe_runtime::registry::ParseFn) {
        vybe_runtime::registry::CompilerOptimizations {
            symbol_membership_index: true,
            user_namespace_cache: true,
            completion_worker_lowering: true,
            compact_source_lines: true,
            array_call_tail_identity: true,
            interpolation_call_tail_identity: true,
            ..Default::default()
        }
    } else {
        Default::default()
    }
}

#[cfg(test)]
#[test]
fn compiler_lookup_accelerators_require_the_selected_optimized_frontend() {
    let optimized = compiler_optimizations(parse);
    assert!(optimized.symbol_membership_index && optimized.user_namespace_cache);
    assert!(optimized.completion_worker_lowering);
    assert!(optimized.compact_source_lines);
    assert!(optimized.array_call_tail_identity && optimized.interpolation_call_tail_identity);
    assert_eq!(compiler_optimizations(parse_with_peg), Default::default());
}

/// Register this language with the shared plugin registry (dylib entry point).
pub fn register() {
    // Preparation is selected only for this parser's explicit grammar bindings.
    use grameo::program::Program;
    let preparation_bound = [Rule::kw_include, Rule::kw_include_once, Rule::kw_require, Rule::kw_require_once]
        .into_iter().all(|selector| native_grammar::Parser.source_unit_binding(
            Rule::include_expression as usize, Some(selector as usize)).is_some());
    vybe_compiler::frontend_preparation::register_default(parse,
        vybe_compiler::frontend_preparation::Capabilities {
            prepare: preparation_bound.then_some(prepare_source_units_for_compiler),
            early_hints: false,
            resolve_path: preparation_bound.then_some(vybe_compiler::frontend_preparation::php_source_path),
            request_setup_reuse: false,
            request_code_space: native_grammar::Parser.code_space()==grameo::source_units::CodeSpace::Request,
            helper_definitions: emitter::dispatch::helper_definitions(),
            argument_emitters: emitter::argument_emission::ARGUMENT_DEFINITIONS,
            deferred_class_bodies: Some(normalize_class::normalize_class_deferred),
        });
    if preparation_bound {
        vybe_compiler::statement_codegen::register_prepared_owned(parse,&PHP_PREPARED_STATEMENTS);
    } else { vybe_compiler::statement_codegen::register(parse,produce_grameo_statements); }
    emitter::dispatch::register_typed_results();
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
            compiler_optimizations: Some(compiler_optimizations),
            variable_namespace: Some(&VARIABLE_NAMESPACE),
            function_literal: Some(|chunk, line| {
                let host = chunk.add_import("php:array", "newClosure");
                chunk.emit_call(host, 1, line);
            }),
            rebind_finalizer: Some(|chunk, line| {
                let host = chunk.add_import("php:object", "finalizeUnreachable");
                chunk.emit_call(host, 2, line);
                chunk.emit_op(vybe_runtime::opcode::Op::DROP, line);
            }),
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
    fn init(&self, fw: &mut vybe_runtime::Framework<'_>) {
        register();
        array_read::register(fw);
        globals::register(fw);
        numeric::register(fw);
        string_scan::register(fw);
        random::register(fw);
        sodium::register(fw);
        cookies::register(fw);
        serialization::register(fw);
        spl_filesystem::register(fw);
        dynamic_symbols::register(fw);
        regex::register(fw);
        string_bytes::register(fw);
        html_entities::register(fw);
        datetime::register(fw);
        logging::register(fw);
        environment::register(fw);
        filesystem::register(fw);
        pdo_rows::register(fw);
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

#[cfg(test)]
mod grammar_prefix_tests {
    use super::{PhpParser, Rule};


    fn shape(rule: Rule, source: &str) -> Vec<(Rule, String)> {
        let pair = PhpParser::parse(rule, source)
            .expect("grammar fixture must parse")
            .next()
            .unwrap();
        assert_eq!(pair.as_str().trim(), source);
        pair.into_inner()
            .map(|child| (child.as_rule(), child.as_str().trim().to_owned()))
            .collect()
    }

    #[test]
    fn keyword_dispatch_preserves_boundaries_and_contextual_names() {
        for keyword in [
            "if", "elseif", "else", "while", "do", "for", "foreach", "as",
            "switch", "case", "default", "break", "continue", "return", "goto",
            "declare", "function", "class", "extends", "implements", "interface",
            "trait", "enum", "new", "clone", "echo", "print", "null", "true",
            "false", "instanceof", "include", "include_once", "require",
            "require_once", "throw", "try", "catch", "finally", "static", "public",
            "private", "protected", "abstract", "final", "const", "match", "fn",
            "use", "namespace", "yield", "list", "global", "and", "or", "xor",
            "self", "parent", "isset", "empty", "unset", "insteadof", "endif",
            "endwhile", "endfor", "endforeach", "endswitch",
        ] {
            assert!(PhpParser::parse(Rule::identifier, keyword).is_err(), "{keyword}");
            assert!(PhpParser::parse(Rule::identifier, &keyword.to_ascii_uppercase()).is_err());
            let extended = format!("{keyword}_value");
            assert!(shape(Rule::identifier, &extended).is_empty());
        }
        for name in ["readonly", "get", "set", "from"] {
            assert!(shape(Rule::identifier, name).is_empty());
        }
    }

    #[test]
    fn qualified_names_keep_reserved_namespace_segments() {
        for name in [
            "Service", "Twig\\Environment", "Fn\\Service", "Match\\Node",
            "\\Twig\\Environment", "\\Fn\\Service", "readonly", "function_name",
        ] {
            assert!(shape(Rule::qualified_name, name).is_empty());
        }
        for name in ["fn", "class", "namespace", "MATCH"] {
            assert!(PhpParser::parse(Rule::qualified_name, name).is_err(), "{name}");
        }
    }

    #[test]
    fn variable_dispatch_preserves_outer_pair_types() {
        for (source, rule) in [
            ("$value", Rule::simple_variable),
            ("$$value", Rule::dynamic_variable),
            ("$$$value", Rule::dynamic_variable),
            ("${$name}", Rule::dynamic_variable),
            ("$${$name}", Rule::dynamic_variable),
        ] {
            assert_eq!(shape(Rule::variable, source), vec![(rule, source.to_owned())]);
        }
    }

    #[test]
    fn list_prefix_factoring_preserves_expression_children() {
        for (source, children) in [
            ("", vec![]), ("&", vec![]), ("$value", vec!["$value"]),
            ("&$value", vec!["$value"]),
            ("'key' => &$value", vec!["'key'", "$value"]),
            ("&$key => $value", vec!["&$key", "$value"]),
            ("[1, [2]]", vec!["[1, [2]]"]),
        ] {
            let expected: Vec<_> = children.into_iter()
                .map(|child| (Rule::expression, child.to_owned())).collect();
            assert_eq!(shape(Rule::list_element, source), expected, "{source}");
        }
        PhpParser::parse(Rule::program_pure, "list(, $a, &$b) = $values;").unwrap();
    }

    #[test]
    fn primary_prefix_dispatch_preserves_walker_pairs() {
        for (source, rule) in [
            ("$value", Rule::variable), ("$$name", Rule::variable),
            ("${$name}", Rule::variable), ("1", Rule::literal),
            ("'text'", Rule::literal), ("null", Rule::literal),
            ("true", Rule::literal), ("false", Rule::literal),
            ("Service", Rule::qualified_name), ("Fn\\Service", Rule::qualified_name),
        ] {
            assert_eq!(shape(Rule::primary_expression, source),
                vec![(rule, source.to_owned())], "{source}");
        }
    }

    #[test]
    fn class_member_dispatch_preserves_method_and_property_pairs() {
        let source = "class Sample {
            public $value = 1;
            public function value() { return $this->value; }
            private static $count = 0;
            public static function count() { return self::$count; }
        }";
        let members: Vec<_> = PhpParser::parse(Rule::program_pure, source).unwrap()
            .flatten()
            .filter(|pair| matches!(pair.as_rule(), Rule::method_declaration | Rule::property_declaration))
            .map(|pair| (pair.as_rule(), pair.as_str().trim().to_owned())).collect();
        assert_eq!(members, vec![
            (Rule::property_declaration, "public $value = 1;".to_owned()),
            (Rule::method_declaration, "public function value() { return $this->value; }".to_owned()),
            (Rule::property_declaration, "private static $count = 0;".to_owned()),
            (Rule::method_declaration, "public static function count() { return self::$count; }".to_owned()),
        ]);
    }

    #[test]
    fn whitespace_runs_preserve_tokens_and_string_contents() {
        let source = " \t\r\n$value \t=\r\n' a  b ';\n\t/* block */ \r\n// line\nreturn \t$value;\n";
        let variables: Vec<_> = PhpParser::parse(Rule::program_pure, source).unwrap()
            .flatten()
            .filter(|pair| pair.as_rule() == Rule::simple_variable)
            .map(|pair| pair.as_str().to_owned())
            .collect();
        assert_eq!(variables, vec!["$value", "$value"]);
        assert_eq!(shape(Rule::primary_expression, "' a  b '"),
            vec![(Rule::literal, "' a  b '".to_owned())]);
    }

    #[test]
    fn statement_dispatch_keeps_walker_rule_types() {
        for (source, rule) in [
            ("$value = 1;", Rule::expression_statement),
            ("return $value;", Rule::return_statement),
            ("if ($value) { return; }", Rule::if_statement),
            ("foreach ($values as $value) {}", Rule::foreach_statement),
            ("function value() { return 1; }", Rule::function_declaration),
        ] {
            assert_eq!(shape(Rule::pure_top_level_statement, source),
                vec![(rule, source.to_owned())]);
        }
    }

    #[test]
    fn array_prefix_factoring_preserves_walker_pairs() {
        for (source, children) in [
            ("1", vec!["1"]),
            ("[1, [2, [3]]]", vec!["[1, [2, [3]]]"]),
            ("'key' => $value", vec!["'key'", "$value"]),
            ("'key' => &$value", vec!["'key'", "$value"]),
            ("&$value", vec!["$value"]),
            ("...$values", vec!["$values"]),
            // Preserve the pre-existing grammar shape even for this invalid
            // PHP reference-key form; semantic rejection is not changed here.
            ("&$key => $value", vec!["&$key", "$value"]),
        ] {
            let expected: Vec<_> = children
                .into_iter()
                .map(|child| (Rule::expression, child.to_owned()))
                .collect();
            assert_eq!(shape(Rule::array_element, source), expected, "{source}");
        }
    }

    #[test]
    fn foreach_prefix_factoring_preserves_walker_pairs() {
        for (source, children) in [
            ("$value", vec!["$value"]),
            ("&$value", vec!["$value"]),
            ("$key => $value", vec!["$key", "$value"]),
            ("$key => &$value", vec!["$key", "$value"]),
            ("[$left, $right]", vec!["[$left, $right]"]),
        ] {
            let expected: Vec<_> = children
                .into_iter()
                .map(|child| (Rule::postfix_expression, child.to_owned()))
                .collect();
            assert_eq!(shape(Rule::foreach_target, source), expected, "{source}");
        }
    }

    #[test]
    fn deeply_nested_plain_arrays_parse_without_key_prefix_replay() {
        let mut source = "1".to_owned();
        for _ in 0..16 {
            source = format!("[{source}]");
        }
        let pairs = PhpParser::parse(Rule::array_element, &source).unwrap();
        assert_eq!(
            pairs.flatten().filter(|pair| pair.as_rule() == Rule::array_element).count(),
            17
        );
        let statement = format!("$value = {source};");
        PhpParser::parse(Rule::program_pure, &statement).unwrap();
    }

    #[test]
    #[ignore = "explicit frontend phase benchmark against local WordPress sources"]
    fn profile_wordpress_frontend_phases() {
        use std::sync::{Arc, Mutex};
        use vybe_runtime::debugger::{DebugReport, DebugReportScope};
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/webroot/wordpress/wp-includes");
        for name in ["functions.php", "class-wp-hook.php", "class-wp-theme-json.php",
                     "formatting.php", "script-loader.php"] {
            let source = std::fs::read_to_string(root.join(name)).unwrap();
            // Warm process-wide immutable libraries outside the measured scope.
            drop(super::parse(&source).unwrap());
            let report = Arc::new(Mutex::new(DebugReport::default()));
            let started = std::time::Instant::now();
            let module = {
                let _scope = DebugReportScope::enter(Some(report.clone()));
                super::parse(&source).unwrap()
            };
            let elapsed = started.elapsed();
            eprintln!("frontend_phases file={name} bytes={} total_ms={:.3}",
                source.len(), elapsed.as_secs_f64() * 1000.0);
            for (phase, duration) in &report.lock().unwrap().completed {
                eprintln!("frontend_phase file={name} phase={phase:?} ms={:.3}",
                    duration.as_secs_f64() * 1000.0);
            }
            // Tree teardown is deliberately outside construction timing.
            drop(module);
        }
    }

    #[test]
    #[ignore = "explicit grammar timing benchmark against local WordPress sources"]
    fn profile_wordpress_grammar_parse() {
        use std::time::Instant;
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/webroot/wordpress/wp-includes");
        for name in [
            "functions.php",
            "class-wp-hook.php",
            "class-wp-theme-json.php",
            "formatting.php",
            "script-loader.php",
        ] {
            let raw = std::fs::read_to_string(root.join(name)).unwrap();
            let source = super::normalize_source_for_parser(&raw);
            // Select the applicable entry outside the measured interval. No
            // AST walking, file IO or source normalization is timed as parsing.
            let rule = if PhpParser::parse(Rule::program_pure, &source).is_ok() {
                Rule::program_pure
            } else {
                Rule::program
            };
            let mut samples = Vec::with_capacity(5);
            let mut pair_count = 0;
            for _ in 0..5 {
                let start = Instant::now();
                let pairs = PhpParser::parse(rule, &source).unwrap();
                samples.push(start.elapsed());
                pair_count = pairs.flatten().count();
            }
            samples.sort_unstable();
            eprintln!(
                "grammar_parse file={name} bytes={} pairs={pair_count} entry={rule:?} min_ms={:.3} median_ms={:.3}",
                source.len(), samples[0].as_secs_f64() * 1000.0,
                samples[2].as_secs_f64() * 1000.0,
            );
        }
    }

    #[test]
    #[ignore = "explicit grammar corpus benchmark against local WordPress sources"]
    fn profile_wordpress_grammar_corpus() {
        use std::path::{Path, PathBuf};
        use std::time::{Duration, Instant};

        fn collect_php(directory: &Path, files: &mut Vec<PathBuf>) -> std::io::Result<()> {
            for entry in std::fs::read_dir(directory)? {
                let entry = entry?;
                let kind = entry.file_type()?;
                let path = entry.path();
                if kind.is_dir() {
                    collect_php(&path, files)?;
                } else if kind.is_file() && path.extension().is_some_and(|ext| ext == "php") {
                    files.push(path);
                }
            }
            Ok(())
        }

        let root = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../examples/webroot/wordpress/wp-includes");
        let mut files = Vec::new();
        collect_php(&root, &mut files).unwrap();
        files.sort_unstable();
        // File IO and normalization stay outside parser timing. This measures
        // the corpus, not just the subset loaded by one WordPress request.
        let sources: Vec<_> = files.into_iter().map(|path| {
            let raw = std::fs::read_to_string(&path).unwrap();
            (path, super::normalize_source_for_parser(&raw))
        }).collect();
        let mut total = Duration::ZERO;
        let mut bytes = 0usize;
        let mut pairs_count = 0usize;
        let mut methods = 0usize;
        let mut properties = 0usize;
        let mut rejected = 0usize;
        let mut timings = Vec::with_capacity(sources.len());
        for (path, source) in &sources {
            let rule = if PhpParser::parse(Rule::program_pure, source).is_ok() {
                Rule::program_pure
            } else {
                Rule::program
            };
            let start = Instant::now();
            let parsed = PhpParser::parse(rule, source);
            let elapsed = start.elapsed();
            total += elapsed;
            bytes += source.len();
            let name = path.strip_prefix(&root).unwrap().to_path_buf();
            timings.push((elapsed, name.clone()));
            match parsed {
                Ok(pairs) => {
                    for pair in pairs.flatten() {
                        pairs_count += 1;
                        methods += usize::from(pair.as_rule() == Rule::method_declaration);
                        properties += usize::from(pair.as_rule() == Rule::property_declaration);
                    }
                }
                Err(error) => {
                    rejected += 1;
                    // Never print source text: even a corpus can contain secrets.
                    eprintln!("grammar_corpus_rejected file={} location={:?}",
                        name.display(), error.offset);
                }
            }
        }
        timings.sort_unstable_by(|left, right| right.0.cmp(&left.0));
        eprintln!("grammar_corpus files={} bytes={bytes} pairs={pairs_count} methods={methods} properties={properties} rejected={rejected} parse_ms={:.3}",
            sources.len(), total.as_secs_f64() * 1000.0);
        for (elapsed, name) in timings.iter().take(12) {
            eprintln!("grammar_corpus_hot file={} parse_ms={:.3}",
                name.display(), elapsed.as_secs_f64() * 1000.0);
        }
    }
}

fn prepare_source_units_for_compiler(
    source: &str,
    caller: &std::path::Path,
    sink: &mut vybe_compiler::frontend_preparation::PreparationSink<'_>,
) -> Result<vybe_compiler::frontend_preparation::FrontendPreparation, String> {
    prepare_source_units_common(source,caller,sink,StatementContextMode::None).map(|(prepared,_)| prepared)
}
static PHP_PREPARED_STATEMENTS: vybe_compiler::statement_codegen::PreparedOwnedHooks =
    vybe_compiler::statement_codegen::PreparedOwnedHooks {
        prepare: prepare_source_units_with_owned_context,
        validate: walker::validate_prepared_context,
    };
/// Request-owned declaration discovery, extracted from the original parse.
pub fn prepare_source_units_with_statement_context(source: &str, caller: &std::path::Path,
    sink: &mut vybe_compiler::frontend_preparation::PreparationSink<'_>)
    -> Result<(vybe_compiler::frontend_preparation::FrontendPreparation,
        vybe_compiler::statement_codegen::ProducerContext),String> {
    let (prepared,context) = prepare_source_units_common(source,caller,sink,StatementContextMode::Replay)?;
    Ok((prepared,context.ok_or("PHP prepared statement context missing")?))
}
enum StatementContextMode { None, Replay, Owned }
fn prepare_source_units_with_owned_context(source: &str, caller: &std::path::Path,
    sink: &mut vybe_compiler::frontend_preparation::PreparationSink<'_>)
    -> Result<(vybe_compiler::frontend_preparation::FrontendPreparation,
        vybe_compiler::statement_codegen::ProducerContext),String> {
    let (prepared, context) = prepare_source_units_common(source, caller, sink, StatementContextMode::Owned)?;
    Ok((prepared, context.ok_or("PHP owned statement context missing")?))
}
#[cfg(feature = "service-verification")]
pub fn verify_owned_statement_preparation(source: &str, caller: &std::path::Path)
    -> Result<(vybe_ast::Module, usize, bool), String> {
    let before = walker::facts_capture_count();
    let (prepared, context) = prepare_source_units_with_owned_context(source, caller, &mut |_| {})?;
    walker::validate_prepared_context(source, &context)?;
    let wrong_source_rejected = walker::validate_prepared_context(&format!("{source} "), &context).is_err();
    let wrong_type: vybe_compiler::statement_codegen::ProducerContext = std::sync::Arc::new(());
    let checks = wrong_source_rejected && walker::validate_prepared_context(source, &wrong_type).is_err()
        && walker::uncommitted_owned_context_rejected();
    Ok((prepared.module, walker::facts_capture_count() - before, checks))
}
fn prepare_source_units_common(source: &str, caller: &std::path::Path,
    sink: &mut vybe_compiler::frontend_preparation::PreparationSink<'_>, context_mode: StatementContextMode)
    -> Result<(vybe_compiler::frontend_preparation::FrontendPreparation,
        Option<vybe_compiler::statement_codegen::ProducerContext>),String> {
    use grameo::source_units::{Missing, Phase, PreparationContext, Repetition, Scope};
    let mut committed = |requests: Vec<grameo::source_units::Request<String>>| {
        // The grammar defines load semantics. This adapter only prepares the
        // supported source-execution policy; it neither evaluates operands nor
        // publishes include_once state, declarations, or execution permissions.
        let operands = requests.into_iter().filter(|request| {
            request.loader == "source"
                && request.policy.phase == Phase::Execute
                && request.policy.scope == Scope::Caller
                && matches!(request.policy.missing, Missing::Warning | Missing::Error)
                && matches!(request.policy.repetition, Repetition::Always | Repetition::Once)
                && request.preparation == PreparationContext::Immediate
        }).map(|request| request.target);
        sink(vybe_compiler::frontend_preparation::php_operand_hints(operands, caller));
    };
    let (module,context) = match context_mode {
        StatementContextMode::Replay => {
            let (module,context)=walker::parse_with_statement_context(source,&mut committed)?;
            (module,Some(context))
        }
        StatementContextMode::Owned => {
            let (module,context)=walker::parse_with_owned_statement_context(source,&mut committed)?;
            (module,Some(context))
        }
        StatementContextMode::None => (parse_with_source_unit_preparation(source,&mut committed)?,None),
    };
    Ok((vybe_compiler::frontend_preparation::FrontendPreparation { module, hints: None },context))
}


/// Opt-in completed-capture statement producer; ordinary PEG APIs are unchanged.
pub use walker::{produce_grameo_statements_prepared, produce_grameo_statements_prepared_with_receipt, produce_grameo_statements, produce_grameo_statements_with_receipt, StatementStreamReceipt};
