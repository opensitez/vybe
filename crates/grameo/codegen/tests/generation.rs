use grameo_codegen::generate;

#[test]
fn source_units_are_explicit_typed_metadata_and_default_output_is_unchanged() {
    let plain = "load = { keyword ~ target } keyword = @{ \"load\" } target = @{ ASCII_ALPHA+ }";
    let annotated = format!("// @grameo source_unit load target=target selector=keyword loader=source phase=execute repeat=once missing=error scope=caller\n{plain}");
    assert_eq!(generate(plain).unwrap(), generate(&annotated).unwrap());
    let grammar = grameo::compile_with_options(&annotated,grameo::CompileOptions {optimizations:true,..Default::default()}).unwrap();
    let generated = grameo_codegen::emit(&grammar).unwrap();
    assert!(generated.contains("fn source_unit_binding"));
    assert!(generated.contains("scope:Scope::Caller"));
    assert!(generated.contains("repetition:Repetition::Once"));
    assert!(generated.contains("missing:Missing::Error"));
    assert_eq!(generated,grameo_codegen::emit(&grammar).unwrap());
}

#[test]
fn php_source_units_add_only_policy_metadata_to_the_existing_matcher() {
    let source = include_str!("../../../../languages/php/src/grammar.pest");
    let without = source.lines().filter(|line| !line.contains("@grameo source_unit"))
        .collect::<Vec<_>>().join("\n");
    let emit = |source: &str| {
        let grammar = grameo::compile_with_options(source,grameo::CompileOptions {optimizations:true,..Default::default()}).unwrap();
        grameo_codegen::emit_with_api(&grammar,grameo_codegen::ParserApi::HYBRID_CAPTURES).unwrap()
    };
    let mut annotated = emit(source);
    let start = annotated.find("    fn source_unit_binding(").unwrap();
    let end = start + annotated[start..].find("\n    }\n").unwrap() + "\n    }\n".len();
    annotated.replace_range(start..end, "");
    assert_eq!(annotated, emit(&without));
}

#[test]
fn ast_annotations_generate_static_constructors_and_a_capture_free_entry() {
    let grammar = r#"
        // @grameo ast root = Root(value: token)
        root = { token ~ EOI }
        // @grameo ast token = Token(text: source)
        token = @{ ASCII_ALPHA+ }
    "#;
    let generated = grameo_codegen::generate_with_ast_bindings(grammar).unwrap();
    assert_eq!(generated, grameo_codegen::generate_with_ast_bindings(grammar).unwrap());
    assert!(generated.contains("pub enum AstConstructor"));
    assert!(generated.contains("pub static AST_BINDINGS"));
    assert!(generated.contains("pub fn build_bound_ast"));
    assert!(generated.contains("pub fn build_bound_ast_with_context"));
    assert!(!generated.contains("@grameo ast"));
    let capture_only = grameo_codegen::generate_with_api(grammar, grameo_codegen::ParserApi::CAPTURE_ONLY).unwrap();
    assert!(!capture_only.contains("AST_BINDINGS"));
    let invalid = grammar.replace("Root(value:", "Self(value:");
    assert!(generate(&invalid).is_ok());
    assert!(!generate(grammar).unwrap().contains("AST_BINDINGS"));
    assert_eq!(grameo_codegen::generate_with_ast_bindings(&invalid).unwrap_err()[0].code, "G043");
}

#[test]
fn deterministic_generation_and_escaping() {
    let grammar = r#"type = { "\"\\\n" | ^"é" | 'a'..'z' }"#;
    let a = generate(grammar).unwrap();
    assert_eq!(a, generate(grammar).unwrap());
    assert!(a.contains("r#type"));
    assert!(a.contains("insensitive: true"));
    assert!(!a.contains("CompiledGrammar"));
}

#[test]
fn invalid_grammar_and_unrepresentable_rust_names_are_located() {
    assert_eq!(generate("root = { missing }").unwrap_err()[0].code, "G010");
    let errors = generate("self = { \"x\" }").unwrap_err();
    assert_eq!(errors[0].code, "G013");
    assert_eq!(errors[0].span, grameo::Span { start: 0, end: 4 });
}

#[test]
fn experimental_literal_choice_dispatch_is_explicit_and_deterministic() {
    use grameo_codegen::{GenerationOptions, generate, generate_with_options};
    let source = r#"root = { !("and" | "or" | ^"Not") ~ ANY }"#;
    let plain = generate(source).unwrap();
    assert!(!plain.contains("fn literal_choice_failure"));
    assert_eq!(
        plain,
        generate_with_options(source, GenerationOptions::default()).unwrap()
    );
    let options = GenerationOptions {
        literal_choice_failures: true,
    };
    let enabled = generate_with_options(source, options).unwrap();
    assert!(enabled.contains("fn literal_choice_failure"));
    assert!(enabled.contains("literal_matches(source, \"Not\", true)"));
    assert_eq!(enabled, generate_with_options(source, options).unwrap());
    let no_choice = generate_with_options(r#"root = { "x" }"#, options).unwrap();
    assert!(!no_choice.contains("fn literal_choice_failure"));
}

#[test]
fn embeddings_select_entry_points_without_changing_standalone_defaults() {
    use grameo_codegen::{ParserApi, generate_with_api};
    let grammar = r#"root = { word ~ EOI } word = @{ ASCII_ALPHA+ }"#;
    let full = generate(grammar).unwrap();
    for method in ["pub fn recognize(", "pub fn build_ast<", "pub fn build_captures<"] {
        assert!(full.contains(method));
    }
    let captures = generate_with_api(grammar, ParserApi::CAPTURE_ONLY).unwrap();
    assert!(captures.contains("pub fn parse<"));
    assert!(captures.contains("pub fn parse_pairs<"));
    assert!(!captures.contains("pub fn recognize("));
    assert!(!captures.contains("pub fn build_ast<"));
    assert!(!captures.contains("pub fn build_captures<"));
    let hybrid = generate_with_api(grammar, ParserApi::HYBRID_CAPTURES).unwrap();
    assert!(hybrid.contains("pub fn build_captures<"));
    assert!(!hybrid.contains("pub fn build_ast<"));
}

#[test]
fn fragment_tables_are_emitted_only_when_referenced() {
    for source in [
        include_str!("../../../../languages/dart/src/grammar.pest"),
        include_str!("../../../../languages/fortran/src/grammar.pest"),
        include_str!("../../../../languages/kotlin/src/grammar.pest"),
    ] {
        let generated = generate(source).unwrap();
        for line in generated.lines().filter(|line| line.starts_with("static FRAGMENT_RULE_")) {
            let name = line.split_once(':').unwrap().0.trim_start_matches("static ");
            assert!(generated.contains(&format!("Some(&{name})")), "unreferenced table {name}");
        }
    }
}
