use grameo_codegen::generate;

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
