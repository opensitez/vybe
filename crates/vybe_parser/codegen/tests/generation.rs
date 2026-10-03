use vybe_parser_codegen::generate;

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
    assert_eq!(errors[0].span, vybe_parser::Span { start: 0, end: 4 });
}

#[test]
fn all_repository_grammars_generate() {
    let languages = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../../languages");
    let mut count = 0;
    for language in std::fs::read_dir(languages).unwrap() {
        let path = language.unwrap().path().join("src/grammar.pest");
        if path.exists() {
            let source = std::fs::read_to_string(&path).unwrap();
            generate(&source).unwrap_or_else(|errors| panic!("{}: {errors:?}", path.display()));
            count += 1;
        }
    }
    assert_eq!(count, 18);
}
