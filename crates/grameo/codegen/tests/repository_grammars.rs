use grameo_codegen::generate;

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

