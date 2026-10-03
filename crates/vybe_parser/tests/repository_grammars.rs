use std::{fs, path::Path};
use vybe_parser::{ExprKind, compile};

#[test]
fn all_existing_language_grammars_compile_without_language_crate_dependencies() {
    let languages = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../languages");
    let mut paths: Vec<_> = fs::read_dir(languages)
        .unwrap()
        .map(|entry| entry.unwrap().path().join("src/grammar.pest"))
        .filter(|path| path.is_file())
        .collect();
    paths.sort();
    assert_eq!(
        paths.len(),
        18,
        "update the corpus inventory deliberately if languages change"
    );
    let mut failures = Vec::new();
    let mut rule_count = 0;
    for path in paths {
        let source = fs::read_to_string(&path).unwrap();
        match compile(&source) {
            Ok(grammar) => {
                assert!(
                    grammar.syntax().rules.len() > 100,
                    "{} unexpectedly truncated",
                    path.display()
                );
                let language = path
                    .parent()
                    .unwrap()
                    .parent()
                    .unwrap()
                    .file_name()
                    .unwrap()
                    .to_str()
                    .unwrap();
                let entry = if language == "lua" {
                    "chunk"
                } else {
                    "program"
                };
                assert!(
                    grammar.rule_id(entry).is_some(),
                    "{} lost entry rule {entry}",
                    path.display()
                );
                rule_count += grammar.syntax().rules.len();
                for (id, expr) in grammar.syntax().expressions.iter().enumerate() {
                    assert!(source.get(expr.span.start..expr.span.end).is_some());
                    if matches!(expr.kind, ExprKind::Reference(_)) {
                        assert!(grammar.reference(id).is_some());
                    }
                }
            }
            Err(errors) => failures.extend(
                errors
                    .into_iter()
                    .map(|error| error.render(&path.display().to_string(), &source)),
            ),
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
    assert!(rule_count > 4000);
}
