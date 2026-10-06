#[test]
fn repeated_alternatives_report_each_expectation_once_in_first_seen_order() {
    let choices = (0..1200)
        .map(|i| format!("\"{}\"", ["a", "b", "c"][i % 3]))
        .collect::<Vec<_>>()
        .join(" | ");
    let grammar = grameo::compile(&format!("root = {{ ({choices}) ~ EOI }}")).unwrap();
    let error = grammar.parse_source("root", "").unwrap_err();
    assert_eq!(error.kind, grameo::ParseErrorKind::Syntax);
    assert_eq!(error.offset, 0);
    assert_eq!(error.expected, ["\"a\"", "\"b\"", "\"c\""]);

    let choices = (0..640)
        .map(|i| format!("\"word{i}\""))
        .collect::<Vec<_>>()
        .join(" | ");
    let grammar = grameo::compile(&format!("root = {{ {choices} }}")).unwrap();
    let error = grammar.parse_source("root", "").unwrap_err();
    let expected = (0..640).map(|i| format!("\"word{i}\"")).collect::<Vec<_>>();
    assert_eq!(error.expected, expected);
}

#[test]
fn farther_failure_discards_prior_expectations_after_duplicate_compactions() {
    let choices = (0..1200)
        .map(|i| format!("\"{}\"", ["a", "b", "c"][i % 3]))
        .collect::<Vec<_>>()
        .join(" | ");
    let grammar = grameo::compile(&format!("root = {{ ({choices})* ~ \"!\" ~ EOI }}"))
        .unwrap();
    let error = grammar.parse_source("root", "b").unwrap_err();
    assert_eq!(error.offset, 1);
    assert_eq!(error.expected, ["\"a\"", "\"b\"", "\"c\"", "\"!\""]);
}
