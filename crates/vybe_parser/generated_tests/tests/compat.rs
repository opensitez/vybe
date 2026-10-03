use vybe_parser_generated_tests::{expression, lua};

#[test]
fn typed_pairs_outlive_parser_and_iterators_preserve_children_and_clones() {
    let pairs = expression::Parser
        .parse_pairs(expression::Rule::expression, "12 + 3")
        .unwrap();
    let root = pairs.clone().next().unwrap();
    assert_eq!(root.as_rule(), expression::Rule::expression);
    assert_eq!(root.as_str(), "12 + 3");
    let span = root.as_span();
    assert_eq!((span.start(), span.end()), (0, 6));
    assert_eq!(span.as_str(), root.as_str());
    let copy = root.clone();
    drop(pairs);
    let children: Vec<_> = root
        .into_inner()
        .map(|pair| (pair.as_rule(), pair.as_str()))
        .collect();
    assert_eq!(
        children,
        vec![
            (expression::Rule::number, "12"),
            (expression::Rule::operator, "+"),
            (expression::Rule::number, "3"),
            (expression::Rule::EOI, "")
        ]
    );
    assert_eq!(copy.into_inner().count(), 4);
    let eoi = expression::Parser
        .parse_pairs(expression::Rule::EOI, "")
        .unwrap()
        .next()
        .unwrap();
    assert_eq!(eoi.as_rule(), expression::Rule::EOI);
    assert!(
        expression::Parser
            .parse_pairs(expression::Rule::EOI, "x")
            .is_err()
    );
}

#[test]
fn compatibility_positions_are_one_based_scalars_and_utf8_byte_spans() {
    let mut pairs = lua::Parser
        .parse_pairs(lua::Rule::string, "\"é😀\"")
        .unwrap();
    let span = pairs.next().unwrap().as_span();
    assert_eq!(span.end(), 8);
    assert_eq!(span.start_pos().pos(), 0);
    assert_eq!(span.start_pos().line_col(), (1, 1));
    assert_eq!(span.end_pos().line_col(), (1, 5));
}
