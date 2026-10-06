//! Regression for mandatory trivia between repetition segments/optional items.
use pest::Parser;
use pest_derive::Parser;
use grameo::{
    MatchOptions,
    engine::{parse_program, recognize_program},
    program::Program,
};
#[derive(Parser)]
#[grammar = "tests/fixtures.pest"]
struct Reference;
type Shape = Vec<(String, usize, usize, usize)>;
fn reference_shape(pairs: pest::iterators::Pairs<'_, Rule>) -> Shape {
    fn visit(pair: pest::iterators::Pair<'_, Rule>, depth: usize, shape: &mut Shape) {
        shape.push((
            format!("{:?}", pair.as_rule()),
            pair.as_span().start(),
            pair.as_span().end(),
            depth,
        ));
        for child in pair.into_inner() {
            visit(child, depth + 1, shape);
        }
    }
    let mut shape = Vec::new();
    for pair in pairs {
        visit(pair, 0, &mut shape);
    }
    shape
}
fn native_shape(tree: &grameo::ParseTree<'_, '_>) -> Shape {
    fn visit(pair: grameo::Pair<'_, '_, '_>, depth: usize, shape: &mut Shape) {
        let span = pair.as_span();
        shape.push((pair.rule_name().into(), span.start, span.end, depth));
        for child in pair.into_inner() {
            visit(child, depth + 1, shape);
        }
    }
    let mut shape = Vec::new();
    for pair in tree.pairs() {
        visit(pair, 0, &mut shape);
    }
    shape
}
fn check<P: Program>(program: &P, entry: &str, input: &str, reference: &Result<Shape, ()>) {
    let captured = parse_program(program, entry, input, MatchOptions::default());
    let recognized = recognize_program(program, entry, input, MatchOptions::default());
    match (captured, recognized, reference) {
        (Ok(tree), Ok(recognition), Ok(expected)) => {
            assert_eq!(&native_shape(&tree), expected, "{entry} {input:?}");
            assert_eq!(recognition.consumed, tree.consumed());
        }
        (Err(captured), Err(recognized), Err(())) => {
            assert_eq!(captured.kind, grameo::ParseErrorKind::Syntax);
            assert_eq!(recognized, captured);
        }
        (captured, recognized, reference) => {
            panic!("{entry} {input:?}: {captured:?} {recognized:?} {reference:?}")
        }
    }
}
#[test]
fn repetition_segments_match_reference_acceptance_consumption_and_trivia_captures() {
    let grammar = grameo::compile(include_str!("fixtures.pest")).unwrap();
    let mut inputs = vec![
        String::new(),
        "a //c\nx".into(),
        " a //c\nx".into(),
        "aa //c\nx".into(),
        " ab a".into(),
    ];
    for len in 1..=4 {
        for mut number in 0..4usize.pow(len) {
            let mut input = String::new();
            for _ in 0..len {
                input.push(['a', 'b', ' ', '\t'][number % 4]);
                number /= 4;
            }
            inputs.push(input);
        }
    }
    for rule in [
        Rule::rep_star,
        Rule::rep_plus,
        Rule::rep_exact,
        Rule::rep_minimum,
        Rule::rep_bounded,
        Rule::rep_maximum,
        Rule::rep_two,
        Rule::rep_three,
        Rule::rep_atomic,
        Rule::rep_predicate,
        Rule::rep_rollback,
    ] {
        let entry = format!("{rule:?}");
        for input in &inputs {
            let reference = Reference::parse(rule, input)
                .map(reference_shape)
                .map_err(|_| ());
            check(&grammar, &entry, input, &reference);
            check(
                &vybe_parser_generated_tests::fixtures::Parser,
                &entry,
                input,
                &reference,
            );
        }
    }
    for input in [" ab a", "ab a", "ab ab a", "a ba", "ab //c\na", "ab //c\nx"] {
        let reference = Reference::parse(Rule::rep_stateful, input)
            .map(reference_shape)
            .map_err(|_| ());
        check(&grammar, "rep_stateful", input, &reference);
        check(
            &vybe_parser_generated_tests::fixtures::Parser,
            "rep_stateful",
            input,
            &reference,
        );
    }
    // Independent expected positions keep the regression meaningful even if
    // both native backends regress together.
    assert_eq!(
        grammar.parse_source("rep_plus", "a ").unwrap().consumed(),
        2
    );
    assert_eq!(
        grammar.parse_source("rep_star", "a ").unwrap().consumed(),
        1
    );
    assert_eq!(
        grammar
            .parse_source("rep_maximum", " a")
            .unwrap()
            .consumed(),
        2
    );
    assert_eq!(
        grammar
            .parse_source("rep_maximum", " x")
            .unwrap()
            .consumed(),
        1
    );
}
