//! Diagnostic granularity is deliberately separate from recognition parity.
use pest::Parser;
use pest_derive::Parser;
#[derive(Parser)]
#[grammar_inline = "program = { SOI ~ token ~ EOI } token = @{ \"ab\" ~ \"cd\" }"]
struct Reference;
#[test]
fn terminal_failure_position_is_distinct_from_reference_rule_attempt_position() {
    let grammar =
        grameo::compile("program = { SOI ~ token ~ EOI } token = @{ \"ab\" ~ \"cd\" }")
            .unwrap();
    let reference = Reference::parse(Rule::program, "abx").unwrap_err();
    let offset = match reference.location {
        pest::error::InputLocation::Pos(offset) | pest::error::InputLocation::Span((offset, _)) => {
            offset
        }
    };
    let ours = grammar.parse_source("program", "abx").unwrap_err();
    assert_eq!(offset, 0);
    assert_eq!(ours.kind, grameo::ParseErrorKind::Syntax);
    assert_eq!(ours.offset, 2);
    assert_eq!(ours.expected, ["\"cd\""]);
}
