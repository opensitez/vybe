use grameo::{CaptureRule, program::Program};

#[test]
fn generated_capture_ids_preserve_rule_identity_and_reject_invalid_ids() {
    macro_rules! check {
        ($module:ident) => {{
            use vybe_parser_generated_tests::$module::{Parser, Rule};
            let mut count = 0;
            while let Some(rule) = Rule::from_capture(CaptureRule::User(count)) {
                assert_eq!(Parser.rule_id(rule.name()), Some(count));
                count += 1;
            }
            assert!(count > 0);
            assert_eq!(Rule::from_capture(CaptureRule::User(count)), None);
            assert_eq!(Rule::from_capture(CaptureRule::User(usize::MAX)), None);
            assert_eq!(Rule::from_capture(CaptureRule::EndOfInput), Some(Rule::EOI));
        }};
    }
    check!(fixtures);
    check!(ast);
    check!(lua);
}
