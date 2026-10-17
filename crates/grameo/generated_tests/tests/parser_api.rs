use grameo::{builder::Builder, Span};
use vybe_parser_generated_tests::{capture_api, hybrid_api};

#[derive(Default)]
struct Hooks(Vec<(usize, Span)>);
impl Builder for Hooks {
    fn checkpoint(&self) -> usize { self.0.len() }
    fn rollback(&mut self, mark: usize) { self.0.truncate(mark); }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, rule: usize, span: Span, _: &str) { self.0.push((rule, span)); }
}

#[test]
fn selected_generated_apis_compile_and_preserve_typed_captures_and_hooks() {
    let pairs = capture_api::Parser.parse_pairs(capture_api::Rule::root, "hello").unwrap();
    let root = pairs.into_iter().next().unwrap();
    assert_eq!(root.as_rule(), capture_api::Rule::root);
    assert_eq!(root.as_str(), "hello");
    let mut hooks = Hooks::default();
    let tree = hybrid_api::Parser.build_captures(hybrid_api::Rule::root, "hello", &mut hooks).unwrap();
    assert_eq!(tree.consumed(), 5);
    assert_eq!(hooks.0.iter().map(|(rule, _)| *rule).collect::<Vec<_>>(),
        [hybrid_api::Rule::token as usize, hybrid_api::Rule::root as usize]);
    let before = hooks.0.clone();
    assert!(hybrid_api::Parser.build_captures(hybrid_api::Rule::root, "hello!", &mut hooks).is_err());
    assert_eq!(hooks.0, before);
}
