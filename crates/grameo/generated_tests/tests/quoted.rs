use grameo::{MatchOptions, Span, builder::Builder, engine, program::Program};
use vybe_parser_generated_tests::quoted::{Parser, Rule};

#[derive(Default)]
struct Hooks(Vec<(usize, Span)>);
impl Builder for Hooks {
    fn checkpoint(&self) -> usize { self.0.len() }
    fn rollback(&mut self, mark: usize) { self.0.truncate(mark); }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, rule: usize, span: Span, _: &str) { self.0.push((rule, span)); }
    fn opaque_rule(&self, _: usize) -> bool { true }
}

#[test]
fn generated_quote_scanning_preserves_typed_captures_and_semantic_hooks() {
    let grammar = grameo::compile(include_str!("quoted.grammar")).unwrap();
    assert!(Parser.quoted_fragment(Parser.rule_id("token").unwrap()).is_some());
    for source in ["''", "'hello 雪'", "'a\\'b'", "'a\\\\b'", "'a\\qb'", "'unclosed", "'a'junk"] {
        let mut actual_hooks = Hooks::default();
        let mut expected_hooks = Hooks::default();
        let actual = engine::build_capture_program(&Parser, "root", source, MatchOptions::default(), &mut actual_hooks)
            .map(|tree| tree.into_typed_pairs::<Rule>().flatten().map(|pair| (pair.as_rule().name(), pair.as_span().start(), pair.as_span().end())).collect::<Vec<_>>());
        let expected = engine::build_capture_program(&grammar, "root", source, MatchOptions::default(), &mut expected_hooks)
            .map(|tree| {
                let mut out = Vec::new();
                let mut pending: Vec<_> = tree.pairs().collect();
                while let Some(pair) = pending.pop() {
                    let span = pair.as_span();
                    out.push((pair.rule_name().to_owned(), span.start, span.end));
                    pending.extend(pair.into_inner());
                }
                out
            });
        match (actual, expected) {
            (Ok(a), Ok(mut b)) => {
                // Both views expose the same capture identities and spans.
                let mut a: Vec<_> = a.into_iter().map(|(name, start, end)| (name.to_owned(), start, end)).collect();
                a.sort_unstable(); b.sort_unstable(); assert_eq!(a, b);
            }
            (Err(a), Err(b)) => assert_eq!(a, b),
            _ => panic!("{source}: recognition differs"),
        }
        assert_eq!(actual_hooks.0, expected_hooks.0, "{source}: hooks differ");
    }
}
