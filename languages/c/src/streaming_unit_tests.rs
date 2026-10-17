use grameo::{MatchOptions, ParseTree, Span, builder::Builder, compat::Pairs};

use crate::{Rule, native_grammar};

#[derive(Default)]
struct UnitBuilder {
    commits: Vec<usize>,
}

impl Builder for UnitBuilder {
    fn observes_rule(&self, _: usize) -> bool { false }
    fn opaque_rule(&self, _: usize) -> bool { true }
    fn checkpoint(&self) -> usize { self.commits.len() }
    fn rollback(&mut self, mark: usize) { self.commits.truncate(mark); }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: Span, _: &str) {}
    fn committed(&mut self, _: &str, end: usize) { self.commits.push(end); }
}

fn signature(tree: ParseTree<'_, '_>) -> Vec<(Rule, usize, usize)> {
    let pairs: Pairs<'_, Rule> = tree.into_typed_pairs();
    pairs.flatten().filter(|pair| !matches!(pair.as_rule(),
        Rule::program | Rule::frontend_unit | Rule::frontend_end
        | Rule::frontend_trivia | Rule::EOI))
        .map(|pair| (pair.as_rule(), pair.as_span().start(), pair.as_span().end()))
        .collect()
}

#[test]
fn committed_units_preserve_whole_program_captures_and_absolute_spans() {
    for source in [
        "",
        " /* only trivia */ \n",
        "; ; int x; ;",
        "/* \u{3bb} */ typedef int T; // gap\nT f(T x) { return x + 1; }\n",
        "struct S { int x; }; int g; int f(void) { return g; } int g = 3;",
        "#define N 3\nconst int values[N] = {1,2,3};\n_Static_assert(1, \"ok\");",
        "int (*callback)(int); int f(int x) { if (x) { return callback(x); } return 0; }",
    ] {
        assert_unit_capture_parity(source);
    }
}

fn assert_unit_capture_parity(source: &str) {
    let parser = native_grammar::Parser;
    let reference = signature(grameo::engine::parse_program(
        &parser, "program", source, MatchOptions::default()).unwrap());
    let mut builder = UnitBuilder::default();
    let mut offset = 0;
    let mut actual = Vec::new();
    loop {
        if let Ok(end) = grameo::engine::parse_program_at(
            &parser, "frontend_end", source, offset, MatchOptions::default(),
        ) {
            assert_eq!(end.consumed(), source.len());
            break;
        }
        let unit = grameo::engine::build_capture_program_at(
            &parser, "frontend_unit", source, offset, MatchOptions::default(), &mut builder,
        ).unwrap();
        let end = unit.consumed();
        assert!(end > offset, "unit must make progress: {source:?}");
        actual.extend(signature(unit));
        assert_eq!(builder.commits.last(), Some(&end));
        offset = end;
    }
    if actual != reference {
        let mismatch = actual.iter().zip(&reference).position(|(left, right)| left != right);
        panic!("unit capture mismatch: source_bytes={} first_difference={mismatch:?} actual={} expected={}",
            source.len(), actual.len(), reference.len());
    }
    assert!(builder.commits.windows(2).all(|ends| ends[0] < ends[1]));
    eprintln!("committed-unit parity: source_bytes={} captures={} units={}",
        source.len(), actual.len(), builder.commits.len());
}

#[test]
#[ignore = "whole saved translation-unit capture/span parity"]
fn committed_units_preserve_complete_saved_translation_unit() {
    let path = std::env::var("VYBE_C_FRONTEND_PREPROCESSED")
        .expect("provide saved preprocessed C source");
    assert_unit_capture_parity(&std::fs::read_to_string(path).unwrap());
}

#[test]
fn failed_later_unit_preserves_prior_commits_and_rejects_invalid_boundaries() {
    let parser = native_grammar::Parser;
    let source = "int first; int broken(void) { return 1 + ; }";
    let mut builder = UnitBuilder::default();
    let first = grameo::engine::build_capture_program_at(
        &parser, "frontend_unit", source, 0, MatchOptions::default(), &mut builder,
    ).unwrap();
    let end = first.consumed();
    let previous = builder.commits.clone();
    assert!(grameo::engine::build_capture_program_at(
        &parser, "frontend_unit", source, end, MatchOptions::default(), &mut builder,
    ).is_err());
    assert_eq!(builder.commits, previous);
    assert!(grameo::engine::build_capture_program_at(
        &parser, "frontend_unit", source, source.len() + 1, MatchOptions::default(), &mut builder,
    ).is_err());
    assert_eq!(builder.commits, previous);
    let utf8 = "/* \u{3bb} */ int x;";
    assert!(grameo::engine::build_capture_program_at(
        &parser, "frontend_unit", utf8, 4, MatchOptions::default(), &mut builder,
    ).is_err());
    assert_eq!(builder.commits, previous);
}
