use grameo::{
    CompiledGrammar, MatchOptions, Pair, ParseTree,
    engine::{parse_program, recognize_program},
    program::{Instruction, Program, RuleSpec, Scope},
};
#[derive(Debug)]
struct Ordinary<'a>(&'a CompiledGrammar);
impl Program for Ordinary<'_> {
    fn fast_ascii_repetitions(&self) -> bool {
        false
    }
    fn atomic_expression(&self, id: usize) -> bool {
        self.0.atomic_expression(id)
    }
    fn leading_bytes(&self, id: usize) -> Option<grameo::lexical::LeadingBytes> {
        let _ = id; None
    }
    fn rule_id(&self, name: &str) -> Option<usize> {
        self.0.rule_id(name)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        self.0.rule(id)
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        self.0.instruction(id)
    }
    fn scope(&self, id: usize) -> Scope {
        self.0.scope(id)
    }
}
fn shape(tree: ParseTree<'_, '_>) -> (usize, Vec<(String, usize, usize, usize)>) {
    fn visit(pair: Pair<'_, '_, '_>, depth: usize, nodes: &mut Vec<(String, usize, usize, usize)>) {
        let span = pair.as_span();
        nodes.push((pair.rule_name().to_owned(), span.start, span.end, depth));
        for child in pair.into_inner() {
            visit(child, depth + 1, nodes);
        }
    }
    let mut nodes = Vec::new();
    for pair in tree.pairs() {
        visit(pair, 0, &mut nodes);
    }
    (tree.consumed(), nodes)
}
#[test]
fn silent_atomic_rule_runs_match_unbatched_tasks_at_every_resource_boundary() {
    for token in [
        r#"token = @{ part* } part = _{ ASCII_ALPHANUMERIC | "_" }"#,
        r#"token = @{ part+ } part = _{ ASCII_ALPHA | "_" }"#,
        r#"token = @{ part{2,4} } part = _{ ASCII_DIGIT }"#,
        r#"token = @{ part* ~ "!" } part = _{ !"x" ~ ASCII_ALPHA }"#,
        r#"token = @{ part* } part = _{ XID_CONTINUE }"#,
        r#"token = @{ part* } part = _{ nested } nested = _{ ASCII_DIGIT }"#,
        r#"token = @{ part* } part = !{ ASCII_DIGIT }"#,
        r#"token = @{ part* } part = ${ ASCII_DIGIT }"#,
    ] {
        let grammar =
            grameo::compile(&format!("root = {{ SOI ~ token ~ EOI }} {token}")).unwrap();
        let ordinary = Ordinary(&grammar);
        let long = "7".repeat(64);
        for input in [
            "", "1", "123", "12345", "123x", "123!", "!123!", "!123x", "é", "a_b12", "é_7", &long,
        ] {
            for depth in 0..=4 {
                for budget in 0..=280 {
                    let options = MatchOptions {
                        max_steps: budget,
                        max_rule_depth: depth,
                    };
                    assert_eq!(
                        parse_program(&grammar, "root", input, options).map(shape),
                        parse_program(&ordinary, "root", input, options).map(shape),
                        "{token} {input:?} {options:?}"
                    );
                    assert_eq!(
                        recognize_program(&grammar, "root", input, options),
                        recognize_program(&ordinary, "root", input, options),
                        "{token} {input:?} {options:?}"
                    );
                }
                let options = MatchOptions { max_steps: usize::MAX, max_rule_depth: depth };
                assert_eq!(parse_program(&grammar, "root", input, options).map(shape),
                    parse_program(&ordinary, "root", input, options).map(shape));
                assert_eq!(recognize_program(&grammar, "root", input, options),
                    recognize_program(&ordinary, "root", input, options));
            }
        }
    }
}
