use grameo::{
    MatchOptions, Pair, ParseTree,
    engine::parse_program,
    program::{Instruction, Program, RuleSpec, Scope},
};

#[derive(Debug)]
struct Scheduled<'a>(&'a grameo::CompiledGrammar);
impl Program for Scheduled<'_> {
    fn atomic_tail_dispatch(&self) -> bool {
        false
    }
    fn atomic_expression(&self, id: usize) -> bool {
        self.0.atomic_expression(id)
    }
    fn leading_bytes(&self, id: usize) -> Option<grameo::lexical::LeadingBytes> {
        self.0.leading_bytes(id)
    }
    fn terminal_choice(&self, id: usize) -> Option<grameo::lexical::TerminalChoice<'_>> {
        self.0.terminal_choice(id)
    }
    fn failure_terminals(&self, id: usize) -> &[usize] {
        self.0.failure_terminals(id)
    }
    fn literal_choice_failure(&self, id: usize, input: &str) -> Option<usize> {
        self.0.literal_choice_failure(id, input)
    }
    fn trivia_failure_prefix(&self, id: usize) -> Option<grameo::lexical::FailurePrefix> {
        self.0.trivia_failure_prefix(id)
    }
    fn fast_ascii_repetitions(&self) -> bool {
        self.0.fast_ascii_repetitions()
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

fn shape(
    tree: ParseTree<'_, '_>,
) -> (
    usize,
    Vec<(grameo::CaptureRule, grameo::Span, usize)>,
) {
    fn visit(
        pair: Pair<'_, '_, '_>,
        depth: usize,
        out: &mut Vec<(grameo::CaptureRule, grameo::Span, usize)>,
    ) {
        out.push((pair.as_rule(), pair.as_span(), depth));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut out = Vec::new();
    for pair in tree.pairs() {
        visit(pair, 0, &mut out);
    }
    assert_eq!(out.len(), tree.capture_count());
    (tree.consumed(), out)
}
#[test]
fn atomic_token_runs_preserve_trees_diagnostics_and_resource_boundaries() {
    for token in [
        r#"token = @{ !keyword ~ (ASCII_ALPHA | "_") ~ (ASCII_ALPHANUMERIC | "_")* } keyword = _{ "if" | "then" }"#,
        r#"token = @{ ASCII_DIGIT+ ~ ("." ~ ASCII_DIGIT+)? }"#,
        r#"token = @{ "'" ~ (!"'" ~ ANY)* ~ "'" }"#,
        r#"token = @{ "'" ~ (!("'" | "\\") ~ ANY)* ~ "'" }"#,
        r#"token = @{ "'" ~ (!"'" ~ !"\\" ~ ANY)* ~ "'" }"#,
        r#"token = @{ "'" ~ (!"''" ~ ANY)* ~ "''" }"#,
        r#"token = @{ part+ } part = { XID_CONTINUE | ('中'..'文') }"#,
        r#"token = @{ (ASCII_ALPHA | "_"){1,3} }"#,
        r#"token = @{ ("" | ASCII_ALPHA){0,2} }"#,
        r#"token = @{ child } child = !{ ASCII_ALPHA+ }"#,
        r#"token = @{ child } child = ${ ASCII_ALPHA+ }"#,
        r#"token = @{ PUSH(ASCII_ALPHA+) ~ POP }"#,
        r#"token = { NEWLINE }"#,
    ] {
        let source = format!("root = {{ SOI ~ token ~ EOI }} {token}");
        let grammar = grameo::resolve(grameo::parse(&source).unwrap()).unwrap();
        let scheduled = Scheduled(&grammar);
        for input in [
            "", "a", "abc123", "abc_123", "if", "123.45", "'é😀'", "'abc'", "'abc''", "'a\\b'",
            "中é", "aa", "\n", "\r", "!",
        ] {
            for depth in [0, 1, 2, 3, 6] {
                for budget in 0..=160 {
                    let options = MatchOptions {
                        max_steps: budget,
                        max_rule_depth: depth,
                    };
                    assert_eq!(
                        parse_program(&grammar, "root", input, options).map(shape),
                        parse_program(&scheduled, "root", input, options).map(shape),
                        "{source}; {input:?}; {options:?}"
                    );
                }
            }
        }
    }
    let grammar = grameo::compile(
        r#"root = { token ~ EOI } token = @{ ASCII_ALPHA ~ (ASCII_ALPHANUMERIC | "_")* }"#,
    )
    .unwrap();
    let scheduled = Scheduled(&grammar);
    let input = format!("{}é", "a".repeat(8192));
    for budget in [1, 8, 100, 8192, 32768, 65536, 131072, usize::MAX] {
        let options = MatchOptions {
            max_steps: budget,
            max_rule_depth: 16,
        };
        assert_eq!(
            parse_program(&grammar, "root", &input, options).map(shape),
            parse_program(&scheduled, "root", &input, options).map(shape),
            "long token; {options:?}"
        );
    }
}
