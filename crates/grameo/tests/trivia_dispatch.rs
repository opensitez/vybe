use grameo::{
    MatchOptions, Pair, ParseTree,
    engine::{parse_program, recognize_program},
    program::{Instruction, Program, RuleSpec, Scope},
};

#[derive(Debug)]
struct Scheduled<'a>(&'a grameo::CompiledGrammar);
impl Program for Scheduled<'_> {
    fn fast_rule_failures(&self) -> bool {
        false
    }
    fn tail_dispatch(&self) -> bool {
        false
    }
    fn fast_trivia_failures(&self) -> bool {
        false
    }
    fn fast_predicate_failures(&self) -> bool {
        false
    }
    fn trivia_failure_prefix(&self, rule: usize) -> Option<grameo::lexical::FailurePrefix> {
        self.0.trivia_failure_prefix(rule)
    }
    fn fast_ascii_repetitions(&self) -> bool {
        self.0.fast_ascii_repetitions()
    }
    fn rule_id(&self, name: &str) -> Option<usize> {
        self.0.rule_id(name)
    }
    fn rule(&self, rule: usize) -> RuleSpec<'_> {
        self.0.rule(rule)
    }
    fn instruction(&self, expression: usize) -> Instruction<'_> {
        self.0.instruction(expression)
    }
    fn scope(&self, scope: usize) -> Scope {
        self.0.scope(scope)
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
fn fused_trivia_preserves_captures_full_errors_and_every_short_budget_boundary() {
    for trivia in [
        "",
        r##"COMMENT = _{ "#" ~ (!NEWLINE ~ ANY)* }"##,
        r##"WHITESPACE = _{ " " } COMMENT = _{ "#" ~ (!NEWLINE ~ ANY)* }"##,
        r##"WHITESPACE = _{ " " | COMMENT } COMMENT = _{ "#" ~ (!NEWLINE ~ ANY)* }"##,
        r##"WHITESPACE = { " " | COMMENT } COMMENT = { "#" ~ (!NEWLINE ~ ANY)* }"##,
        r##"WHITESPACE = _{ " " | COMMENT } COMMENT = _{ prefix ~ POP } prefix = @{ "!" }"##,
        r##"WHITESPACE = _{ " " | COMMENT } COMMENT = _{ "!" ~ PUSH("x") ~ "z" }"##,
        r##"WHITESPACE = _{ " " | COMMENT } COMMENT = _{ ^"é" ~ ANY | ('中'..'文') ~ ANY }"##,
        r##"WHITESPACE = _{ " " } COMMENT = _{ "" }"##,
    ] {
        let source = format!(
            "{trivia}\nroot = {{ SOI ~ (word ~ \"c\" | word ~ \"b\") ~ EOI }} word = {{ \"a\" }}"
        );
        // Lower-level resolve retains the deliberately nullable trivia fixture
        // so its runtime non-progress error is compared as well.
        let grammar = grameo::resolve(grameo::parse(&source).unwrap()).unwrap();
        let scheduled = Scheduled(&grammar);
        for input in [
            "", "ab", "ac", "a b", " a b ", "a#hi\nb", "a!b", "a!xz b", "aéb", "aéxb", "a中xb",
            "ab!",
        ] {
            for depth in 0..=5 {
                for budget in 0..=240 {
                    let options = MatchOptions {
                        max_steps: budget,
                        max_rule_depth: depth,
                    };
                    assert_eq!(
                        parse_program(&grammar, "root", input, options).map(shape),
                        parse_program(&scheduled, "root", input, options).map(shape),
                        "{source}; {input:?}; {options:?}",
                    );
                }
            }
            assert_eq!(
                recognize_program(&grammar, "root", input, MatchOptions::default()),
                recognize_program(&scheduled, "root", input, MatchOptions::default()),
            );
        }
    }
}

#[test]
fn predicate_prefix_dispatch_preserves_limits_unicode_stack_and_atomic_modes() {
    for keyword in [
        r#"keyword = _{ ("if" | "or" | "for") ~ !(ASCII_ALPHANUMERIC | "_") }"#,
        r#"keyword = !{ (nested | ^"if") ~ !(ASCII_ALPHANUMERIC | "_") } nested = { "or" }"#,
        r#"keyword = ${ ("if" | '中'..'文') ~ !ASCII_ALPHA }"#,
        r#"keyword = _{ "!" ~ POP }"#,
        r#"keyword = _{ PUSH("!") ~ "z" }"#,
        r#"keyword = _{ "if"? }"#,
    ] {
        let source = format!(
            r#"{keyword} WHITESPACE = _{{ " " }} root = {{ SOI ~ word ~ EOI }} word = @{{ !keyword ~ ANY+ }}"#
        );
        let grammar = grameo::compile(&source).unwrap();
        let scheduled = Scheduled(&grammar);
        for input in [
            "", "x", "x0", "if", "ifx", "IF", "or", "or_", "for", "é", "中x", "!", "!z", "!a",
        ] {
            for depth in 0..=7 {
                for budget in 0..=240 {
                    let options = MatchOptions {
                        max_steps: budget,
                        max_rule_depth: depth,
                    };
                    assert_eq!(
                        parse_program(&grammar, "root", input, options).map(shape),
                        parse_program(&scheduled, "root", input, options).map(shape),
                        "{source}; {input:?}; {options:?}",
                    );
                }
            }
        }
    }
}

#[test]
fn ordinary_rule_and_terminal_choice_dispatch_preserves_diagnostics_and_work() {
    for source in [
        r#"root = { SOI ~ (first ~ "!" | second ~ "?") ~ EOI }
            first = { nested } nested = ${ (("ab" | "a") | ^"É" | '中'..'文') }
            second = !{ PUSH(("x" | "y")) ~ POP }"#,
        r#"root = { SOI ~ token{1,3} ~ EOI }
            token = @{ (ASCII_ALPHA | "_" | '中'..'文') ~ (ASCII_ALPHANUMERIC | "_")* }"#,
        r#"root = { SOI ~ (item | alternative) ~ EOI }
            item = { "abc" ~ "def" } alternative = { "a" ~ ("x" | "y" | "x") }"#,
        r#"root = { SOI ~ (item | fallback) ~ EOI }
            item = { ("if" | "in" | "else" | "while" | "do" | "end" | "return" | "then") ~ !(ASCII_ALPHANUMERIC | "_") }
            fallback = @{ (ASCII_ALPHA | "_")+ }
            WHITESPACE = _{ " " | "\t" }"#,
    ] {
        let grammar = grameo::compile(source).unwrap();
        let scheduled = Scheduled(&grammar);
        for input in [
            "", "a", "ab", "ab!", "a?", "abcx", "abcdef", "É!", "é!", "中!", "x", "if", "iffy",
            "  else",
        ] {
            for depth in 0..=5 {
                for budget in 0..=240 {
                    let options = MatchOptions {
                        max_steps: budget,
                        max_rule_depth: depth,
                    };
                    assert_eq!(
                        parse_program(&grammar, "root", input, options).map(shape),
                        parse_program(&scheduled, "root", input, options).map(shape),
                        "capture: {source}; {input:?}; {options:?}",
                    );
                    assert_eq!(
                        recognize_program(&grammar, "root", input, options),
                        recognize_program(&scheduled, "root", input, options),
                        "recognition: {source}; {input:?}; {options:?}",
                    );
                }
            }
        }
    }
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
