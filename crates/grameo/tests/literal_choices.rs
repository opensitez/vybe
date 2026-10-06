use grameo::{
    CompiledGrammar, MatchOptions, Span,
    builder::{BuildError, Builder},
    engine::{build_program, parse_program, recognize_program},
    program::{Instruction, Program, RuleSpec, Scope},
};

fn compiled(source: &str) -> CompiledGrammar {
    let mut grammar = grameo::compile(source).unwrap();
    assert_eq!(grammar.literal_choices().count(), 0);
    grammar.set_literal_choice_failures(true);
    grammar
}

#[derive(Debug)]
struct Scheduled<'a>(&'a CompiledGrammar);
impl Program for Scheduled<'_> {
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
    fn fast_ascii_repetitions(&self) -> bool {
        true
    }
    fn trivia_failure_prefix(&self, rule: usize) -> Option<grameo::lexical::FailurePrefix> {
        self.0.trivia_failure_prefix(rule)
    }
}

#[derive(Debug, PartialEq, Eq)]
enum Event {
    Begin(usize, usize),
    Finish(usize, Span),
}
struct Trace {
    events: Vec<Event>,
    reject: Option<(usize, bool)>,
}
impl Trace {
    fn new(reject: Option<(usize, bool)>) -> Self {
        Self {
            events: vec![Event::Begin(usize::MAX, 17)],
            reject,
        }
    }
}
impl Builder for Trace {
    fn checkpoint(&self) -> usize {
        self.events.len()
    }
    fn rollback(&mut self, mark: usize) {
        self.events.truncate(mark);
    }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, _: usize, _: Span, _: &str) {}
    fn try_begin(&mut self, rule: usize, offset: usize) -> Result<(), BuildError> {
        self.events.push(Event::Begin(rule, offset));
        if self.reject == Some((rule, false)) {
            Err(BuildError {
                span: Span {
                    start: offset,
                    end: offset,
                },
                message: "rejected begin".into(),
            })
        } else {
            Ok(())
        }
    }
    fn try_finish(&mut self, rule: usize, span: Span, _: &str) -> Result<(), BuildError> {
        self.events.push(Event::Finish(rule, span));
        if self.reject == Some((rule, true)) {
            Err(BuildError {
                span,
                message: "rejected finish".into(),
            })
        } else {
            Ok(())
        }
    }
}

fn captures(
    tree: grameo::ParseTree<'_, '_>,
) -> (usize, Vec<(grameo::CaptureRule, Span, usize)>) {
    let mut pairs = tree.pairs().collect::<Vec<_>>();
    pairs.reverse();
    let mut nodes = Vec::new();
    while let Some(pair) = pairs.pop() {
        let children = pair.into_inner().collect::<Vec<_>>();
        nodes.push((pair.as_rule(), pair.as_span(), children.len()));
        pairs.extend(children.into_iter().rev());
    }
    assert_eq!(nodes.len(), tree.capture_count());
    (tree.consumed(), nodes)
}

const GRAMMAR: &str = r##"
WHITESPACE = _{ " " | COMMENT }
COMMENT = _{ "#" ~ (!NEWLINE ~ ANY)* }
root = { SOI ~ (trial | ci | range | stack) ~ EOI }
trial = { (plain ~ "!") | plain }
plain = { "x" }
ci = @{ ^"Ab" }
range = ${ '中'..'文' }
stack = { PUSH("z") ~ ((pop ~ "!") | (peek ~ pop)) }
pop = { POP }
peek = { PEEK }
probe = { &plain ~ plain }
silent = _{ "x" }
atomic = @{ transparent }
transparent = !{ "x" }
any = { ANY }
eoi = { EOI }
empty = { "" }
choice = { (("and" | "or") | (^"Not" | "while")) | ("if" | "é") }
guard = { !choice ~ ANY* }
positive = { &choice ~ choice }
keyword = _{ ("and" | "or" | "not" | "while" | "if") ~ !ASCII_ALPHANUMERIC }
name = @{ !keyword ~ ASCII_ALPHA ~ ASCII_ALPHANUMERIC* }
"##;

#[test]
fn literal_choice_failure_preserves_captures_diagnostics_work_depth_and_builder_trace() {
    let grammar = compiled(GRAMMAR);
    let scheduled = Scheduled(&grammar);
    for (entry, input) in [
        ("guard", "x"),
        ("guard", "a"),
        ("guard", "and"),
        ("guard", "And"),
        ("guard", "é"),
        ("guard", "中"),
        ("guard", ""),
        ("positive", "not"),
        ("positive", "NOT"),
        ("positive", "while"),
        ("positive", "n"),
        ("choice", "bad"),
        ("choice", "and"),
        ("name", "and"),
        ("name", "andrew"),
        ("name", "self"),
        ("name", "x"),
        ("name", "if"),
        ("root", ""),
        ("root", "x"),
        ("root", "x!"),
        ("root", "x?"),
        ("root", "aB"),
        ("root", "中"),
        ("root", "z z"),
        ("root", "z z!"),
        ("root", "x # hi\n"),
        ("probe", "x"),
        ("silent", "x"),
        ("atomic", "x"),
        ("any", "🙂"),
        ("any", ""),
        ("pop", ""),
        ("eoi", ""),
        ("eoi", "x"),
        ("empty", "🙂"),
    ] {
        for max_rule_depth in 0..=7 {
            for max_steps in 0..=150 {
                let options = MatchOptions {
                    max_rule_depth,
                    max_steps,
                };
                assert_eq!(
                    recognize_program(&grammar, entry, input, options),
                    recognize_program(&scheduled, entry, input, options),
                    "{entry}: {input:?}; {options:?}"
                );
                assert_eq!(
                    parse_program(&grammar, entry, input, options).map(captures),
                    parse_program(&scheduled, entry, input, options).map(captures),
                    "captures {entry}: {input:?}; {options:?}"
                );
                let mut fast = Trace::new(None);
                let mut slow = Trace::new(None);
                let a = build_program(&grammar, entry, input, options, &mut fast);
                let b = build_program(&scheduled, entry, input, options, &mut slow);
                assert_eq!(a, b, "builder {entry}: {input:?}; {options:?}");
                assert_eq!(
                    fast.events, slow.events,
                    "events {entry}: {input:?}; {options:?}"
                );
                if a.is_err() {
                    assert_eq!(fast.events, [Event::Begin(usize::MAX, 17)]);
                }
            }
        }
    }
}

#[test]
fn literal_choice_failure_preserves_partially_mutated_hook_failures_and_budget_ordering() {
    let grammar = compiled(GRAMMAR);
    let scheduled = Scheduled(&grammar);
    for (entry, source) in [
        ("choice", "and"),
        ("guard", "x"),
        ("name", "andrew"),
        ("plain", "x"),
        ("ci", "aB"),
        ("range", "中"),
        ("any", "🙂"),
        ("eoi", ""),
    ] {
        let rule = grammar.rule_id(entry).unwrap();
        for finish in [false, true] {
            for max_steps in 0..=150 {
                let options = MatchOptions {
                    max_steps,
                    ..MatchOptions::default()
                };
                let mut fast = Trace::new(Some((rule, finish)));
                let mut slow = Trace::new(Some((rule, finish)));
                assert_eq!(
                    build_program(&grammar, entry, source, options, &mut fast),
                    build_program(&scheduled, entry, source, options, &mut slow)
                );
                assert_eq!(fast.events, [Event::Begin(usize::MAX, 17)]);
                assert_eq!(slow.events, fast.events);
            }
        }
    }
}

#[test]
fn only_maximal_nonempty_literal_choices_are_certified_with_linear_storage() {
    use grameo::program::Program;
    for source in [
        r#"root = { ("x" | "") }"#,
        r#"root = { ("x" | PUSH("y")) }"#,
        r#"root = { ("x" | POP) }"#,
        r#"root = { ("x" | other) } other={"y"}"#,
    ] {
        let grammar = compiled(source);
        assert_eq!(grammar.literal_choices().count(), 0, "{source}");
    }
    let source = format!(
        "root = {{ !({}) ~ ANY }}",
        (0..1000)
            .map(|n| format!("\"word{n}\""))
            .collect::<Vec<_>>()
            .join(" | ")
    );
    let grammar = compiled(&source);
    let choices = grammar.literal_choices().collect::<Vec<_>>();
    assert_eq!(choices.len(), 1);
    let (id, choice) = choices[0];
    assert_eq!(choice.alternatives.len(), 1000);
    assert_eq!(grammar.literal_choice_failure(id, "other"), Some(3000 - 1)); // choice plus grouping
    assert_eq!(grammar.literal_choice_failure(id, "word17"), None);
    let before = recognize_program(&grammar, "root", "other", MatchOptions::default());
    let mut grammar = grammar;
    grammar.set_literal_choice_failures(false);
    assert_eq!(grammar.literal_choices().count(), 0);
    assert_eq!(
        before,
        recognize_program(&grammar, "root", "other", MatchOptions::default())
    );
    grammar.set_literal_choice_failures(true);
    assert_eq!(grammar.literal_choices().count(), 1);
    assert_eq!(
        before,
        recognize_program(&grammar, "root", "other", MatchOptions::default())
    );
}
