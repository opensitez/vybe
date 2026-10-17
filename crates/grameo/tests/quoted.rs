use std::cell::Cell;
use grameo::{CompiledGrammar, MatchOptions, engine, program::{Instruction, Program, RuleSpec, Scope}, quoted};

#[derive(Debug)]
struct Certified {
    grammar: CompiledGrammar,
    fragments: Vec<Option<quoted::Fragment>>,
    scans: Cell<usize>,
}
impl Certified {
    fn new(source: &str) -> Self {
        let grammar = grameo::compile(source).unwrap();
        let fragments = (0..grammar.syntax().rules.len()).map(|id| quoted::certify(&grammar, id)).collect();
        Self { grammar, fragments, scans: Cell::new(0) }
    }
}
impl Program for Certified {
    fn match_observers_are_noop(&self) -> bool { true }
    fn rule_id(&self, name: &str) -> Option<usize> { self.grammar.rule_id(name) }
    fn rule(&self, id: usize) -> RuleSpec<'_> { self.grammar.rule(id) }
    fn instruction(&self, id: usize) -> Instruction<'_> { self.grammar.instruction(id) }
    fn scope(&self, id: usize) -> Scope { self.grammar.scope(id) }
    fn atomic_expression(&self, id: usize) -> bool { self.grammar.atomic_expression(id) }
    fn delimited_fragment(&self, id: usize) -> Option<quoted::DelimitedFragment<'_>> {
        quoted::certify_delimited(&self.grammar, id)
    }
    fn quoted_fragment(&self, id: usize) -> Option<quoted::Fragment> {
        let result = self.fragments[id];
        if result.is_some() { self.scans.set(self.scans.get() + 1); }
        result
    }
}
const GRAMMAR: &str = r#"root = { SOI ~ token ~ EOI }
token = @{ "'" ~ ("\\'" | "\\\\" | (!"'" ~ ANY))* ~ "'" }"#;

fn captures(tree: &grameo::ParseTree<'_, '_>) -> Vec<(String, grameo::Span)> {
    let mut out = Vec::new();
    let mut pending: Vec<_> = tree.pairs().collect();
    while let Some(pair) = pending.pop() {
        out.push((pair.rule_name().to_owned(), pair.as_span()));
        pending.extend(pair.into_inner());
    }
    out
}

#[test]
fn quoted_tokens_preserve_captures_errors_unicode_and_budget_boundaries() {
    let program = Certified::new(GRAMMAR);
    assert!(program.fragments[program.rule_id("token").unwrap()].is_some());
    for source in ["''", "'hello'", "'雪 café'", "'a\\'b'", "'a\\\\b'", "'a\\qb'", "'a\nline'", "'unterminated", "'trailing\\", "'a'junk"] {
        let oracle = engine::parse_program(&program.grammar, "root", source, MatchOptions::default());
        let actual = engine::parse_program(&program, "root", source, MatchOptions::default());
        match (actual, oracle) {
            (Ok(a), Ok(b)) => { assert_eq!(a.consumed(), b.consumed()); assert_eq!(captures(&a), captures(&b)); }
            (Err(a), Err(b)) => assert_eq!(a, b, "{source}"),
            _ => panic!("recognition differs for {source}"),
        }
        for budget in 0..200 {
            let options = MatchOptions { max_steps: budget, ..MatchOptions::default() };
            let before = program.scans.get();
            let a = engine::parse_program(&program, "root", source, options).map(|tree| captures(&tree));
            let b = engine::parse_program(&program.grammar, "root", source, options).map(|tree| captures(&tree));
            assert_eq!(a, b, "{source} at budget {budget}");
            assert_eq!(program.scans.get(), before, "explicit budgets must retain exact accounting");
        }
    }
    assert!(program.scans.get() > 0);
    let large = format!("'{}'", "ordinary text 雪 ".repeat(20_000));
    let a = engine::parse_program(&program, "root", &large, MatchOptions::default()).unwrap();
    let b = engine::parse_program(&program.grammar, "root", &large, MatchOptions::default()).unwrap();
    assert_eq!(captures(&a), captures(&b));
}

#[test]
fn newline_restricted_quotes_preserve_escape_order_diagnostics_and_captures() {
    for guard in [r#"NEWLINE = { "\n" | "\r\n" }"#, r#"NEWLINE = { "\r\n" | "\r" | "\n" }"#] {
        let program = Certified::new(&format!(r#"root = {{ token ~ EOI }} token = @{{ "'" ~ (!"'" ~ !NEWLINE ~ ANY)* ~ "'" }} {guard}"#));
        assert!(program.fragments[program.rule_id("token").unwrap()].is_some());
        for source in ["'a\rb'", "'a\nb'", "'a\r\nb'", "'a\r'", "'a\r", "'雪\r文'"] {
            let actual = engine::parse_program(&program, "root", source, MatchOptions::default()).map(|tree| captures(&tree));
            let expected = engine::parse_program(&program.grammar, "root", source, MatchOptions::default()).map(|tree| captures(&tree));
            assert_eq!(actual, expected, "{guard}, {source:?}");
        }
    }
    for escape in [r#""\\'" | "\\\\""#, r#""\\" ~ ANY"#] {
        let program = Certified::new(&format!(r#"root = {{ &token ~ token ~ EOI }} token = @{{ "'" ~ ({escape} | (!"'" ~ !NEWLINE ~ ANY))* ~ "'" }}"#));
        let fragment = program.fragments[program.rule_id("token").unwrap()].unwrap();
        assert!(fragment.forbidden_newline);
        for source in ["''", "'雪é'", "'a\\'b'", "'a\\qb'", "'a\nb'", "'a\rb'", "'a\r\nb'", "'a\\\nb'", "'a\\\rb'", "'a\\\r\nb'", "'missing", "'slash\\", "'ok'junk"] {
            for budget in [usize::MAX, 0, 10, 50, 200] {
                let options = MatchOptions { max_steps: budget, ..MatchOptions::default() };
                let actual = engine::parse_program(&program, "root", source, options).map(|tree| captures(&tree));
                let expected = engine::parse_program(&program.grammar, "root", source, options).map(|tree| captures(&tree));
                assert_eq!(actual, expected, "{escape}, {source:?}, budget {budget}");
            }
        }
        let source = format!("'{}'", "long ordinary 雪 text ".repeat(10_000));
        assert_eq!(quoted::scan(fragment, &source, 0), Some(source.len()));
    }
    for predicate in ["NEWLINE", "ANY"] {
        let grammar = grameo::compile(&format!(r#"token = @{{ "'" ~ (!"'" ~ {predicate} ~ ANY)* ~ "'" }}"#)).unwrap();
        assert!(quoted::certify(&grammar, grammar.rule_id("token").unwrap()).is_none());
    }
}

#[test]
fn quote_certificates_reject_effects_calls_and_shadowing_alternatives() {
    for body in [r#"(ANY | "\\'")*"#, r#"(!"'" ~ ANY | "\\'")*"#, r#"PUSH((!"'" ~ ANY)*)"#, "helper*", r#"(!"x" ~ ANY)*"#] {
        let grammar = grameo::compile(&format!("token = @{{ \"'\" ~ {body} ~ \"'\" }} helper = {{ ANY }}")).unwrap();
        assert!(quoted::certify(&grammar, grammar.rule_id("token").unwrap()).is_none(), "{body}");
    }
}

#[test]
fn actual_php_quoted_rules_are_certified_without_language_name_checks() {
    let grammar = grameo::compile(include_str!("../../../languages/php/src/grammar.pest")).unwrap();
    for name in ["sq_string", "dq_string"] {
        assert!(quoted::certify(&grammar, grammar.rule_id(name).unwrap()).is_some(), "{name}");
    }
}

#[test]
fn any_scalar_escapes_and_doubled_delimiters_match_their_peg_definitions() {
    for (body, inputs) in [
        (r#"("\\" ~ ANY | (!"'" ~ ANY))*"#, vec!["'\\雪'", "'\\é\\''", "'\\\\'", "'\\'"]),
        (r#"("''" | (!"'" ~ ANY))*"#, vec!["'a''b'", "''''", "'''", "'雪''é'"]),
    ] {
        let program = Certified::new(&format!("root = {{ token ~ EOI }} token = @{{ \"'\" ~ {body} ~ \"'\" }}"));
        assert!(program.fragments[program.rule_id("token").unwrap()].is_some());
        for source in inputs {
            let actual = engine::parse_program(&program, "root", source, MatchOptions::default()).map(|tree| captures(&tree));
            let expected = engine::parse_program(&program.grammar, "root", source, MatchOptions::default()).map(|tree| captures(&tree));
            assert_eq!(actual, expected, "{source}");
        }
    }
}

#[cfg(target_os = "macos")]
#[test]
#[ignore = "paired CPU measurement on real WordPress capture parsing"]
fn paired_wordpress_quote_scanning_cpu() {
    unsafe extern "C" { fn clock_gettime(id: i32, time: *mut [i64; 2]) -> i32; }
    fn cpu() -> u128 {
        let mut time = [0i64; 2];
        assert_eq!(unsafe { clock_gettime(16, &mut time) }, 0);
        time[0] as u128 * 1_000_000_000 + time[1] as u128
    }
    let fast = Certified::new(include_str!("../../../languages/php/src/grammar.pest"));
    let mut baseline = Certified::new(include_str!("../../../languages/php/src/grammar.pest"));
    baseline.fragments.fill(None);
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../examples/webroot/wordpress/wp-includes");
    for name in ["class-wp-hook.php", "class-wp-query.php", "formatting.php"] {
        let text = std::fs::read_to_string(root.join(name)).unwrap();
        let source = text.strip_prefix("<?php").unwrap();
        let mut times = [Vec::new(), Vec::new()];
        for round in 0..9 {
            for enabled in if round % 2 == 0 { [false, true] } else { [true, false] } {
                let parser = if enabled { &fast } else { &baseline };
                let start = cpu();
                let tree = engine::parse_program(parser, "program_pure", source, MatchOptions::default()).unwrap();
                std::hint::black_box(&tree);
                drop(tree);
                let elapsed = cpu() - start;
                if round != 0 { times[enabled as usize].push(elapsed); }
            }
        }
        for samples in &mut times { samples.sort_unstable(); }
        eprintln!("quoted_captures file={name} baseline_cpu_ms={:.3} fast_cpu_ms={:.3} ratio={:.3}",
            times[0][4] as f64 / 1e6, times[1][4] as f64 / 1e6, times[1][4] as f64 / times[0][4] as f64);
    }
}

#[test]
fn hybrid_trivia_and_lookahead_keep_captures_hooks_and_errors() {
    use grameo::{builder::Builder, Span};
    #[derive(Debug)]
    struct Ordinary<'a>(&'a CompiledGrammar);
    impl Program for Ordinary<'_> {
        fn rule_id(&self, name: &str) -> Option<usize> { self.0.rule_id(name) }
        fn rule(&self, id: usize) -> RuleSpec<'_> { self.0.rule(id) }
        fn instruction(&self, id: usize) -> Instruction<'_> { self.0.instruction(id) }
        fn scope(&self, id: usize) -> Scope { self.0.scope(id) }
    }
    #[derive(Default)]
    struct Hooks(Vec<(usize, Span)>);
    impl Builder for Hooks {
        fn checkpoint(&self) -> usize { self.0.len() }
        fn rollback(&mut self, mark: usize) { self.0.truncate(mark); }
        fn begin(&mut self, _: usize, _: usize) {}
        fn finish(&mut self, rule: usize, span: Span, _: &str) { self.0.push((rule, span)); }
    }
    let program = Certified::new(r##"root = { &token ~ token ~ EOI }
WHITESPACE = _{ " " }
COMMENT = _{ "#" ~ token }
token = @{ "'" ~ ("\\'" | "\\\\" | (!"'" ~ ANY))* ~ "'" }"##);
    let oracle = Ordinary(&program.grammar);
    for source in ["'hello'", "'hello' #'comment'", "'雪\\'é'", "'hello' #'broken", "'broken", "'hello' junk"] {
        for budget in [usize::MAX, 0, 10, 50, 500] {
            let options = MatchOptions { max_steps: budget, ..MatchOptions::default() };
            let mut fast_hooks = Hooks::default();
            let mut slow_hooks = Hooks::default();
            let actual = engine::build_capture_program(&program, "root", source, options, &mut fast_hooks).map(|tree| captures(&tree));
            let expected = engine::build_capture_program(&oracle, "root", source, options, &mut slow_hooks).map(|tree| captures(&tree));
            assert_eq!(actual, expected, "{source}, budget {budget}");
            assert_eq!(fast_hooks.0, slow_hooks.0, "{source}, budget {budget}");
        }
    }
    assert!(program.scans.get() > 0, "suppressed semantic contexts use certified scans");
}

#[test]
fn bulk_delimited_scanning_preserves_unicode_terminators_errors_and_budgets() {
    for (open, close) in [("/*", "*/"), ("<!--", "-->"), ("始", "終")] {
        let program = Certified::new(&format!("root = {{ token ~ EOI }} token = @{{ {open:?} ~ (!{close:?} ~ ANY)* ~ {close:?} }}"));
        assert!(program.delimited_fragment(program.rule_id("token").unwrap()).is_some());
        for source in [format!("{open}{close}"), format!("{open}雪 café /* nested {close}"), format!("{open}missing"), format!("{open}{close}junk"), format!("{open}{}{close}", "文 / * ".repeat(10000))] {
            for budget in [usize::MAX, 0, 10, 30, 100] {
                let options = MatchOptions { max_steps: budget, ..MatchOptions::default() };
                let a = engine::parse_program(&program, "root", &source, options).map(|tree| captures(&tree));
                let b = engine::parse_program(&program.grammar, "root", &source, options).map(|tree| captures(&tree));
                assert_eq!(a, b, "delimiters {open}/{close}, budget {budget}");
            }
        }
    }
    for body in [r#"(!"x" ~ ANY)*"#, r#"(!"*/" ~ helper)*"#, r#"PUSH((!"*/" ~ ANY)*)"#, r#"(!"*/" ~ ANY)+"#] {
        let program = Certified::new(&format!("token = @{{ \"/*\" ~ {body} ~ \"*/\" }} helper = {{ ANY }}"));
        assert!(program.delimited_fragment(program.rule_id("token").unwrap()).is_none());
    }
    let php = grameo::compile(include_str!("../../../languages/php/src/grammar.pest")).unwrap();
    assert!(quoted::certify_delimited(&php, php.rule_id("block_comment").unwrap()).is_some());
}
