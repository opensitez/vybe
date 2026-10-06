//! Actual dispatched rule/expression attempts, including repeated offsets.
//! cargo run --release --features all-grammars --example lua_profile
use std::{cell::RefCell, collections::HashMap};
use grameo::{MatchOptions, engine::parse_program, program::Program};
use vybe_parser_generated_tests::lua as js;
#[derive(Debug, Default)]
struct Counts {
    rules: HashMap<usize, usize>,
    expressions: HashMap<usize, usize>,
    rule_sites: HashMap<(usize, usize, u8), usize>,
    expression_sites: HashMap<(usize, usize), usize>,
}
#[derive(Debug)]
struct Profile<P>(P, RefCell<Counts>);
impl<P: Program> Program for Profile<P> {
    fn atomic_expression(&self, id: usize) -> bool {
        self.0.atomic_expression(id)
    }
    fn observe_rule_attempt(&self, id: usize, offset: usize, flags: u8) {
        let mut counts = self.1.borrow_mut();
        *counts.rules.entry(id).or_default() += 1;
        *counts.rule_sites.entry((id, offset, flags)).or_default() += 1;
    }
    fn observe_expression(&self, id: usize, offset: usize) {
        let mut counts = self.1.borrow_mut();
        *counts.expressions.entry(id).or_default() += 1;
        *counts.expression_sites.entry((id, offset)).or_default() += 1;
    }
    fn terminal_choice(&self, id: usize) -> Option<grameo::lexical::TerminalChoice<'_>> {
        self.0.terminal_choice(id)
    }
    fn failure_terminals(&self, id: usize) -> &[usize] {
        self.0.failure_terminals(id)
    }
    fn trivia_failure_prefix(&self, id: usize) -> Option<grameo::lexical::FailurePrefix> {
        self.0.trivia_failure_prefix(id)
    }
    fn literal_choice_failure(&self, id: usize, input: &str) -> Option<usize> {
        self.0.literal_choice_failure(id, input)
    }
    fn rule_id(&self, name: &str) -> Option<usize> {
        self.0.rule_id(name)
    }
    fn rule(&self, id: usize) -> grameo::program::RuleSpec<'_> {
        self.0.rule(id)
    }
    fn instruction(&self, id: usize) -> grameo::program::Instruction<'_> {
        self.0.instruction(id)
    }
    fn scope(&self, id: usize) -> grameo::program::Scope {
        self.0.scope(id)
    }
    fn fast_ascii_repetitions(&self) -> bool {
        self.0.fast_ascii_repetitions()
    }
    fn source_pratt(&self, rule: usize) -> Option<grameo::program::SourcePratt<'_>> {
        self.0.source_pratt(rule)
    }
}
fn main() {
    for (label, source) in [
        (
            "declarations100",
            (0..100)
                .map(|n| format!("local x{n} = {n}\n"))
                .collect::<String>(),
        ),
        (
            "expressions100",
            (0..100)
                .map(|n| format!("local x{n} = f(1 + 2 * 3, (4 - 5) / 6)\n"))
                .collect::<String>(),
        ),
    ] {
        report(js::Parser, &format!("original {label}"), &source);
    }
}
fn report<P: Program>(parser: P, label: &str, source: &str) {
    let program = Profile(parser, RefCell::new(Counts::default()));
    let tree = parse_program(&program, "chunk", source, MatchOptions::default()).unwrap();
    let counts = program.1.borrow();
    println!(
        "{label}: {} bytes, {} captures",
        source.len(),
        tree.capture_count()
    );
    println!(
        "  rule attempts {}; distinct rule/offset/flags {}; repeated attempts {}",
        counts.rules.values().sum::<usize>(),
        counts.rule_sites.len(),
        counts.rules.values().sum::<usize>() - counts.rule_sites.len()
    );
    println!(
        "  expression evals {}; distinct expr/offset {}; repeated evals {}",
        counts.expressions.values().sum::<usize>(),
        counts.expression_sites.len(),
        counts.expressions.values().sum::<usize>() - counts.expression_sites.len()
    );
    let mut rules = counts.rules.iter().collect::<Vec<_>>();
    rules.sort_by_key(|(_, n)| std::cmp::Reverse(**n));
    for (id, n) in rules.into_iter().take(25) {
        println!("  rule {:>7} {}", n, program.rule(*id).name);
    }
    let mut expressions = counts.expressions.iter().collect::<Vec<_>>();
    expressions.sort_by_key(|(_, n)| std::cmp::Reverse(**n));
    for (id, n) in expressions.into_iter().take(15) {
        println!(
            "  expression {:>7} #{} {:?}",
            n,
            id,
            program.instruction(*id)
        );
    }
    let mut sites = counts
        .rule_sites
        .iter()
        .filter(|((_, _, flags), n)| **n > 1 && flags & 8 == 0)
        .collect::<Vec<_>>();
    sites.sort_by_key(|(_, n)| std::cmp::Reverse(**n));
    for ((id, offset, flags), n) in sites.into_iter().take(20) {
        println!(
            "  repeated {:>5} {} @{} flags={} next={:?}",
            n,
            program.rule(*id).name,
            offset,
            flags,
            &source[*offset..source.len().min(offset + 28)]
        );
    }
}
