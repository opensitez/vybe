//! Token certificates do not authorize skipping a custom Program's observers.
use std::cell::{Cell, RefCell};
use grameo::{CompiledGrammar, MatchOptions, engine, program::{Instruction, Program, RuleSpec, Scope}, quoted};

#[derive(Debug)]
struct Tracing {
    grammar: CompiledGrammar,
    certificates: bool,
    queries: Cell<usize>,
    events: RefCell<Vec<(usize, usize, Option<u8>)>>,
}
impl Program for Tracing {
    fn rule_id(&self, name: &str) -> Option<usize> { self.grammar.rule_id(name) }
    fn rule(&self, id: usize) -> RuleSpec<'_> { self.grammar.rule(id) }
    fn instruction(&self, id: usize) -> Instruction<'_> { self.grammar.instruction(id) }
    fn scope(&self, id: usize) -> Scope { self.grammar.scope(id) }
    fn atomic_expression(&self, id: usize) -> bool { self.grammar.atomic_expression(id) }
    fn observe_expression(&self, id: usize, offset: usize) {
        self.events.borrow_mut().push((id, offset, None));
    }
    fn observe_rule_attempt(&self, id: usize, offset: usize, flags: u8) {
        self.events.borrow_mut().push((id, offset, Some(flags)));
    }
    fn quoted_fragment(&self, id: usize) -> Option<quoted::Fragment> {
        self.queries.set(self.queries.get() + 1);
        self.certificates.then(|| quoted::certify(&self.grammar, id)).flatten()
    }
    fn delimited_fragment(&self, id: usize) -> Option<quoted::DelimitedFragment<'_>> {
        self.queries.set(self.queries.get() + 1);
        self.certificates.then(|| quoted::certify_delimited(&self.grammar, id)).flatten()
    }
}
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
fn token_scanners_preserve_observers_at_both_rule_entry_paths() {
    for token in [
        r#"token = @{ "'" ~ ("\\'" | "\\\\" | (!"'" ~ ANY))* ~ "'" }"#,
        r#"token = @{ "/*" ~ (!"*/" ~ ANY)* ~ "*/" }"#,
    ] {
        for nested in [false, true] {
            let source = format!("root = {{ SOI ~ {} ~ EOI }}\n{}\n{token}",
                if nested { "wrapper" } else { "token" },
                if nested { "wrapper = @{ token }" } else { "" });
            for input in ["''", "'snow 雪'", "'a\\'b'", "'missing", "/*snow 雪*/", "/*missing"] {
                for (work, depth) in [(usize::MAX, usize::MAX), (0, 0), (40, 2), (1000, 20)] {
                    let new = |certificates| Tracing { grammar: grameo::compile(&source).unwrap(), certificates,
                        queries: Cell::new(0), events: RefCell::new(Vec::new()) };
                    let control = new(false);
                    let candidate = new(true);
                    let options = MatchOptions { max_steps: work, max_rule_depth: depth };
                    let a = engine::parse_program(&control, "root", input, options).map(|tree| captures(&tree));
                    let b = engine::parse_program(&candidate, "root", input, options).map(|tree| captures(&tree));
                    assert_eq!(a, b, "{input:?} nested={nested} work={work} depth={depth}");
                    assert_eq!(*control.events.borrow(), *candidate.events.borrow());
                    assert_eq!(candidate.queries.get(), 0, "observer permission precedes certificates");
                }
            }
        }
    }
}
