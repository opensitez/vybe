use grameo::{builder::{Builder, BuildError}, engine::{build_capture_program, build_capture_program_at}, MatchOptions, Span};

#[derive(Default)]
struct Hooks { events: Vec<(usize, Span)>, commits: Vec<usize>, reject: bool, target: usize }
impl Builder for Hooks {
    fn checkpoint(&self) -> usize { self.events.len() }
    fn rollback(&mut self, mark: usize) { self.events.truncate(mark); }
    fn begin(&mut self, _: usize, _: usize) {}
    fn finish(&mut self, rule: usize, span: Span, _: &str) { self.events.push((rule, span)); }
    fn try_finish(&mut self, rule: usize, span: Span, source: &str) -> Result<(), BuildError> {
        self.finish(rule, span, source);
        if self.reject { Err(BuildError { span, message: "rejected semantic unit".into() }) } else { Ok(()) }
    }
    fn committed(&mut self, _: &str, end: usize) { self.commits.push(end); }
    fn supports_pratt(&self, rule: usize) -> bool { rule == self.target }
}

#[test]
fn successful_units_keep_absolute_captures_and_failed_next_unit_preserves_prior_values() {
    let grammar = grameo::compile(r#"unit = { word ~ ";" } word = @{ ASCII_ALPHA+ }"#).unwrap();
    let source = "λ:one;two;bad?";
    let mut hooks = Hooks::default();
    for (start, end, raw) in [(3, 7, "one;"), (7, 11, "two;")] {
        let tree = grammar.build_captures_at("unit", source, start, &mut hooks).unwrap();
        assert_eq!(tree.consumed(), end);
        assert_eq!(tree.source().as_ptr(), source.as_ptr());
        let root = tree.pairs().next().unwrap();
        assert_eq!(root.as_span(), Span { start, end });
        assert_eq!(root.as_str(), raw);
    }
    let prior = hooks.events.clone();
    assert_eq!(hooks.commits, [7, 11]);
    let error = grammar.build_captures_at("unit", source, 11, &mut hooks).unwrap_err();
    assert_eq!(error.offset, 14);
    assert_eq!(hooks.events, prior);
    assert_eq!(hooks.commits, [7, 11]);
    hooks.reject = true;
    assert!(grammar.build_captures_at("unit", source, 3, &mut hooks).is_err());
    assert_eq!(hooks.events, prior);
    assert_eq!(hooks.commits, [7, 11]);
}

#[test]
fn boundary_validation_and_full_input_anchors_preserve_the_builder() {
    let grammar = grameo::compile(r#"unit = { ANY } soi = { SOI ~ ANY } eoi = { ANY ~ EOI }"#).unwrap();
    let mut hooks = Hooks::default();
    for start in [1, 4, usize::MAX] {
        let error = grammar.build_captures_at("unit", "λ:x", start, &mut hooks).unwrap_err();
        assert_eq!(error.offset, start);
        assert!(hooks.events.is_empty() && hooks.commits.is_empty());
    }
    assert!(grammar.build_captures_at("soi", "λ:x", 3, &mut hooks).is_err());
    assert!(grammar.build_captures_at("eoi", "λ:xy", 3, &mut hooks).is_err());
    assert!(grammar.build_captures_at("missing", "λ:x", 3, &mut hooks).is_err());
    assert!(hooks.events.is_empty() && hooks.commits.is_empty());
}

#[test]
fn semantic_offset_entry_matches_zero_entry_hooks_errors_pratt_and_resource_budgets() {
    let grammar = grameo::compile_with_options(r#"
// @grameo pratt expr atom=atom trailing=consume
// @grameo operator expr plus precedence=10 fixity=infix_left
WHITESPACE = _{ " " }
unit = { expr ~ ";" }
expr = { atom ~ (plus ~ atom)* }
atom = @{ ASCII_DIGIT+ }
plus = @{ "+" }
"#, grameo::CompileOptions { optimizations: true, ..Default::default() }).unwrap();
    let target = grammar.rule_id("expr").unwrap();
    for raw in ["1;tail", "1 + 2;tail", "1+?", "1+2?"] {
        let full = format!("λ:{raw}");
        for max_rule_depth in 0..12 { for max_steps in (0..100).chain([usize::MAX]) {
            let options = MatchOptions { max_steps, max_rule_depth };
            let mut baseline = Hooks { target, ..Default::default() };
            let mut offset = Hooks { target, ..Default::default() };
            let a = build_capture_program(&grammar, "unit", raw, options, &mut baseline);
            let b = build_capture_program_at(&grammar, "unit", &full, 3, options, &mut offset);
            match (a,b) {
                (Ok(a), Ok(b)) => {
                    assert_eq!(a.consumed()+3,b.consumed());
                    assert_eq!(a.capture_count(),b.capture_count());
                    assert_eq!(baseline.events.iter().map(|(r,s)|(*r,Span{start:s.start+3,end:s.end+3})).collect::<Vec<_>>(),offset.events);
                    assert_eq!(baseline.commits.iter().map(|end|end+3).collect::<Vec<_>>(),offset.commits);
                },
                (Err(a), Err(b)) => {
                    assert_eq!(a.kind,b.kind); assert_eq!(a.offset+3,b.offset);
                    assert_eq!(a.expected,b.expected); assert_eq!(a.message,b.message);
                    assert!(baseline.events.is_empty() && offset.events.is_empty());
                    assert!(baseline.commits.is_empty() && offset.commits.is_empty());
                },
                _ => panic!("unit entry differs for {raw:?}, depth={max_rule_depth}, work={max_steps}"),
            }
        }}
    }
}
