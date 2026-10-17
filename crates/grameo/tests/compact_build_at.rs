use grameo::{
    CompiledGrammar, MatchOptions, ParseTree, Span,
    builder::{BuildError, Builder},
    engine::{
        build_capture_program_at, build_compact_capture_program_at,
        build_stackfree_compact_capture_program_at,
    },
    program::{CompactScope, Instruction, Program, RuleSpec, Scope, SourcePratt},
    tree::{CaptureHandle, Pair},
};

// CompiledGrammar deliberately does not promise compact scopes. This test
// adapter certifies its scope words explicitly; stack-free admission is checked
// separately against the same whole-arena exclusions used by code generation.
#[derive(Debug)]
struct CompactProgram(CompiledGrammar, bool);

impl CompactProgram {
    fn new(source: &str, fragments: bool) -> Self {
        let grammar = grameo::compile_with_options(
            source,
            grameo::CompileOptions {
                optimizations: true,
                ..Default::default()
            },
        )
        .unwrap();
        for scope in 0..grammar.scopes().len() {
            let _ = CompactScope::new(scope);
        }
        Self(grammar, fragments)
    }

    fn certify_stack_free(&self) {
        use grameo::Builtin;
        for id in 0..self.0.syntax().expressions.len() {
            assert!(!matches!(
                self.instruction(id),
                Instruction::Push(_)
                    | Instruction::PushLiteral(_)
                    | Instruction::PeekSlice { .. }
                    | Instruction::Builtin(
                        Builtin::Peek
                            | Builtin::PeekAll
                            | Builtin::Pop
                            | Builtin::PopAll
                            | Builtin::Drop
                    )
            ));
        }
    }
}

impl Program for CompactProgram {
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
    fn compact_capture_scope(&self, id: usize) -> CompactScope {
        CompactScope::new(id)
    }
    fn atomic_fragments(&self) -> Option<()> {
        self.1.then_some(())
    }
    fn source_pratt(&self, id: usize) -> Option<SourcePratt<'_>> {
        self.0.source_pratt(id)
    }
    fn source_pratt_operator_rhs_atom(&self, id: usize, op: usize) -> Option<usize> {
        self.0.source_pratt_operator_rhs_atom(id, op)
    }
    fn source_pratt_operator_trivia(&self, id: usize) -> Option<usize> {
        self.0.source_pratt_operator_trivia(id)
    }
    fn source_pratt_rhs_atom(&self, id: usize) -> Option<usize> {
        self.0.source_pratt_rhs_atom(id)
    }
    fn source_pratt_rhs_operator(&self, id: usize) -> Option<usize> {
        self.0.source_pratt_rhs_operator(id)
    }
}

#[derive(Debug, PartialEq, Eq)]
enum Event {
    Begin(usize, usize),
    Finish(usize, Span, Option<CaptureHandle>),
    Reduce(usize, Span),
}

#[derive(Debug, Default, PartialEq, Eq)]
struct Hooks {
    events: Vec<Event>,
    commits: Vec<usize>,
    pratt: Option<usize>,
    policy: u8,
}

impl Builder for Hooks {
    fn checkpoint(&self) -> usize {
        self.events.len()
    }
    fn rollback(&mut self, mark: usize) {
        self.events.truncate(mark);
    }
    fn begin(&mut self, rule: usize, offset: usize) {
        self.events.push(Event::Begin(rule, offset));
    }
    fn finish(&mut self, _: usize, _: Span, _: &str) {
        unreachable!("checked capture hook is used");
    }
    fn observes_capture(&self, _: usize) -> bool {
        true
    }
    fn try_finish_capture(
        &mut self,
        rule: usize,
        span: Span,
        _: &str,
        capture: Option<CaptureHandle>,
    ) -> Result<(), BuildError> {
        self.events.push(Event::Finish(rule, span, capture));
        if self.policy == 1 {
            Err(BuildError {
                span,
                message: "finish rejection".into(),
            })
        } else {
            Ok(())
        }
    }
    fn try_committed(&mut self, _: &str, end: usize) -> Result<(), BuildError> {
        if self.policy == 2 {
            return Err(BuildError {
                span: Span { start: end, end },
                message: "commit rejection".into(),
            });
        }
        self.commits.push(end);
        Ok(())
    }
    fn supports_pratt(&self, rule: usize) -> bool {
        self.pratt == Some(rule)
    }
    fn accepts_pratt(&self, _: usize, _: Span) -> bool {
        self.policy != 3
    }
    fn reduce(&mut self, rule: usize, _: grameo::pratt::Fixity, span: Span) {
        self.events.push(Event::Reduce(rule, span));
    }
}

fn captures(tree: &ParseTree<'_, '_>) -> Vec<(usize, String, Span)> {
    fn visit(pair: Pair<'_, '_, '_>, depth: usize, out: &mut Vec<(usize, String, Span)>) {
        out.push((depth, pair.rule_name().into(), pair.as_span()));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut out = Vec::new();
    for pair in tree.pairs() {
        visit(pair, 0, &mut out);
    }
    out
}

fn compare(
    program: &CompactProgram,
    entry: &str,
    source: &str,
    start: usize,
    options: MatchOptions,
    pratt: Option<usize>,
    policy: u8,
    stack_free: bool,
) {
    let seed = || Hooks {
        events: vec![Event::Begin(usize::MAX, 0)],
        commits: vec![0],
        pratt,
        policy,
    };
    let mut reference = seed();
    let mut candidate = seed();
    let a = build_capture_program_at(program, entry, source, start, options, &mut reference);
    let b = if stack_free {
        build_stackfree_compact_capture_program_at(
            program,
            entry,
            source,
            start,
            options,
            &mut candidate,
        )
    } else {
        build_compact_capture_program_at(program, entry, source, start, options, &mut candidate)
    };
    let failed = a.is_err();
    match (a, b) {
        (Ok(a), Ok(b)) => {
            assert_eq!(a.consumed(), b.consumed());
            assert_eq!(a.capture_count(), b.capture_count());
            assert_eq!(captures(&a), captures(&b));
            assert_eq!(a.source().as_ptr(), source.as_ptr());
            assert_eq!(b.source().as_ptr(), source.as_ptr());
        }
        (Err(a), Err(b)) => {
            assert_eq!(a.kind, b.kind);
            assert_eq!(a.offset, b.offset);
            assert_eq!(a.expected, b.expected);
            assert_eq!(a.message, b.message);
        }
        _ => panic!("offset backends differ: {source:?}, {start}, {options:?}, policy {policy}"),
    }
    assert_eq!(reference, candidate);
    if failed {
        assert_eq!(reference, seed());
    } else {
        assert_eq!(reference.commits.len(), 2);
    }
}

#[test]
fn compact_offset_captures_pratt_hooks_errors_and_budgets_match_env() {
    let source = r#"
// @grameo pratt expr atom=atom trailing=consume
// @grameo operator expr plus precedence=10 fixity=infix_left
WHITESPACE = _{ " " }
unit = { &(expr ~ ";") ~ ((expr ~ "!") | (expr ~ ";")) }
expr = { atom ~ (plus ~ atom)* }
atom = @{ ASCII_DIGIT+ }
plus = @{ "+" }
"#;
    for fragments in [false, true] {
        let program = CompactProgram::new(source, fragments);
        program.certify_stack_free();
        assert!(
            program
                .source_pratt(program.rule_id("expr").unwrap())
                .is_some()
        );
        let mut hooks = Hooks {
            pratt: program.rule_id("expr"),
            ..Default::default()
        };
        let _ = build_stackfree_compact_capture_program_at(
            &program,
            "unit",
            "λ:1+2;tail",
            3,
            MatchOptions::default(),
            &mut hooks,
        )
        .unwrap();
        assert!(
            hooks
                .events
                .iter()
                .any(|event| matches!(event, Event::Reduce(..)))
        );
        for input in ["λ:1;tail", "λ:1 + 2;tail", "λ:1+?", "λ:1+2?"] {
            for pratt in [None, program.rule_id("expr")] {
                for policy in 0..4 {
                    for max_rule_depth in [0, 1, 4, 8, 32] {
                        for max_steps in (0..120).chain([usize::MAX]) {
                            let options = MatchOptions {
                                max_steps,
                                max_rule_depth,
                            };
                            for stack_free in [false, true] {
                                compare(
                                    &program, "unit", input, 3, options, pratt, policy, stack_free,
                                );
                            }
                        }
                    }
                }
            }
        }
    }
}

#[test]
fn compact_offsets_validate_utf8_full_source_anchors_and_entries() {
    let program = CompactProgram::new("unit={ANY} soi={SOI~ANY} eoi={ANY~EOI}", false);
    program.certify_stack_free();
    for source in ["λ:xy", "λ", ""] {
        for start in [0, 1, 2, 3, source.len(), source.len() + 1, usize::MAX] {
            for entry in ["unit", "soi", "eoi", "missing", "ANY"] {
                for stack_free in [false, true] {
                    compare(
                        &program,
                        entry,
                        source,
                        start,
                        MatchOptions::default(),
                        None,
                        0,
                        stack_free,
                    );
                }
            }
        }
    }
}

#[test]
fn compact_units_commit_once_and_failed_next_unit_preserves_prior_receipts() {
    let program = CompactProgram::new("unit={word~\";\"} word=@{ASCII_ALPHA+}", false);
    program.certify_stack_free();
    for stack_free in [false, true] {
        let mut reference = Hooks::default();
        let mut candidate = Hooks::default();
        for start in [3, 7, 11] {
            let source = "λ:one;two;bad?";
            let a = build_capture_program_at(
                &program,
                "unit",
                source,
                start,
                MatchOptions::default(),
                &mut reference,
            );
            let b = if stack_free {
                build_stackfree_compact_capture_program_at(
                    &program,
                    "unit",
                    source,
                    start,
                    MatchOptions::default(),
                    &mut candidate,
                )
            } else {
                build_compact_capture_program_at(
                    &program,
                    "unit",
                    source,
                    start,
                    MatchOptions::default(),
                    &mut candidate,
                )
            };
            assert_eq!(a.is_ok(), start != 11);
            assert_eq!(a.is_ok(), b.is_ok());
            if let Ok(tree) = b {
                assert_eq!(tree.consumed(), start + 4);
                let root = tree.pairs().next().unwrap();
                assert_eq!(
                    root.as_span(),
                    Span {
                        start,
                        end: start + 4
                    }
                );
                let word = root.into_inner().next().unwrap();
                assert_eq!(
                    word.as_span(),
                    Span {
                        start,
                        end: start + 3
                    }
                );
                assert_eq!(word.as_str(), if start == 3 { "one" } else { "two" });
            }
            assert_eq!(reference, candidate);
        }
        assert_eq!(candidate.commits, [7, 11]);
        assert!(candidate.events.iter().all(|event| match event {
            Event::Begin(_, offset) => *offset < 11,
            Event::Finish(_, span, _) | Event::Reduce(_, span) => span.end <= 11,
        }));
    }
}

#[test]
fn compact_stackful_offsets_keep_predicate_and_alternative_undo() {
    for fragments in [false, true] {
        let program = CompactProgram::new(
            r#"unit={PUSH("ab")~&PEEK~(POP~"!"|PEEK~POP~";")}"#,
            fragments,
        );
        for source in ["λ:abab!", "λ:ababab;", "λ:abab?", "λ:ab"] {
            for policy in 0..3 {
                for max_steps in (0..100).chain([usize::MAX]) {
                    compare(
                        &program,
                        "unit",
                        source,
                        3,
                        MatchOptions {
                            max_steps,
                            ..Default::default()
                        },
                        None,
                        policy,
                        false,
                    );
                }
            }
        }
    }
}

#[test]
fn existing_env_offset_api_does_not_require_compact_certification() {
    let grammar = grameo::compile(r#"unit={PUSH("ab")~POP~EOI}"#).unwrap();
    let mut hooks = Hooks::default();
    let tree = build_capture_program_at(
        &grammar,
        "unit",
        "λ:abab",
        3,
        MatchOptions::default(),
        &mut hooks,
    )
    .unwrap();
    assert_eq!(tree.consumed(), 7);
    assert_eq!(hooks.commits, [7]);
    let prior = hooks.events.len();
    assert!(
        build_capture_program_at(
            &grammar,
            "unit",
            "λ:abac",
            3,
            MatchOptions::default(),
            &mut hooks
        )
        .is_err()
    );
    assert_eq!(hooks.events.len(), prior);
    assert_eq!(hooks.commits, [7]);
}

#[test]
#[should_panic(expected = "program has not certified compact capture scopes")]
fn ordinary_compiled_grammar_is_not_implicitly_compact() {
    let grammar = grameo::compile("unit={ANY}").unwrap();
    let _ = build_stackfree_compact_capture_program_at(
        &grammar,
        "unit",
        "x",
        0,
        MatchOptions::default(),
        &mut Hooks::default(),
    );
}

#[derive(Debug)]
struct MaximumScope;

impl Program for MaximumScope {
    fn rule_id(&self, name: &str) -> Option<usize> {
        (name == "unit").then_some(usize::MAX)
    }
    fn rule(&self, id: usize) -> RuleSpec<'_> {
        assert_eq!(id, usize::MAX);
        RuleSpec {
            name: "unit",
            mode: grameo::RuleMode::Normal,
            expression: usize::MAX,
            scope: usize::MAX,
        }
    }
    fn instruction(&self, id: usize) -> Instruction<'_> {
        assert_eq!(id, usize::MAX);
        Instruction::Literal {
            text: "x",
            insensitive: false,
        }
    }
    fn scope(&self, id: usize) -> Scope {
        assert!(id == 0 || id == usize::MAX);
        Scope::default()
    }
}

#[test]
fn existing_env_offset_api_keeps_full_width_custom_scope_ids() {
    let mut hooks = Hooks::default();
    let tree = build_capture_program_at(
        &MaximumScope,
        "unit",
        "λ:x",
        3,
        MatchOptions::default(),
        &mut hooks,
    )
    .unwrap();
    assert_eq!(tree.consumed(), 4);
    assert_eq!(
        tree.pairs().next().unwrap().as_span(),
        Span { start: 3, end: 4 }
    );
    assert_eq!(hooks.commits, [4]);
}
