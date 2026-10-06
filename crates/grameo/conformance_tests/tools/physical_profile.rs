//! Untimed observational profile; all Program metadata and certificates forwarded.
use std::{cell::RefCell, collections::BTreeMap, env, fs};
use vybe_parser::{
    MatchOptions,
    grammar::{ExprId, RuleId},
    lexical::{ByteClass, FailurePrefix, LeadingBytes, TerminalChoice},
    program::{Instruction, Program, RuleSpec, Scope, SourcePratt},
};
#[derive(Debug, Default)]
struct Counts {
    rules: BTreeMap<usize, u64>,
    expressions: BTreeMap<usize, u64>,
    opcodes: BTreeMap<&'static str, u64>,
    offsets: BTreeMap<(usize, usize), u64>,
    environment_offsets: BTreeMap<(usize, usize, u8), u64>,
    flags: BTreeMap<u8, u64>,
    metadata: BTreeMap<&'static str, (u64, u64)>,
}
#[derive(Debug)]
struct Profile<'a, P> {
    inner: &'a P,
    counts: RefCell<Counts>,
}
impl<P> Profile<'_, P> {
    fn note(&self, name: &'static str, available: bool) {
        let mut c = self.counts.borrow_mut();
        let n = c.metadata.entry(name).or_default();
        n.0 += 1;
        n.1 += u64::from(available);
    }
}
fn opcode(i: Instruction<'_>) -> &'static str {
    match i {
        Instruction::Literal { .. } => "Literal",
        Instruction::Range { .. } => "Range",
        Instruction::Call(_) => "Call",
        Instruction::Builtin(_) => "Builtin",
        Instruction::Sequence { .. } => "Sequence",
        Instruction::Choice { .. } => "Choice",
        Instruction::Repeat { .. } => "Repeat",
        Instruction::Predicate { .. } => "Predicate",
        Instruction::Group(_) => "Group",
        Instruction::Push(_) => "Push",
        Instruction::PushLiteral(_) => "PushLiteral",
        Instruction::PeekSlice { .. } => "PeekSlice",
    }
}
macro_rules! flag {
    ($name:ident) => {
        fn $name(&self) -> bool {
            let v = self.inner.$name();
            self.note(stringify!($name), v);
            v
        }
    };
}
impl<P: Program> Program for Profile<'_, P> {
    fn observe_rule_attempt(&self, id: RuleId, offset: usize, flags: u8) {
        let mut c = self.counts.borrow_mut();
        *c.rules.entry(id).or_default() += 1;
        *c.offsets.entry((id, offset)).or_default() += 1;
        *c.environment_offsets
            .entry((id, offset, flags))
            .or_default() += 1;
        *c.flags.entry(flags).or_default() += 1;
        drop(c);
        self.inner.observe_rule_attempt(id, offset, flags);
    }
    fn observe_expression(&self, id: ExprId, offset: usize) {
        let op = opcode(self.inner.instruction(id));
        let mut c = self.counts.borrow_mut();
        *c.expressions.entry(id).or_default() += 1;
        *c.opcodes.entry(op).or_default() += 1;
        drop(c);
        self.inner.observe_expression(id, offset);
    }
    fn leading_bytes(&self, id: ExprId) -> Option<LeadingBytes> {
        let v = self.inner.leading_bytes(id);
        self.note("leading_bytes", v.is_some());
        v
    }
    fn atomic_expression(&self, id: ExprId) -> bool {
        let v = self.inner.atomic_expression(id);
        self.note("atomic_expression", v);
        v
    }
    fn terminal_candidates(&self, id: ExprId, b: Option<u8>) -> Option<&[usize]> {
        let v = self.inner.terminal_candidates(id, b);
        self.note("terminal_candidates", v.is_some());
        v
    }
    fn terminal_choice(&self, id: ExprId) -> Option<TerminalChoice<'_>> {
        let v = self.inner.terminal_choice(id);
        self.note("terminal_choice", v.is_some());
        v
    }
    fn failure_terminals(&self, id: RuleId) -> &[ExprId] {
        let v = self.inner.failure_terminals(id);
        self.note("failure_terminals", !v.is_empty());
        v
    }
    fn literal_choice_failure(&self, id: ExprId, source: &str) -> Option<usize> {
        let v = self.inner.literal_choice_failure(id, source);
        self.note("literal_choice_failure", v.is_some());
        v
    }
    fn trivia_failure_prefix(&self, id: RuleId) -> Option<FailurePrefix> {
        let v = self.inner.trivia_failure_prefix(id);
        self.note("trivia_failure_prefix", v.is_some());
        v
    }
    flag!(fast_rule_failures);
    flag!(atomic_tail_dispatch);
    flag!(tail_continuations);
    flag!(tail_dispatch);
    flag!(fast_trivia_failures);
    flag!(fast_predicate_failures);
    flag!(fast_ascii_repetitions);
    fn rule_id(&self, name: &str) -> Option<RuleId> {
        self.note("rule_id", false);
        self.inner.rule_id(name)
    }
    fn rule(&self, id: RuleId) -> RuleSpec<'_> {
        self.note("rule", false);
        self.inner.rule(id)
    }
    fn instruction(&self, id: ExprId) -> Instruction<'_> {
        self.note("instruction", false);
        self.inner.instruction(id)
    }
    fn source_pratt(&self, id: RuleId) -> Option<SourcePratt<'_>> {
        let v = self.inner.source_pratt(id);
        self.note("source_pratt", v.is_some());
        v
    }
    fn scope(&self, id: usize) -> Scope {
        self.note("scope", false);
        self.inner.scope(id)
    }
    fn whitespace_class(&self) -> Option<ByteClass> {
        let v = self.inner.whitespace_class();
        self.note("whitespace_class", v.is_some());
        v
    }
    fn whitespace_prefix_class(&self) -> Option<ByteClass> {
        let v = self.inner.whitespace_prefix_class();
        self.note("whitespace_prefix_class", v.is_some());
        v
    }
}
type Shape = Vec<(String, usize, usize, usize)>;
fn shape(tree: vybe_parser::ParseTree<'_, '_>) -> (usize, Shape) {
    fn visit(p: vybe_parser::Pair<'_, '_, '_>, d: usize, s: &mut Shape) {
        let span = p.as_span();
        s.push((p.rule_name().into(), span.start, span.end, d));
        for c in p.into_inner() {
            visit(c, d + 1, s);
        }
    }
    let mut s = vec![];
    for p in tree.pairs() {
        visit(p, 0, &mut s);
    }
    (tree.consumed(), s)
}
fn emit<P: Program>(p: &Profile<'_, P>, phase: &str) {
    let c = p.counts.borrow();
    println!(
        "TOTAL\t{phase}\t{}\t{}\t{}\t{}",
        c.rules.values().sum::<u64>(),
        c.expressions.values().sum::<u64>(),
        c.offsets.len(),
        c.offsets.values().map(|&n| n - 1).sum::<u64>()
    );
    println!(
        "EXACT_FLAGS_REPEATS\t{phase}\t{}",
        c.environment_offsets.values().map(|&n| n - 1).sum::<u64>()
    );
    for (&id, &n) in &c.rules {
        let repeated = c
            .offsets
            .iter()
            .filter(|((r, _), _)| *r == id)
            .map(|(_, n)| n - 1)
            .sum::<u64>();
        println!(
            "RULE\t{phase}\t{id}\t{}\t{n}\t{repeated}",
            p.inner.rule(id).name
        );
        let spec = p.inner.rule(id);
        let mut pending = vec![spec.expression];
        let mut visited = std::collections::BTreeSet::new();
        let mut calls = std::collections::BTreeSet::new();
        let mut opcodes = BTreeMap::<&str, usize>::new();
        while let Some(expr) = pending.pop() {
            if !visited.insert(expr) {
                continue;
            }
            let instruction = p.inner.instruction(expr);
            *opcodes.entry(opcode(instruction)).or_default() += 1;
            match instruction {
                Instruction::Call(rule) => {
                    calls.insert(p.inner.rule(rule).name);
                }
                Instruction::Sequence { left, right } | Instruction::Choice { left, right } => {
                    pending.extend([left, right])
                }
                Instruction::Repeat { expression, .. }
                | Instruction::Predicate { expression, .. } => pending.push(expression),
                Instruction::Group(child) | Instruction::Push(child) => pending.push(child),
                _ => {}
            }
        }
        let exact = c
            .environment_offsets
            .iter()
            .filter(|((r, _, _), _)| *r == id)
            .map(|(_, n)| n - 1)
            .sum::<u64>();
        println!(
            "RULEMETA\t{phase}\t{id}\t{:?}\t{}\t{}\t{exact}\t{}\t{:?}",
            spec.mode,
            spec.expression,
            p.inner.atomic_expression(spec.expression),
            calls.into_iter().collect::<Vec<_>>().join(","),
            opcodes
        );
    }
    for (&id, &n) in &c.expressions {
        println!("EXPR\t{phase}\t{id}\t{n}\t{:?}", p.inner.instruction(id));
    }
    for (name, n) in &c.opcodes {
        println!("OPCODE\t{phase}\t{name}\t{n}");
    }
    for (flags, n) in &c.flags {
        println!("FLAGS\t{phase}\t{flags}\t{n}");
    }
    for (name, (queries, available)) in &c.metadata {
        println!("META\t{phase}\t{name}\t{queries}\t{available}");
    }
}
fn run<P: Program>(inner: &P, entry: &str, source: &str, grammar: &str) {
    let options = MatchOptions::default();
    let plain = shape(vybe_parser::engine::parse_program(inner, entry, source, options).unwrap());
    let wrapper = Profile {
        inner,
        counts: RefCell::new(Counts::default()),
    };
    let observed =
        shape(vybe_parser::engine::parse_program(&wrapper, entry, source, options).unwrap());
    assert_eq!(plain, observed);
    emit(&wrapper, "capture");
    let resolved = vybe_parser::compile(grammar).unwrap();
    for &id in wrapper.counts.borrow().rules.keys() {
        let name = inner.rule(id).name;
        let rule = resolved
            .syntax()
            .rules
            .iter()
            .find(|r| r.name == name)
            .unwrap();
        assert_eq!(resolved.rule_id(name), Some(id));
        assert_eq!(rule.expression, inner.rule(id).expression);
        let text = &grammar[rule.span.start..rule.span.end];
        let hex = text.bytes().map(|b| format!("{b:02x}")).collect::<String>();
        println!("SOURCE\t{id}\t{hex}");
    }
    println!("CAPTURES\t{}\t{}", plain.0, plain.1.len());
    let expected = vybe_parser::engine::recognize_program(inner, entry, source, options).unwrap();
    let wrapper = Profile {
        inner,
        counts: RefCell::new(Counts::default()),
    };
    let actual = vybe_parser::engine::recognize_program(&wrapper, entry, source, options).unwrap();
    assert_eq!(expected.consumed, actual.consumed);
    assert_eq!(expected.steps, actual.steps);
    emit(&wrapper, "recognition");
    println!("STEPS\t{}", actual.steps);
}
fn main() {
    std::thread::Builder::new()
        .stack_size(64 * 1024 * 1024)
        .spawn(|| {
            let a = env::args().collect::<Vec<_>>();
            let source = fs::read_to_string(&a[2]).unwrap();
            let grammar =
                fs::read_to_string(format!("{}/{}/src/grammar.pest", a[3], a[1])).unwrap();
            match a[1].as_str() {
                "ruby" => run(
                    &vybe_parser_generated_tests::all::ruby::Parser,
                    "program",
                    &source,
                    &grammar,
                ),
                "fortran" => run(
                    &vybe_parser_generated_tests::all::fortran::Parser,
                    "program",
                    &source,
                    &grammar,
                ),
                "washm" => run(
                    &vybe_parser_generated_tests::all::washm::Parser,
                    "program",
                    &source,
                    &grammar,
                ),
                _ => panic!("unsupported language"),
            }
        })
        .unwrap()
        .join()
        .unwrap();
}
