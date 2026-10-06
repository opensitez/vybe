//! Unchanged JavaScript grammar: interleaved capture benchmarks and tree parity.
//! cargo run --release --features javascript-benchmark --example javascript_baseline
//! Timings include capture allocation/destruction, exclude AST/compiler work.
//! Before disables only atomic token execution; previous dispatch stays enabled.
use pest::Parser as _;
use pest_derive::Parser;
use std::{hint::black_box, time::Instant};
use grameo::{MatchOptions, ParseTree, engine::parse_program};
#[allow(dead_code, clippy::upper_case_acronyms)]
mod js {
    include!(concat!(env!("OUT_DIR"), "/javascript_parser.rs"));
}

#[derive(Parser)]
#[grammar = "../../../languages/js/src/grammar.pest"]
struct ReferenceParser;

#[derive(Debug)]
struct WithoutAtomic<P>(P);
impl<P: grameo::program::Program> grameo::program::Program for WithoutAtomic<P> {
    // atomic_expression retains Program's false default, disabling only the
    // compact atomic token evaluator. All prior dispatch proofs remain enabled.
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

type Shape = Vec<(String, usize, usize, usize)>;
fn reference_shape(pairs: pest::iterators::Pairs<'_, Rule>) -> Shape {
    fn visit(pair: pest::iterators::Pair<'_, Rule>, depth: usize, out: &mut Shape) {
        let span = pair.as_span();
        out.push((
            format!("{:?}", pair.as_rule()),
            span.start(),
            span.end(),
            depth,
        ));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut result = Vec::new();
    for pair in pairs {
        visit(pair, 0, &mut result);
    }
    result
}
fn native_shape(tree: &ParseTree<'_, '_>) -> Shape {
    fn visit(pair: grameo::Pair<'_, '_, '_>, depth: usize, out: &mut Shape) {
        let span = pair.as_span();
        out.push((pair.rule_name().to_owned(), span.start, span.end, depth));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut result = Vec::new();
    for pair in tree.pairs() {
        visit(pair, 0, &mut result);
    }
    result
}
fn paired(source: &str, count: usize) -> [(f64, f64); 3] {
    let before = WithoutAtomic(js::Parser);
    let mut times: [Vec<f64>; 3] = std::array::from_fn(|_| Vec::with_capacity(count));
    for round in 0..count + 10 {
        for index in 0..3 {
            let variant = (index + round) % 3;
            let now = Instant::now();
            match variant {
                0 => {
                    black_box(
                        parse_program(
                            &before,
                            "program",
                            black_box(source),
                            MatchOptions::default(),
                        )
                        .unwrap(),
                    );
                }
                1 => {
                    black_box(
                        js::Parser
                            .parse(js::Rule::program, black_box(source))
                            .unwrap(),
                    );
                }
                _ => {
                    black_box(ReferenceParser::parse(Rule::program, black_box(source)).unwrap());
                }
            }
            if round >= 10 {
                times[variant].push(now.elapsed().as_secs_f64() * 1e6);
            }
        }
    }
    times.map(|mut times| {
        times.sort_by(f64::total_cmp);
        (times[count / 2], times[count * 95 / 100])
    })
}
fn main() {
    let count = std::env::args()
        .nth(1)
        .map(|n| n.parse::<usize>().expect("sample count"))
        .unwrap_or(101);
    assert!(count > 0);
    let grammar =
        grameo::compile(include_str!("../../../../languages/js/src/grammar.pest")).unwrap();
    for (label, source) in [
        ("tiny", "const x = 1 + 2 * 3;".to_owned()),
        ("declarations100", (0..100).map(|n| format!("const x{n} = {n};\n")).collect()),
        ("expressions100", (0..100).map(|n| format!("const x{n} = f(1 + 2 * 3, (4 - 5) / 6);\n")).collect()),
        ("comments100", (0..100).map(|n| format!("// comment {n}\nconst x{n} = 'é😀';\n")).collect()),
        ("modern", "class Box { constructor(value) { this.value = value; } get() { return this.value; } }\nconst f = (x) => x?.value ?? 0;\nconst xs = [1, 2, 3].map(x => ({ value: x }));\nconst text = `sum: ${f(xs[0])}`;\n".to_owned()),
    ] {
        let expected = reference_shape(ReferenceParser::parse(Rule::program, &source).unwrap());
        let generated = js::Parser.parse(js::Rule::program, &source).unwrap();
        let dynamic = grammar.parse_source("program", &source).unwrap();
        let previous = parse_program(&WithoutAtomic(js::Parser), "program", &source, MatchOptions::default()).unwrap();
        assert_eq!(generated.consumed(), source.len());
        assert_eq!(native_shape(&generated), expected, "{label}: generated capture shape");
        assert_eq!(native_shape(&dynamic), expected, "{label}: resolved capture shape");
        assert_eq!(native_shape(&previous), expected, "{label}: previous dispatch capture shape");
        println!("{label}: {} bytes, {} captures", source.len(), generated.capture_count());
        let [before, current, pest] = paired(&source, count);
        println!("  before median/p95 {before:?} us; current {current:?} us; pest {pest:?} us; reduction {:.1}%; native/pest {:.2}x", (1.0 - current.0 / before.0) * 100.0, current.0 / pest.0);
    }
}
