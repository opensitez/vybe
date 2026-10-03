//! Local release baseline, no timing assertions in correctness tests.
//! Run with --release; comparisons include capture construction, not AST walks.
use pest::Parser;
use pest_derive::Parser;
use std::{hint::black_box, time::Instant};
use vybe_parser_generated_tests::lua;

#[derive(Parser)]
#[grammar = "../../../languages/lua/src/grammar.pest"]
struct ReferenceParser;

#[derive(Debug)]
struct WithoutTriviaCertificate<P>(P);
impl<P: vybe_parser::program::Program> vybe_parser::program::Program
    for WithoutTriviaCertificate<P>
{
    fn rule_id(&self, name: &str) -> Option<usize> {
        self.0.rule_id(name)
    }
    fn rule(&self, id: usize) -> vybe_parser::program::RuleSpec<'_> {
        self.0.rule(id)
    }
    fn instruction(&self, id: usize) -> vybe_parser::program::Instruction<'_> {
        self.0.instruction(id)
    }
    fn scope(&self, id: usize) -> vybe_parser::program::Scope {
        self.0.scope(id)
    }
    fn fast_ascii_repetitions(&self) -> bool {
        self.0.fast_ascii_repetitions()
    }
    fn source_pratt(&self, rule: usize) -> Option<vybe_parser::program::SourcePratt<'_>> {
        self.0.source_pratt(rule)
    }
}

fn measure(mut action: impl FnMut(), count: usize) -> (f64, f64) {
    for _ in 0..10 {
        action();
    }
    let mut times = Vec::with_capacity(count);
    for _ in 0..count {
        let now = Instant::now();
        action();
        times.push(now.elapsed().as_secs_f64() * 1e6);
    }
    times.sort_by(f64::total_cmp);
    (times[count / 2], times[(count * 95 / 100).min(count - 1)])
}

fn main() {
    let grammar_source = include_str!("../../../../languages/lua/src/grammar.pest");
    let grammar = vybe_parser::compile(grammar_source).unwrap();
    let count = 101;
    ascii_baseline(count);
    println!(
        "Lua baseline: 10 warmups, {count} samples, median/p95 µs, rustc={}, target={}",
        option_env!("RUSTC_VERSION").unwrap_or("record with rustc -V"),
        std::env::consts::ARCH
    );
    println!(
        "grammar compilation: {:?}",
        measure(
            || {
                black_box(vybe_parser::compile(black_box(grammar_source)).unwrap());
            },
            count
        )
    );
    for (label, source) in [
        ("tiny", "return 1 + 2 * 3".to_owned()),
        (
            "declarations100",
            (0..100).map(|n| format!("local x{n} = {n}\n")).collect(),
        ),
        (
            "expressions100",
            (0..100)
                .map(|n| format!("local x{n} = f(1 + 2 * 3, (4 - 5) / 6)\n"))
                .collect(),
        ),
    ] {
        let reference = ReferenceParser::parse(Rule::chunk, &source).unwrap();
        let dynamic = grammar.parse_source("chunk", &source).unwrap();
        let generated = lua::Parser.parse(lua::Rule::chunk, &source).unwrap();
        assert_eq!(dynamic.consumed(), source.len());
        assert_eq!(generated.consumed(), source.len());
        assert_eq!(reference.clone().flatten().count(), dynamic.capture_count());
        assert_eq!(dynamic.capture_count(), generated.capture_count());
        println!(
            "{label}: {} bytes, {} captures",
            source.len(),
            dynamic.capture_count()
        );
        println!(
            "  reference capture {:?}",
            measure(
                || {
                    black_box(ReferenceParser::parse(Rule::chunk, black_box(&source)).unwrap());
                },
                count
            )
        );
        println!(
            "  resolved capture  {:?}",
            measure(
                || {
                    black_box(grammar.parse_source("chunk", black_box(&source)).unwrap());
                },
                count
            )
        );
        println!(
            "  static capture    {:?}",
            measure(
                || {
                    black_box(
                        lua::Parser
                            .parse(lua::Rule::chunk, black_box(&source))
                            .unwrap(),
                    );
                },
                count
            )
        );
        println!(
            "  static recognize  {:?}",
            measure(
                || {
                    black_box(
                        lua::Parser
                            .recognize(lua::Rule::chunk, black_box(&source))
                            .unwrap(),
                    );
                },
                count
            )
        );
        let baseline = WithoutTriviaCertificate(lua::Parser);
        let options = vybe_parser::MatchOptions::default();
        assert_eq!(
            lua::Parser.recognize(lua::Rule::chunk, &source).unwrap(),
            vybe_parser::engine::recognize_program(&baseline, "chunk", &source, options).unwrap()
        );
        println!(
            "  without trivia certificate {:?}",
            measure(
                || {
                    black_box(
                        vybe_parser::engine::recognize_program(
                            &baseline,
                            "chunk",
                            black_box(&source),
                            options,
                        )
                        .unwrap(),
                    );
                },
                count
            )
        );
        let pratt = vybe_parser_generated_tests::lua_pratt::Parser;
        let entry = vybe_parser_generated_tests::lua_pratt::Rule::chunk;
        assert_eq!(
            pratt.recognize(entry, &source).unwrap().consumed,
            source.len()
        );
        println!(
            "  Pratt recognize   {:?}",
            measure(
                || {
                    black_box(pratt.recognize(entry, black_box(&source)).unwrap());
                },
                count
            )
        );
    }
}

fn ascii_baseline(count: usize) {
    use vybe_parser::{
        MatchOptions,
        engine::recognize_program,
        program::{Instruction, Program, RuleSpec, Scope},
    };
    #[derive(Debug)]
    struct Baseline<'a>(&'a vybe_parser::CompiledGrammar);
    impl Program for Baseline<'_> {
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
    let grammar = vybe_parser::compile("root = @{ ASCII_DIGIT+ ~ EOI }").unwrap();
    let input = "1234567890".repeat(409);
    let ordinary = Baseline(&grammar);
    let options = MatchOptions::default();
    assert_eq!(
        recognize_program(&grammar, "root", &input, options),
        recognize_program(&ordinary, "root", &input, options)
    );
    println!(
        "ASCII digit repetition: {} bytes, 10 warmups, {count} samples, median/p95 µs",
        input.len()
    );
    println!(
        "  continuation {:?}",
        measure(
            || {
                black_box(
                    recognize_program(&ordinary, "root", black_box(&input), options).unwrap(),
                );
            },
            count
        )
    );
    println!(
        "  byte scanner {:?}",
        measure(
            || {
                black_box(recognize_program(&grammar, "root", black_box(&input), options).unwrap());
            },
            count
        )
    );
}
