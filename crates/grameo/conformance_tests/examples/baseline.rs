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
#[derive(Debug)]
struct ScheduledTrivia<P>(P);
impl<P: grameo::program::Program> grameo::program::Program for ScheduledTrivia<P> {
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

#[derive(Debug)]
struct BeforeRuleDispatch<P>(P);
impl<P: grameo::program::Program> grameo::program::Program for BeforeRuleDispatch<P> {
    fn tail_continuations(&self) -> bool {
        false
    }
    fn fast_rule_failures(&self) -> bool {
        false
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

#[derive(Debug)]
struct BeforeAtomic<P>(P);
impl<P: grameo::program::Program> grameo::program::Program for BeforeAtomic<P> {
    fn failure_terminals(&self, id: usize) -> &[usize] {
        self.0.failure_terminals(id)
    }
    fn terminal_choice(&self, id: usize) -> Option<grameo::lexical::TerminalChoice<'_>> {
        self.0.terminal_choice(id)
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

fn paired_capture(source: &str, count: usize) -> ((f64, f64), (f64, f64)) {
    use grameo::engine::parse_program;
    let scheduled = ScheduledTrivia(lua::Parser);
    let options = grameo::MatchOptions::default();
    let mut before = Vec::with_capacity(count);
    let mut after = Vec::with_capacity(count);
    let mut sample = |fast: bool, record: bool| {
        let now = Instant::now();
        if fast {
            black_box(parse_program(&lua::Parser, "chunk", black_box(source), options).unwrap());
        } else {
            black_box(parse_program(&scheduled, "chunk", black_box(source), options).unwrap());
        }
        if record {
            let elapsed = now.elapsed().as_secs_f64() * 1e6;
            if fast {
                after.push(elapsed);
            } else {
                before.push(elapsed);
            }
        }
    };
    for _ in 0..10 {
        sample(false, false);
        sample(true, false);
    }
    for index in 0..count {
        sample(index % 2 != 0, true);
        sample(index % 2 == 0, true);
    }
    before.sort_by(f64::total_cmp);
    after.sort_by(f64::total_cmp);
    (
        (before[count / 2], before[count * 95 / 100]),
        (after[count / 2], after[count * 95 / 100]),
    )
}
fn paired_rules<P: grameo::program::Program>(
    source: &str,
    count: usize,
    before: P,
) -> [(f64, f64); 3] {
    use grameo::engine::parse_program;
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
                            "chunk",
                            black_box(source),
                            grameo::MatchOptions::default(),
                        )
                        .unwrap(),
                    );
                }
                1 => {
                    black_box(
                        lua::Parser
                            .parse(lua::Rule::chunk, black_box(source))
                            .unwrap(),
                    );
                }
                _ => {
                    black_box(ReferenceParser::parse(Rule::chunk, black_box(source)).unwrap());
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

impl<P: grameo::program::Program> grameo::program::Program
    for WithoutTriviaCertificate<P>
{
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
    let grammar = grameo::compile(grammar_source).unwrap();
    let count = std::env::args()
        .skip(1)
        .find_map(|arg| arg.parse::<usize>().ok())
        .unwrap_or(101)
        .max(1);
    let bulk = std::env::args().any(|arg| arg == "--bulk-dispatch-only");
    let rule_dispatch_only =
        bulk || std::env::args().any(|arg| arg == "--rule-dispatch-only");
    let dispatch_only = std::env::args().any(|arg| arg == "--dispatch-only");
    if !dispatch_only && !rule_dispatch_only {
        ascii_baseline(count);
    }
    println!(
        "Lua baseline: 10 warmups, {count} samples, median/p95 µs, rustc={}, target={}",
        option_env!("RUSTC_VERSION").unwrap_or("record with rustc -V"),
        std::env::consts::ARCH
    );
    if !dispatch_only && !rule_dispatch_only {
        println!(
            "grammar compilation: {:?}",
            measure(
                || {
                    black_box(grameo::compile(black_box(grammar_source)).unwrap());
                },
                count
            )
        );
    }
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
        (
            "comments100",
            (0..100)
                .map(|n| format!("-- comment {n}\nlocal x{n} = 'é😀'\n"))
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
        if rule_dispatch_only {
            let [before, after, reference] = if bulk {
                paired_rules(&source, count, BeforeAtomic(lua::Parser))
            } else {
                paired_rules(&source, count, BeforeRuleDispatch(lua::Parser))
            };
            println!("  paired dispatch before {before:?}; after {after:?}; pest {reference:?}");
            continue;
        }
        let (scheduled, fused) = paired_capture(&source, count);
        println!("  paired dispatch scheduled {scheduled:?}; optimized {fused:?}");
        if dispatch_only {
            continue;
        }
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
        let options = grameo::MatchOptions::default();
        assert_eq!(
            lua::Parser.recognize(lua::Rule::chunk, &source).unwrap(),
            grameo::engine::recognize_program(&baseline, "chunk", &source, options).unwrap()
        );
        println!(
            "  without trivia certificate {:?}",
            measure(
                || {
                    black_box(
                        grameo::engine::recognize_program(
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
    use grameo::{
        MatchOptions,
        engine::recognize_program,
        program::{Instruction, Program, RuleSpec, Scope},
    };
    #[derive(Debug)]
    struct Baseline<'a>(&'a grameo::CompiledGrammar);
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
    let grammar = grameo::compile("root = @{ ASCII_DIGIT+ ~ EOI }").unwrap();
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
