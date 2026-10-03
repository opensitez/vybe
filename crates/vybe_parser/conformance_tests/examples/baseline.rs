//! Local release baseline, no timing assertions in correctness tests.
//! Run with --release; comparisons include capture construction, not AST walks.
use pest::Parser;
use pest_derive::Parser;
use std::{hint::black_box, time::Instant};
use vybe_parser_generated_tests::lua;

#[derive(Parser)]
#[grammar = "../../../languages/lua/src/grammar.pest"]
struct ReferenceParser;

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
    }
}
