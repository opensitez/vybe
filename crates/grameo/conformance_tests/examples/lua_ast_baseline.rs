//! Raw expression/module AST baseline; run with --release. Includes source indexing,
//! parsing, AST construction and destruction. No performance assertions.
#[path = "../../generated_tests/tests/support/ast_compare.rs"]
#[allow(dead_code)]
mod compare;
#[path = "../../generated_tests/tests/support/lua_ast.rs"]
mod direct;
use pest_derive::Parser;
use std::{hint::black_box, time::Instant};
use grameo::{MatchOptions, engine::build_program};
use vybe_parser_generated_tests::lua::Parser;
#[derive(Parser)]
#[grammar = "../../../languages/lua/src/grammar.pest"]
struct LuaParser;
#[allow(
    clippy::collapsible_if,
    clippy::while_let_on_iterator,
    clippy::clone_on_copy,
    clippy::manual_strip,
    dead_code
)]
mod walker {
    include!(concat!(env!("OUT_DIR"), "/lua_raw_walker.rs"));
}
#[allow(dead_code)]
mod normalize {
    include!(concat!(env!("OUT_DIR"), "/lua_numeric_helpers.rs"));
}
mod native {
    use vybe_parser_generated_tests::lua::{Parser, Rule};
    struct LuaParser;
    impl LuaParser {
        fn parse(
            rule: Rule,
            source: &str,
        ) -> Result<grameo::compat::Pairs<'_, Rule>, grameo::ParseError> {
            Parser.parse_pairs(rule, source)
        }
    }
    #[allow(
        clippy::collapsible_if,
        clippy::while_let_on_iterator,
        clippy::clone_on_copy,
        clippy::manual_strip,
        dead_code
    )]
    pub mod walker {
        include!(concat!(env!("OUT_DIR"), "/lua_native_walker.rs"));
    }
    #[allow(dead_code)]
    mod normalize {
        include!(concat!(env!("OUT_DIR"), "/lua_numeric_helpers.rs"));
    }
}
fn parse(source: &str) -> vybe_ast::Expression {
    let mut builder = direct::Ast::new(source);
    debug_assert!(builder.is_empty());
    let matched = build_program(
        &Parser,
        "expr",
        source,
        MatchOptions::default(),
        &mut builder,
    )
    .unwrap();
    assert_eq!(matched.consumed, source.len());
    builder.into_expression()
}
#[derive(Debug)]
struct ScheduledChoices<P>(P);
impl<P: grameo::program::Program> grameo::program::Program for ScheduledChoices<P> {
    fn rule_id(&self, name: &str) -> Option<usize> {
        self.0.rule_id(name)
    }
    fn rule(&self, rule: usize) -> grameo::program::RuleSpec<'_> {
        self.0.rule(rule)
    }
    fn instruction(&self, expression: usize) -> grameo::program::Instruction<'_> {
        self.0.instruction(expression)
    }
    fn scope(&self, scope: usize) -> grameo::program::Scope {
        self.0.scope(scope)
    }
    fn source_pratt(&self, rule: usize) -> Option<grameo::program::SourcePratt<'_>> {
        self.0.source_pratt(rule)
    }
    fn fast_ascii_repetitions(&self) -> bool {
        self.0.fast_ascii_repetitions()
    }
    fn trivia_failure_prefix(&self, rule: usize) -> Option<grameo::lexical::FailurePrefix> {
        self.0.trivia_failure_prefix(rule)
    }
}
fn pratt_builder<'a, P: grameo::program::Program>(
    program: &P,
    entry: &str,
    source: &'a str,
) -> direct::Ast<'a> {
    let mut builder = direct::Ast::with_pratt(source);
    let matched = build_program(
        program,
        entry,
        source,
        MatchOptions::default(),
        &mut builder,
    )
    .unwrap();
    assert_eq!(matched.consumed, source.len());
    builder
}
fn parse_pratt(source: &str) -> vybe_ast::Expression {
    pratt_builder(
        &vybe_parser_generated_tests::lua_pratt::Parser,
        "expr",
        source,
    )
    .into_expression()
}
fn parse_pratt_module(source: &str) -> vybe_ast::Module {
    pratt_builder(
        &vybe_parser_generated_tests::lua_pratt::Parser,
        "chunk",
        source,
    )
    .into_module()
}
fn measure_pair(
    mut scheduled: impl FnMut(),
    mut certified: impl FnMut(),
) -> ((f64, f64), (f64, f64)) {
    for _ in 0..10 {
        scheduled();
        certified();
    }
    let mut before = Vec::with_capacity(101);
    let mut after = Vec::with_capacity(101);
    fn sample(action: &mut impl FnMut(), times: &mut Vec<f64>) {
        let start = Instant::now();
        action();
        times.push(start.elapsed().as_secs_f64() * 1e6);
    }
    for index in 0..101 {
        if index % 2 == 0 {
            sample(&mut scheduled, &mut before);
            sample(&mut certified, &mut after);
        } else {
            sample(&mut certified, &mut after);
            sample(&mut scheduled, &mut before);
        }
    }
    before.sort_by(f64::total_cmp);
    after.sort_by(f64::total_cmp);
    ((before[50], before[95]), (after[50], after[95]))
}
fn parse_module(source: &str) -> vybe_ast::Module {
    let mut builder = direct::Ast::new(source);
    let matched = build_program(
        &Parser,
        "chunk",
        source,
        MatchOptions::default(),
        &mut builder,
    )
    .unwrap();
    assert_eq!(matched.consumed, source.len());
    builder.into_module()
}
fn measure(mut action: impl FnMut()) -> (f64, f64) {
    for _ in 0..10 {
        action();
    }
    let mut times = Vec::with_capacity(101);
    for _ in 0..101 {
        let now = Instant::now();
        action();
        times.push(now.elapsed().as_secs_f64() * 1e6);
    }
    times.sort_by(f64::total_cmp);
    (times[50], times[95])
}
fn main() {
    println!(
        "{}; {}; 10 warmups, 101 samples; median/p95 µs",
        env!("RUSTC_VERSION"),
        std::env::consts::ARCH
    );
    for (name, source) in [
        ("tiny", "1 + 2 * 3".into()),
        (
            "operators100",
            std::iter::repeat_n("a * 2", 100)
                .collect::<Vec<_>>()
                .join(" + "),
        ),
        (
            "calls30",
            std::iter::repeat_n("obj:m(f(1,2)).field[3]", 30)
                .collect::<Vec<_>>()
                .join(" + "),
        ),
        (
            "fields200",
            format!(
                "{{{}}}",
                (0..100)
                    .map(|index| format!("key{index}=f({index}), [g({index})]={index}"))
                    .collect::<Vec<_>>()
                    .join(",")
            ),
        ),
        (
            "functions20",
            format!(
                "{{{}}}",
                (0..20)
                    .map(|index| format!("f{index}=function(a,...) local x=a+1 return x,... end"))
                    .collect::<Vec<_>>()
                    .join(",")
            ),
        ),
    ] {
        let expected = walker::expression_for_conformance(&source).unwrap();
        assert_eq!(expected.1, source.len());
        compare::expression(&parse(&source), &expected.0);
        compare::expression(&parse_pratt(&source), &expected.0);
        let captured = native::walker::expression_for_conformance(&source).unwrap();
        assert_eq!(captured.1, source.len());
        compare::expression(&captured.0, &expected.0);
        let reference = measure(|| {
            black_box(walker::expression_for_conformance(black_box(&source)).unwrap());
        });
        let captured = measure(|| {
            black_box(native::walker::expression_for_conformance(black_box(&source)).unwrap());
        });
        let direct = measure(|| {
            black_box(parse(black_box(&source)));
        });
        let pratt = measure(|| {
            black_box(parse_pratt(black_box(&source)));
        });
        println!(
            "{name}: reference+walker={:.2}/{:.2}; native+walker={:.2}/{:.2}; native direct AST={:.2}/{:.2}; native Pratt AST={:.2}/{:.2}",
            reference.0, reference.1, captured.0, captured.1, direct.0, direct.1, pratt.0, pratt.1
        );
        let scheduled = ScheduledChoices(vybe_parser_generated_tests::lua_pratt::Parser);
        compare::expression(
            &pratt_builder(&scheduled, "expr", &source).into_expression(),
            &expected.0,
        );
        let (scheduled, certified) = measure_pair(
            || {
                black_box(pratt_builder(&scheduled, "expr", black_box(&source)).into_expression());
            },
            || {
                black_box(parse_pratt(black_box(&source)));
            },
        );
        println!(
            "literal choice {name}: scheduled={:.2}/{:.2}; certified={:.2}/{:.2}",
            scheduled.0, scheduled.1, certified.0, certified.1
        );
    }
    for (name, source) in [
        ("module_tiny", "local x=1+2; return x".into()),
        ("module_declarations200", (0..200).map(|n| format!("local x{n}={n}+f(1,2*3)\n")).collect::<String>()),
        ("module_methods20", (0..20).map(|n| format!("function t:m{n}(x,...) for k,v in pairs(x) do if v then return self,k,... end end end\n")).collect::<String>()),
        ("module_numeric20", (0..20).map(|n| format!("for i{n}=1,20,0.5 do t[i{n}]=function() return i{n} end end\n")).collect::<String>()),
    ] {
        let expected = walker::parse(&source).unwrap();
        compare::module(&parse_module(&source), &expected);
        compare::module(&parse_pratt_module(&source), &expected);
        compare::module(&native::walker::parse(&source).unwrap(), &expected);
        let reference = measure(|| { black_box(walker::parse(black_box(&source)).unwrap()); });
        let captured = measure(|| { black_box(native::walker::parse(black_box(&source)).unwrap()); });
        let direct = measure(|| { black_box(parse_module(black_box(&source))); });
        let pratt = measure(|| { black_box(parse_pratt_module(black_box(&source))); });
        println!("{name}: reference+walker={:.2}/{:.2}; native+walker={:.2}/{:.2}; native direct AST={:.2}/{:.2}; native Pratt AST={:.2}/{:.2}", reference.0, reference.1, captured.0, captured.1, direct.0, direct.1, pratt.0, pratt.1);
        let scheduled = ScheduledChoices(vybe_parser_generated_tests::lua_pratt::Parser);
        compare::module(&pratt_builder(&scheduled, "chunk", &source).into_module(), &expected);
        let (scheduled, certified) = measure_pair(
            || { black_box(pratt_builder(&scheduled, "chunk", black_box(&source)).into_module()); },
            || { black_box(parse_pratt_module(black_box(&source))); });
        println!("literal choice {name}: scheduled={:.2}/{:.2}; certified={:.2}/{:.2}", scheduled.0, scheduled.1, certified.0, certified.1);

    }
}
