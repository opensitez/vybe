use pest::Parser;
use pest_derive::Parser;
use vybe_parser::{Pair, compile};

#[derive(Parser)]
#[grammar = "tests/fixtures.pest"]
struct ReferenceParser;

type Shape = Vec<(String, usize, usize, usize)>;

#[test]
fn migration_pair_positions_match_reference_unicode_and_line_boundaries() {
    #[derive(Clone, Copy)]
    enum Rules {
        Program,
        Cell,
        Eoi,
    }
    impl vybe_parser::compat::RuleIdentity for Rules {
        fn from_capture(rule: vybe_parser::CaptureRule) -> Option<Self> {
            match rule {
                vybe_parser::CaptureRule::User(0) => Some(Self::Program),
                vybe_parser::CaptureRule::User(1) => Some(Self::Cell),
                vybe_parser::CaptureRule::EndOfInput => Some(Self::Eoi),
                _ => None,
            }
        }
    }
    let grammar = compile("program = { SOI ~ cell* ~ EOI } cell = { ANY }").unwrap();
    for input in ["", "aé😀", "a\rb", "a\r\nb", "\n\r\n\r😀\n"] {
        let root = grammar
            .parse_source("program", input)
            .unwrap()
            .into_typed_pairs::<Rules>()
            .next()
            .unwrap();
        for pair in root.into_inner() {
            let span = pair.as_span();
            assert_eq!(
                span.start_pos().line_col(),
                pest::Position::new(input, span.start()).unwrap().line_col(),
                "start {:?} {}",
                input,
                span.start()
            );
            assert_eq!(
                span.end_pos().line_col(),
                pest::Position::new(input, span.end()).unwrap().line_col(),
                "end {:?} {}",
                input,
                span.end()
            );
        }
    }
}
fn pest_shape<R: pest::RuleType + std::fmt::Debug>(pairs: pest::iterators::Pairs<'_, R>) -> Shape {
    fn visit<R: pest::RuleType + std::fmt::Debug>(
        pair: pest::iterators::Pair<'_, R>,
        depth: usize,
        out: &mut Shape,
    ) {
        out.push((
            format!("{:?}", pair.as_rule()),
            pair.as_span().start(),
            pair.as_span().end(),
            depth,
        ));
        for child in pair.into_inner() {
            visit(child, depth + 1, out);
        }
    }
    let mut out = Vec::new();
    for pair in pairs {
        visit(pair, 0, &mut out);
    }
    out
}
fn vybe_shape(tree: &vybe_parser::ParseTree<'_, '_>) -> Shape {
    fn visit(pair: Pair<'_, '_, '_>, depth: usize, out: &mut Shape) {
        let span = pair.as_span();
        out.push((pair.rule_name().to_owned(), span.start, span.end, depth));
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

#[test]
fn synthetic_mode_stack_whitespace_and_repetition_contracts_match_pest() {
    let grammar = compile(include_str!("fixtures.pest")).unwrap();
    let entries = [
        Rule::main,
        Rule::single,
        Rule::sequence,
        Rule::repeat,
        Rule::bounded,
        Rule::optional,
        Rule::peek,
        Rule::rollback,
        Rule::atomic,
        Rule::compound,
        Rule::nonatomic,
        Rule::outer_atomic,
        Rule::stack,
        Rule::predicate_stack,
        Rule::push_stack,
        Rule::repeat_stack,
        Rule::slice,
        Rule::literal,
        Rule::newline,
    ];
    let mut inputs = vec![
        String::new(),
        "ababbaba".into(),
        "ababb bababa".into(),
        "IFα😀".into(),
        "IFÉ😀".into(),
        "a\r\nb".into(),
        "a\rb".into(),
        "a\nb".into(),
        " a // comment\na ".into(),
    ];
    for len in 1..=4 {
        for mut number in 0..4usize.pow(len) {
            let mut input = String::new();
            for _ in 0..len {
                input.push(['a', 'b', ' ', 'x'][number % 4]);
                number /= 4;
            }
            inputs.push(input);
        }
    }
    let mut comparisons = 0;
    for entry in entries {
        let name = format!("{entry:?}");
        for source in &inputs {
            let reference = std::panic::catch_unwind(|| ReferenceParser::parse(entry, source));
            let ours = grammar.parse_source(&name, source);
            match (reference, ours) {
                (Ok(Ok(reference)), Ok(ours)) => assert_eq!(
                    pest_shape(reference),
                    vybe_shape(&ours),
                    "{name} {source:?}"
                ),
                (Ok(Err(_)), Err(ours)) => assert_eq!(
                    ours.kind,
                    vybe_parser::ParseErrorKind::Syntax,
                    "{name} {source:?}: {ours}"
                ),
                // Malformed guest input can exhaust a grammar stack. Pest
                // panics; Vybe reports an explicit grammar-state error.
                (Err(_), Err(ours)) => assert_eq!(
                    ours.kind,
                    vybe_parser::ParseErrorKind::InvalidStack,
                    "{name} {source:?}: {ours}"
                ),
                (reference, ours) => panic!("{name} {source:?}: pest={reference:?}, vybe={ours:?}"),
            }
            comparisons += 1;
        }
    }
    assert!(comparisons > 6000);
}

mod lua {
    use super::*;
    #[derive(Parser)]
    #[grammar = "../../../languages/lua/src/grammar.pest"]
    struct Lua;
    #[test]
    fn generated_pratt_expression_island_matches_black_box_reference() {
        use vybe_parser_generated_tests::lua_pratt;
        let operators = [
            "+", "-", "*", "/", "//", "%", "^", "..", "<<", ">>", "&", "~", "|", "<=", ">=", "~=",
            "==", "<", ">", "and", "or",
        ];
        for a in operators {
            for b in operators {
                for suffix in ["c", "", "-", "not", "(c)"] {
                    let expression = format!("a {a} b {b} {suffix}");
                    let chunk = format!("return {expression}");
                    for (reference_rule, generated_rule, source) in [
                        (Rule::expr, lua_pratt::Rule::expr, expression.as_str()),
                        (Rule::chunk, lua_pratt::Rule::chunk, chunk.as_str()),
                    ] {
                        match (
                            Lua::parse(reference_rule, source),
                            lua_pratt::Parser.recognize(generated_rule, source),
                        ) {
                            (Ok(mut reference), Ok(ours)) => assert_eq!(
                                reference.next().unwrap().as_span().end(),
                                ours.consumed,
                                "{source}"
                            ),
                            (Err(reference), Err(ours)) => {
                                assert_eq!(ours.kind, vybe_parser::ParseErrorKind::Syntax);
                                let offset = match reference.location {
                                    pest::error::InputLocation::Pos(offset)
                                    | pest::error::InputLocation::Span((offset, _)) => offset,
                                };
                                assert_eq!(offset, ours.offset, "{source}");
                            }
                            (reference, ours) => {
                                panic!("{source}: reference={reference:?}; ours={ours:?}")
                            }
                        }
                    }
                }
            }
        }
    }
    #[test]
    fn actual_lua_grammar_source_trees_match() {
        let grammar = compile(include_str!("../../../../languages/lua/src/grammar.pest")).unwrap();
        for source in [
            "",
            "local x = 1 + 2 * 3\nprint(x)",
            "function f(a) return a+1 end\nprint(f(2))",
            "for i=1,3 do print(i) end",
            "local t = {x=1, 2}\nreturn t.x",
            "if true then print('x') else print('y') end",
            "local s = [[long string]]",
            "local x =",
            "function(",
            "local x=0b10",
        ] {
            match (
                Lua::parse(Rule::chunk, source),
                grammar.parse_source("chunk", source),
            ) {
                (Ok(reference), Ok(ours)) => {
                    assert_eq!(pest_shape(reference), vybe_shape(&ours), "{source:?}")
                }
                (Err(_), Err(ours)) => assert_eq!(ours.kind, vybe_parser::ParseErrorKind::Syntax),
                (reference, ours) => panic!("{source:?}: pest={reference:?}, vybe={ours:?}"),
            }
        }
    }
}
