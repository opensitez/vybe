use vybe_parser::{
    MatchOptions, ParseTree,
    engine::{parse_program, recognize_program},
    program::Program,
};
use vybe_parser_generated_tests::{fixtures, lua};

fn shape(tree: &ParseTree<'_, '_>) -> Vec<(String, usize, usize, usize)> {
    let mut pending: Vec<_> = tree.pairs().map(|pair| (pair, 0)).collect();
    pending.reverse();
    let mut result = Vec::new();
    while let Some((pair, depth)) = pending.pop() {
        let span = pair.as_span();
        result.push((pair.rule_name().to_owned(), span.start, span.end, depth));
        let children: Vec<_> = pair.into_inner().map(|child| (child, depth + 1)).collect();
        pending.extend(children.into_iter().rev());
    }
    result
}

fn compare<P: Program>(source: &str, static_program: &P, inputs: &[&str]) {
    let grammar = vybe_parser::compile(source).unwrap();
    for rule in &grammar.syntax().rules {
        for input in inputs {
            let dynamic = grammar.parse_source(&rule.name, input);
            let generated =
                parse_program(static_program, &rule.name, input, MatchOptions::default());
            match (dynamic, generated) {
                (Ok(a), Ok(b)) => {
                    assert_eq!(a.consumed(), b.consumed());
                    assert_eq!(shape(&a), shape(&b));
                }
                (Err(a), Err(b)) => assert_eq!(a, b, "{} {input:?}", rule.name),
                (a, b) => panic!("{} {input:?}: {a:?} / {b:?}", rule.name),
            }
            assert_eq!(
                grammar.recognize(&rule.name, input),
                recognize_program(static_program, &rule.name, input, MatchOptions::default())
            );
        }
    }
}

#[test]
fn generated_fixture_capture_error_and_recognition_parity() {
    compare(
        include_str!("../../conformance_tests/tests/fixtures.pest"),
        &fixtures::Parser,
        &[
            "", "a", "a b", "ab", "aax", "xx", "é", "\r\n", "aa#hi\nb", "aa", "aaa", "abxy", "abc",
        ],
    );
    assert!(fixtures::Parser.parse(fixtures::Rule::word, "abc").is_ok());
    assert!(
        fixtures::Parser
            .recognize(fixtures::Rule::word, "abc")
            .is_ok()
    );
}

#[test]
fn generated_real_lua_grammar_parity() {
    compare(
        include_str!("../../../../languages/lua/src/grammar.pest"),
        &lua::Parser,
        &[
            "",
            "local x = 1",
            "return 1 + 2 * 3",
            "function f(x) return x end",
            "local =",
            "-- comment\nreturn {a = 1}",
        ],
    );
    let tree = lua::Parser.parse(lua::Rule::chunk, "return 42").unwrap();
    assert_eq!(
        lua::Rule::from_capture(tree.pairs().next().unwrap().as_rule()),
        Some(lua::Rule::chunk)
    );
}
