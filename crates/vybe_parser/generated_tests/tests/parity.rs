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

#[test]
fn generated_lua_pratt_recognition_preserves_operator_ambiguity_and_partial_input() {
    use vybe_parser_generated_tests::lua_pratt;
    let operators = [
        "+", "-", "*", "/", "//", "%", "^", "..", "<<", ">>", "&", "~", "|", "<=", ">=", "~=",
        "==", "<", ">", "and", "or",
    ];
    for a in operators {
        for b in operators {
            for suffix in ["c", "", "-", "not", "(c)"] {
                let input = format!("a {a} b {b} {suffix}");
                for entry in ["expr", "chunk"] {
                    let plain =
                        recognize_program(&lua::Parser, entry, &input, MatchOptions::default())
                            .map(|r| r.consumed)
                            .map_err(|e| (e.kind, e.offset));
                    let pratt = recognize_program(
                        &lua_pratt::Parser,
                        entry,
                        &input,
                        MatchOptions::default(),
                    )
                    .map(|r| r.consumed)
                    .map_err(|e| (e.kind, e.offset));
                    assert_eq!(plain, pratt, "{entry}: {input}");
                }
            }
        }
    }
    // The compatibility/capture path must still expose every original wrapper.
    let input = "return -a^2 + b ~= c and f(3)";
    let a = lua::Parser.parse(lua::Rule::chunk, input).unwrap();
    let b = lua_pratt::Parser
        .parse(lua_pratt::Rule::chunk, input)
        .unwrap();
    assert_eq!(shape(&a), shape(&b));
}
