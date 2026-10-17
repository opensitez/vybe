//! Scope keywords share the inline word-boundary contract. A prefixed identifier
//! must remain an expression rather than opening a different variable's scope.
use grameo::{CompiledGrammar, ParseTree, Span};

fn captures(tree: &ParseTree<'_, '_>) -> Vec<(String, Span, usize)> {
    fn visit(pair: grameo::Pair<'_, '_, '_>, depth: usize, out: &mut Vec<(String, Span, usize)>) {
        out.push((pair.rule_name().into(), pair.as_span(), depth));
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
fn grammars() -> (CompiledGrammar, CompiledGrammar, CompiledGrammar) {
    let source = include_str!("../src/grammar.pest");
    let legacy = source
        .replace(
            "global_stmt = { global_kw ~",
            "global_stmt = { \"global\" ~",
        )
        .replace(
            "nonlocal_stmt = { nonlocal_kw ~",
            "nonlocal_stmt = { \"nonlocal\" ~",
        );
    let expression_only = source
        .replace(
            "global_stmt = { global_kw ~",
            "global_stmt = { !ANY ~ global_kw ~",
        )
        .replace(
            "nonlocal_stmt = { nonlocal_kw ~",
            "nonlocal_stmt = { !ANY ~ nonlocal_kw ~",
        );
    (
        grameo::compile(&source).unwrap(),
        grameo::compile(&legacy).unwrap(),
        grameo::compile(&expression_only).unwrap(),
    )
}
#[test]
fn prefixed_identifiers_have_exact_ordinary_expression_captures() {
    let (production, _, expression_only) = grammars();
    for prefix in ["global", "nonlocal"] {
        for suffix in ["x", "_", "_value", "Counter", "1", "_123"] {
            for tail in ["\n", ";\n", " = 1\n", "(1)\n", " + 2\n"] {
                let input = format!("{prefix}{suffix}{tail}");
                let tree = production.parse_source("program", &input).unwrap();
                let expected = expression_only.parse_source("program", &input).unwrap();
                assert_eq!(captures(&tree), captures(&expected), "{input:?}");
                assert!(
                    captures(&tree)
                        .iter()
                        .all(|(name, _, _)| name != "global_stmt" && name != "nonlocal_stmt")
                );
            }
        }
    }
}
#[test]
fn legitimate_declarations_keep_name_captures_and_spans() {
    let (production, legacy, _) = grammars();
    for keyword in ["global", "nonlocal"] {
        for gap in [" ", "  ", "\t"] {
            for names in ["x", "x, y", "_first, second_2"] {
                for end in ["\n", ";\n", ""] {
                    let input = format!("{keyword}{gap}{names}{end}");
                    let a = production.parse_source("program", &input).unwrap();
                    let b = legacy.parse_source("program", &input).unwrap();
                    let all = captures(&a);
                    let filtered: Vec<_> = all
                        .iter()
                        .filter(|(name, _, _)| name != "global_kw" && name != "nonlocal_kw")
                        .cloned()
                        .collect();
                    assert_eq!(filtered, captures(&b), "{input:?}");
                    let keyword_capture = all
                        .iter()
                        .find(|(name, _, _)| name == &format!("{keyword}_kw"))
                        .unwrap();
                    assert_eq!(
                        keyword_capture.1,
                        Span {
                            start: 0,
                            end: keyword.len()
                        }
                    );
                }
            }
        }
    }
}
