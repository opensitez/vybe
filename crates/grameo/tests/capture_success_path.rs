use grameo::{MatchOptions, ParseTree, compile};
fn shape(tree: &ParseTree<'_, '_>) -> Vec<(String, usize, usize, usize)> {
    let mut pending: Vec<_> = tree.pairs().map(|p| (p, 0)).collect();
    pending.reverse();
    let mut out = Vec::new();
    while let Some((p, d)) = pending.pop() {
        let span = p.as_span();
        out.push((p.rule_name().to_owned(), span.start, span.end, d));
        let children: Vec<_> = p.into_inner().map(|p| (p, d + 1)).collect();
        pending.extend(children.into_iter().rev());
    }
    out
}
#[test]
fn successful_capture_specialization_preserves_trees_and_exact_failures() {
    let grammars = [
        r#"WHITESPACE = _{ " " | "\t" }
entry = { SOI ~ (a | b)* ~ EOI }
a = { "alpha" ~ ASCII_DIGIT+ }
b = { "beta" ~ ("(" ~ a ~ ")")? }"#,
        r#"entry = { SOI ~ (PUSH("a"+) ~ "b" ~ POP)+ ~ EOI }"#,
        r#"entry = { SOI ~ (!"bad" ~ ANY)* ~ EOI }"#,
    ];
    let inputs = [
        "",
        "alpha123 beta(alpha4) beta",
        "alpha123 bad",
        "aaabaaa",
        "aabaaa",
        "日本語 é\r\n",
        "bad",
    ];
    for src in grammars {
        let grammar = compile(src).unwrap();
        for input in inputs {
            let fast = grammar.parse_source("entry", input);
            let regular = grammar.parse_source_with_options(
                "entry",
                input,
                MatchOptions {
                    max_steps: usize::MAX - 1,
                    max_rule_depth: 4096,
                },
            );
            match (fast, regular) {
                (Ok(a), Ok(b)) => {
                    assert_eq!(a.consumed(), b.consumed());
                    assert_eq!(shape(&a), shape(&b));
                }
                (Err(a), Err(b)) => assert_eq!(a, b),
                _ => panic!("acceptance mismatch: {input}"),
            }
        }
    }
}
