use grameo::{
    islands::{Operator, TrailingTrivia},
    pratt::Fixity,
};

pub fn bind(grammar: &mut grameo::CompiledGrammar) -> Result<(), Vec<grameo::Diagnostic>> {
    // Assignment, ternary, equality, casts and postfix syntax remain PEG.
    // Capture parsing continues to use the original grammar and PHP walker.
    grammar.bind_pratt(
        "comparison_expression",
        "boolean_not_expression",
        &[
            Operator {
                rule: "mul_op",
                precedence: 50,
                fixity: Fixity::InfixLeft,
            },
            Operator {
                rule: "additive_op",
                precedence: 40,
                fixity: Fixity::InfixLeft,
            },
            Operator {
                rule: "shift_op",
                precedence: 30,
                fixity: Fixity::InfixLeft,
            },
            Operator {
                rule: "concat_op",
                precedence: 20,
                fixity: Fixity::InfixLeft,
            },
            Operator {
                rule: "cmp_op",
                precedence: 10,
                fixity: Fixity::InfixLeft,
            },
        ],
        TrailingTrivia::Consume,
    )
}

#[cfg(test)]
mod tests {
    #[test]
    fn php_operator_and_heredoc_recognition_matches_peg() {
        std::thread::Builder::new()
            .stack_size(64 * 1024 * 1024)
            .spawn(|| {
                let source = include_str!("../src/grammar.pest");
                let ordinary = grameo::compile(source).unwrap();
                let mut bound = grameo::compile(source).unwrap();
                super::bind(&mut bound).unwrap();
                let check = |input: &str| {
                    let expected = ordinary.recognize("program", input);
                    let actual = bound.recognize("program", input);
                    assert_eq!(
                        actual.as_ref().ok().map(|r| r.consumed),
                        expected.as_ref().ok().map(|r| r.consumed),
                        "{input}"
                    );
                    assert_eq!(actual.is_ok(), expected.is_ok(), "{input}");
                };
                let operators = [
                    "*", "/", "%", "+", "-", "<<", ">>", ".", "<=>", "<=", ">=", "<", ">",
                ];
                for first in operators {
                    for second in operators {
                        check(&format!("<?php $x = $a {first} $b {second} $c;"));
                        check(&format!(
                            "<?php $x = $a /* trivia */ {first} ($b {second} $c);"
                        ));
                        check(&format!("<?php $x = $a {first} $b {second};"));
                    }
                }
                for expression in [
                    "false !== $p = strpos($text, 'x')",
                    "!$a instanceof Thing + 2",
                    "(int) $a ** 2 + $b",
                    "-$a ** 2 * 3",
                    "$a++ + --$b",
                    "$a < $b < $c",
                    "$a << $b < $c",
                    "$a . $b + $c",
                    "$a ? $b + $c : $d * $e",
                    "$a ?? $b + $c",
                    "new Thing($a + 1) * $b",
                    "clone $a + $b",
                    "$a->f()[0] + $b",
                    "$a +",
                    "$a <<",
                    "$a <=>",
                    "$a **",
                    "$a + /* incomplete */",
                    "<<<TEXT\nhello\nTEXT\n . 'world'",
                    "<<<'TEXT'\nhello\nTEXT\n + 2",
                ] {
                    check(&format!("<?php $x = {expression};"));
                }
                let mut large = String::from("<?php\n");
                for i in 0..2000 {
                    large.push_str(&format!("$v{i} = ($a + $b * 3) << 2 . $c <= $d;\n"));
                }
                check(&large);
            })
            .unwrap()
            .join()
            .unwrap();
    }
}
