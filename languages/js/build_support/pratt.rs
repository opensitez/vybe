use grameo::{
    islands::{Operator, TrailingTrivia},
    pratt::Fixity,
};

pub fn bind(grammar: &mut grameo::CompiledGrammar) -> Result<(), Vec<grameo::Diagnostic>> {
    // Ordinary capture parsing remains PEG. The registered JavaScript hybrid
    // frontend explicitly selects arithmetic islands; other owned regions use
    // grammar hooks, with contextual captures for unsupported syntax.
    grammar.bind_pratt(
        "additive",
        "unary",
        &[
            Operator {
                rule: "additive_op",
                precedence: 10,
                fixity: Fixity::InfixLeft,
            },
            Operator {
                rule: "multiplicative_op",
                precedence: 20,
                fixity: Fixity::InfixLeft,
            },
            Operator {
                rule: "exp_op",
                precedence: 30,
                fixity: Fixity::InfixRight,
            },
        ],
        TrailingTrivia::Consume,
    )
}
