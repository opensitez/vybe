fn main() {
    let grammar = grameo::compile(r#"
        WHITESPACE = _{ " " | "\t" | "\n" }
        number = @{ ASCII_DIGIT+ }
        program = { SOI ~ number ~ ("," ~ number)* ~ EOI }
    "#).expect("valid grammar");

    let tree = grammar.parse_source("program", "12, 34").expect("valid input");
    for root in tree.pairs() {
        for pair in root.into_inner() {
            if pair.rule_name() == "number" {
                println!("{}", pair.as_str()); // Prints 12, then 34.
            }
        }
    }

    assert!(grammar.recognize("program", "12, 34").is_ok());
    assert!(grammar.recognize("program", "12, nope").is_err());
}
