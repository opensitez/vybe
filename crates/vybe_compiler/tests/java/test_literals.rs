use crate::helpers::run_main;

#[test]
fn hex_literal_zero_ff() {
    let out = run_main("System.out.println(0xFF);");
    assert_eq!(out, vec!["255"]);
}

#[test]
fn binary_literal_ten() {
    let out = run_main("System.out.println(0b1010);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn double_literal_fraction() {
    let out = run_main("double pi = 3.14; System.out.println(pi);");
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn char_literal_prints_code_unit() {
    let out = run_main("char c = 'A'; System.out.println(c);");
    assert_eq!(out, vec!["A"]);
}

#[test]
fn long_literal_suffix() {
    let out = run_main("long n = 1_000_000L; System.out.println(n);");
    assert_eq!(out, vec!["1000000"]);
}

#[test]
fn octal_literal_eight() {
    let out = run_main("System.out.println(010);");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn float_literal_with_f_suffix() {
    let out = run_main("float x = 2.5f; System.out.println(x);");
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn string_escape_sequence_newline() {
    let out = run_main(r#"String s = "a\nb"; System.out.println(s.length());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn boolean_literal_in_expression() {
    let out = run_main("System.out.println(!false);");
    assert_eq!(out, vec!["true"]);
}
