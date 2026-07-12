use crate::helpers::run_main;

#[test]
fn character_is_digit_accepts_ascii_zero() {
    let out = run_main("System.out.println(Character.isDigit('0'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_digit_accepts_ascii_nine() {
    let out = run_main("System.out.println(Character.isDigit('9'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_digit_rejects_uppercase_letter() {
    let out = run_main("System.out.println(Character.isDigit('A'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_digit_rejects_lowercase_letter() {
    let out = run_main("System.out.println(Character.isDigit('z'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_digit_rejects_space_char() {
    let out = run_main("System.out.println(Character.isDigit(' '));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_digit_rejects_punctuation() {
    let out = run_main("System.out.println(Character.isDigit('-'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_letter_accepts_lowercase_alpha() {
    let out = run_main("System.out.println(Character.isLetter('k'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_letter_accepts_uppercase_alpha() {
    let out = run_main("System.out.println(Character.isLetter('M'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_letter_rejects_digit_char() {
    let out = run_main("System.out.println(Character.isLetter('4'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_letter_rejects_whitespace() {
    let out = run_main("System.out.println(Character.isLetter('\\t'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_letter_rejects_symbol() {
    let out = run_main("System.out.println(Character.isLetter('@'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_whitespace_on_space() {
    let out = run_main("System.out.println(Character.isWhitespace(' '));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_whitespace_on_tab() {
    let out = run_main("System.out.println(Character.isWhitespace('\\t'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_whitespace_on_newline() {
    let out = run_main("System.out.println(Character.isWhitespace('\\n'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_whitespace_rejects_digit() {
    let out = run_main("System.out.println(Character.isWhitespace('7'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_whitespace_rejects_letter() {
    let out = run_main("System.out.println(Character.isWhitespace('a'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_to_upper_case_from_lowercase_a() {
    let out = run_main("System.out.println(Character.toUpperCase('a'));");
    assert_eq!(out, vec!["A"]);
}

#[test]
fn character_to_upper_case_from_lowercase_z() {
    let out = run_main("System.out.println(Character.toUpperCase('z'));");
    assert_eq!(out, vec!["Z"]);
}

#[test]
fn character_to_upper_case_leaves_uppercase_unchanged() {
    let out = run_main("System.out.println(Character.toUpperCase('Q'));");
    assert_eq!(out, vec!["Q"]);
}

#[test]
fn character_to_upper_case_leaves_digit_unchanged() {
    let out = run_main("System.out.println(Character.toUpperCase('3'));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn character_to_lower_case_from_uppercase_z() {
    let out = run_main("System.out.println(Character.toLowerCase('Z'));");
    assert_eq!(out, vec!["z"]);
}

#[test]
fn character_to_lower_case_from_uppercase_b() {
    let out = run_main("System.out.println(Character.toLowerCase('B'));");
    assert_eq!(out, vec!["b"]);
}

#[test]
fn character_to_lower_case_leaves_lowercase_unchanged() {
    let out = run_main("System.out.println(Character.toLowerCase('m'));");
    assert_eq!(out, vec!["m"]);
}

#[test]
fn character_to_lower_case_leaves_digit_unchanged() {
    let out = run_main("System.out.println(Character.toLowerCase('8'));");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn character_for_digit_builds_decimal_five() {
    let out = run_main("System.out.println(Character.forDigit(5, 10));");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn character_for_digit_builds_decimal_zero() {
    let out = run_main("System.out.println(Character.forDigit(0, 10));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn character_for_digit_builds_hex_ten_as_lowercase_a() {
    let out = run_main("System.out.println(Character.forDigit(10, 16));");
    assert_eq!(out, vec!["a"]);
}

#[test]
fn character_for_digit_builds_hex_fifteen_as_lowercase_f() {
    let out = run_main("System.out.println(Character.forDigit(15, 16));");
    assert_eq!(out, vec!["f"]);
}

#[test]
fn character_for_digit_returns_nul_for_negative_digit() {
    let out = run_main("System.out.println((int) Character.forDigit(-1, 10));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn character_for_digit_returns_nul_when_digit_out_of_radix_range() {
    let out = run_main("System.out.println((int) Character.forDigit(10, 10));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn character_digit_reads_decimal_five() {
    let out = run_main("System.out.println(Character.digit('5', 10));");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn character_digit_reads_decimal_zero() {
    let out = run_main("System.out.println(Character.digit('0', 10));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn character_digit_reads_hex_uppercase_a_as_ten() {
    let out = run_main("System.out.println(Character.digit('A', 16));");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn character_digit_reads_hex_lowercase_c_as_twelve() {
    let out = run_main("System.out.println(Character.digit('c', 16));");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn character_digit_returns_minus_one_for_non_digit_in_radix() {
    let out = run_main("System.out.println(Character.digit('x', 10));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn character_digit_returns_minus_one_for_punctuation() {
    let out = run_main("System.out.println(Character.digit('#', 16));");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn character_get_numeric_value_of_ascii_digit_seven() {
    let out = run_main("System.out.println(Character.getNumericValue('7'));");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn character_get_numeric_value_of_uppercase_hex_letter() {
    let out = run_main("System.out.println(Character.getNumericValue('A'));");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn character_get_numeric_value_of_lowercase_hex_letter() {
    let out = run_main("System.out.println(Character.getNumericValue('f'));");
    assert_eq!(out, vec!["15"]);
}

#[test]
fn character_get_numeric_value_returns_minus_one_for_symbol() {
    let out = run_main("System.out.println(Character.getNumericValue('!'));");
    assert_eq!(out, vec!["-1"]);
}
