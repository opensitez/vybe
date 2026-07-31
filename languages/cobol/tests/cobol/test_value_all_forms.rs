use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn value_all_literal_fills_field() {
    let out = run_prints(&p(
        "01 S PIC X(8) VALUE ALL \"AB\".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["ABABABAB"]);
}

#[test]
fn value_all_single_char_fills() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE ALL \"*\".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["*****"]);
}

#[test]
fn value_zeros_numeric_field() {
    let out = run_prints(&p(
        "01 N PIC 9(5) VALUE ZEROS.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["00000"]);
}

#[test]
fn value_zeroes_synonymous() {
    let out = run_prints(&p(
        "01 N PIC 9(3) VALUE ZEROES.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["000"]);
}

#[test]
fn value_zero_singular() {
    let out = run_prints(&p(
        "01 N PIC 9(4) VALUE ZERO.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0000"]);
}

#[test]
fn value_spaces_fills_alpha() {
    let out = run_prints(&p(
        "01 S PIC X(6) VALUE SPACES.",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["      "]);
}

#[test]
fn value_space_singular() {
    let out = run_prints(&p(
        "01 S PIC X(4) VALUE SPACE.",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["    "]);
}

#[test]
fn value_high_values_compiles() {
    compile_ok(&p(
        "01 S PIC X(4) VALUE HIGH-VALUES.",
        "    DISPLAY S.",
    ));
}

#[test]
fn value_high_value_singular_compiles() {
    compile_ok(&p(
        "01 S PIC X(4) VALUE HIGH-VALUE.",
        "    DISPLAY S.",
    ));
}

#[test]
fn value_low_values_compiles() {
    compile_ok(&p(
        "01 S PIC X(4) VALUE LOW-VALUES.",
        "    DISPLAY S.",
    ));
}

#[test]
fn value_low_value_singular_compiles() {
    compile_ok(&p(
        "01 S PIC X(4) VALUE LOW-VALUE.",
        "    DISPLAY S.",
    ));
}

#[test]
fn value_null_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER VALUE NULL.",
        "    CONTINUE.",
    ));
}

#[test]
fn value_nulls_compiles() {
    compile_ok(&p(
        "01 P USAGE POINTER VALUE NULLS.",
        "    CONTINUE.",
    ));
}

#[test]
fn value_quotes_single_char() {
    compile_ok(&p(
        "01 S PIC X VALUE QUOTE.",
        "    DISPLAY S.",
    ));
}

#[test]
fn value_literal_string() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"HELLO\".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["HELLO"]);
}

#[test]
fn value_numeric_literal() {
    let out = run_prints(&p(
        "01 N PIC 9(4) VALUE 1234.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["1234"]);
}

#[test]
fn value_signed_negative_literal() {
    let out = run_prints(&p(
        "01 N PIC S9(4) VALUE -9.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["-0009"]);
}

#[test]
fn value_decimal_literal() {
    let out = run_prints(&p(
        "01 D PIC 9(3)V99 VALUE 12.34.",
        "    DISPLAY D.",
    ));
    assert_eq!(out, vec!["01234"]);
}

#[test]
fn value_all_zero_string_fills_with_zero_chars() {
    let out = run_prints(&p(
        "01 S PIC X(6) VALUE ALL \"0\".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["000000"]);
}

#[test]
fn value_zeros_on_alpha_field() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE ZEROS.",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["00000"]);
}

#[test]
fn value_spaces_on_numeric_field() {
    let out = run_prints(&p(
        "01 N PIC 9(5) VALUE SPACES.",
        "    DISPLAY N.",
    ));
    // SPACES on numeric initializes to 0
    assert_eq!(out, vec!["00000"]);
}

#[test]
fn value_all_space_char() {
    let out = run_prints(&p(
        "01 S PIC X(4) VALUE ALL \" \".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["    "]);
}

#[test]
fn value_all_three_char_pattern_truncated() {
    let out = run_prints(&p(
        "01 S PIC X(7) VALUE ALL \"123\".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["1231231"]);
}

#[test]
fn value_spaces_and_then_moved_to() {
    let out = run_prints(&p(
        "01 S PIC X(8) VALUE SPACES.",
        "    MOVE \"OK\" TO S.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["OK      "]);
}

#[test]
fn value_zeros_and_then_added_to() {
    let out = run_prints(&p(
        "01 N PIC 9(4) VALUE ZEROS.",
        "    ADD 100 TO N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0100"]);
}

#[test]
fn value_high_values_compared_compiles() {
    compile_ok(&p(
        "01 S PIC X(4) VALUE HIGH-VALUES.",
        "    IF S > \"ZZZZ\"\n        DISPLAY \"HIGH\"\n    ELSE\n        DISPLAY \"NOT HIGH\"\n    END-IF.",
    ));
}

#[test]
fn value_low_values_compared_compiles() {
    compile_ok(&p(
        "01 S PIC X(4) VALUE LOW-VALUES.",
        "    IF S < \"AAAA\"\n        DISPLAY \"LOW\"\n    ELSE\n        DISPLAY \"NOT LOW\"\n    END-IF.",
    ));
}

#[test]
fn value_all_in_group_item() {
    let out = run_prints(&p(
        "01 GRP.\n   05 P1 PIC X(3) VALUE ALL \"A\".\n   05 P2 PIC X(3) VALUE ALL \"B\".",
        "    DISPLAY GRP.",
    ));
    assert_eq!(out, vec!["AAABBB"]);
}

#[test]
fn value_clause_on_level77() {
    let out = run_prints(&p(
        "77 STANDALONE PIC X(5) VALUE \"FIVE5\".",
        "    DISPLAY STANDALONE.",
    ));
    assert_eq!(out, vec!["FIVE5"]);
}
