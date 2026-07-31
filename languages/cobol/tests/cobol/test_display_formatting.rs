use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn display_single_string_literal() {
    let out = run_prints(&p("", "    DISPLAY \"HELLO WORLD\"."));
    assert_eq!(out, vec!["HELLO WORLD"]);
}

#[test]
fn display_numeric_variable() {
    let out = run_prints(&p(
        "01 N PIC 9(4) VALUE 2024.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["2024"]);
}

#[test]
fn display_two_literals_sequential() {
    let out = run_prints(&p(
        "",
        "    DISPLAY \"LINE1\".\n    DISPLAY \"LINE2\".",
    ));
    assert_eq!(out, vec!["LINE1", "LINE2"]);
}

#[test]
fn display_multiple_items_on_one_statement() {
    let out = run_prints(&p(
        "01 A PIC X(3) VALUE \"FOO\".\n01 B PIC X(3) VALUE \"BAR\".",
        "    DISPLAY A \" \" B.",
    ));
    assert_eq!(out, vec!["FOO BAR"]);
}

#[test]
fn display_numeric_after_compute() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R = 7 * 8.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["56"]);
}

#[test]
fn display_alphanumeric_right_padded_pic() {
    let out = run_prints(&p(
        "01 S PIC X(8) VALUE \"COBOL\".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["COBOL   "]);
}

#[test]
fn display_zero_value_numeric() {
    let out = run_prints(&p(
        "01 N PIC 9(4) VALUE ZERO.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0000"]);
}

#[test]
fn display_spaces_figurative() {
    let out = run_prints(&p(
        "01 S PIC X(4) VALUE SPACES.",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["    "]);
}

#[test]
fn display_concatenated_string_and_number() {
    let out = run_prints(&p(
        "01 N PIC 9(3) VALUE 42.",
        "    DISPLAY \"VALUE: \" N.",
    ));
    assert_eq!(out, vec!["VALUE: 042"]);
}

#[test]
fn display_boolean_condition_result() {
    let out = run_prints(&p(
        "01 A PIC 9 VALUE 5.\n01 B PIC 9 VALUE 3.",
        "    IF A > B\n        DISPLAY \"GREATER\"\n    ELSE\n        DISPLAY \"NOT GREATER\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["GREATER"]);
}

#[test]
fn display_three_numeric_fields() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 1.\n01 Y PIC 9 VALUE 2.\n01 Z PIC 9 VALUE 3.",
        "    DISPLAY X.\n    DISPLAY Y.\n    DISPLAY Z.",
    ));
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn display_signed_negative_value() {
    let out = run_prints(&p(
        "01 N PIC S9(4) VALUE -42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["-0042"]);
}

#[test]
fn display_decimal_field() {
    let out = run_prints(&p(
        "01 D PIC 9(3)V99 VALUE 123.45.",
        "    DISPLAY D.",
    ));
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn display_group_item_shows_raw_bytes() {
    let out = run_prints(&p(
        "01 GRP.\n   05 GA PIC X(3) VALUE \"ABC\".\n   05 GB PIC X(3) VALUE \"DEF\".",
        "    DISPLAY GRP.",
    ));
    assert_eq!(out, vec!["ABCDEF"]);
}

#[test]
fn display_after_move_spaces() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"HELLO\".",
        "    MOVE SPACES TO S.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["     "]);
}

#[test]
fn display_numeric_after_add() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 100.\n01 B PIC 9(3) VALUE 55.",
        "    ADD B TO A.\n    DISPLAY A.",
    ));
    assert_eq!(out, vec!["155"]);
}

#[test]
fn display_after_subtract() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 200.\n01 B PIC 9(3) VALUE 75.",
        "    SUBTRACT B FROM A.\n    DISPLAY A.",
    ));
    assert_eq!(out, vec!["125"]);
}

#[test]
fn display_inside_perform_times() {
    let out = run_prints(&p(
        "",
        "    PERFORM 4 TIMES\n        DISPLAY \"*\"\n    END-PERFORM.",
    ));
    assert_eq!(out, vec!["*", "*", "*", "*"]);
}

#[test]
fn display_inside_if_true_branch() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 1.",
        "    IF X = 1\n        DISPLAY \"ONE\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ONE"]);
}

#[test]
fn display_inside_if_false_branch() {
    let out = run_prints(&p(
        "01 X PIC 9 VALUE 0.",
        "    IF X = 1\n        DISPLAY \"ONE\"\n    ELSE\n        DISPLAY \"ZERO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["ZERO"]);
}

#[test]
fn display_alphanumeric_value_all_literal() {
    let out = run_prints(&p(
        "01 S PIC X(6) VALUE ALL \"XY\".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["XYXYXY"]);
}

#[test]
fn display_upon_console_compiles() {
    compile_ok(&p(
        "",
        "    DISPLAY \"OUTPUT\" UPON CONSOLE.",
    ));
}

#[test]
fn display_with_no_advancing_compiles() {
    compile_ok(&p(
        "",
        "    DISPLAY \"PROMPT\" WITH NO ADVANCING.",
    ));
}

#[test]
fn display_multiple_literals_space_separated() {
    let out = run_prints(&p(
        "",
        "    DISPLAY \"A\" \"B\" \"C\".",
    ));
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn display_high_value_field_compiles() {
    compile_ok(&p(
        "01 S PIC X(4).",
        "    MOVE HIGH-VALUES TO S.\n    DISPLAY S.",
    ));
}

#[test]
fn display_numeric_leading_zero_preserved() {
    let out = run_prints(&p(
        "01 N PIC 9(5) VALUE 7.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["00007"]);
}

#[test]
fn display_field_after_initialize() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE \"HELLO\".",
        "    INITIALIZE S.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["     "]);
}

#[test]
fn display_two_different_types() {
    let out = run_prints(&p(
        "01 NAME PIC X(4) VALUE \"JOHN\".\n01 AGE PIC 9(3) VALUE 30.",
        "    DISPLAY NAME \" \" AGE.",
    ));
    assert_eq!(out, vec!["JOHN 030"]);
}

#[test]
fn display_empty_literal() {
    let out = run_prints(&p(
        "",
        "    DISPLAY \"\".",
    ));
    assert_eq!(out, vec![""]);
}

#[test]
fn display_long_literal() {
    let out = run_prints(&p(
        "",
        "    DISPLAY \"ABCDEFGHIJKLMNOPQRSTUVWXYZ\".",
    ));
    assert_eq!(out, vec!["ABCDEFGHIJKLMNOPQRSTUVWXYZ"]);
}

#[test]
fn display_result_of_multiply_giving() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 9.\n01 B PIC 9(3) VALUE 9.\n01 R PIC 9(5) VALUE 0.",
        "    MULTIPLY A BY B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["81"]);
}
