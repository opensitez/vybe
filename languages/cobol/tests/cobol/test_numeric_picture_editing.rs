use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn pic_9_basic_display() {
    let out = run_prints(&p(
        "01 N PIC 9 VALUE 7.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["7"]);
}

#[test]
fn pic_99_two_digits() {
    let out = run_prints(&p(
        "01 N PIC 99 VALUE 42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["42"]);
}

#[test]
fn pic_9_parenthetic_notation() {
    let out = run_prints(&p(
        "01 N PIC 9(4) VALUE 1234.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["1234"]);
}

#[test]
fn pic_x_single_char() {
    let out = run_prints(&p(
        "01 C PIC X VALUE \"A\".",
        "    DISPLAY C.",
    ));
    assert_eq!(out, vec!["A"]);
}

#[test]
fn pic_x_parenthetic_fills() {
    let out = run_prints(&p(
        "01 S PIC X(6) VALUE \"COBOL\".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["COBOL "]);
}

#[test]
fn pic_a_alphabetic() {
    compile_ok(&p(
        "01 S PIC A(5) VALUE \"HELLO\".",
        "    DISPLAY S.",
    ));
}

#[test]
fn pic_9_v_99_decimal() {
    let out = run_prints(&p(
        "01 D PIC 9V99 VALUE 3.14.",
        "    DISPLAY D.",
    ));
    assert_eq!(out, vec!["314"]);
}

#[test]
fn pic_99_v_9999_decimal_leading_zero() {
    let out = run_prints(&p(
        "01 D PIC 99V9999 VALUE 01.2345.",
        "    DISPLAY D.",
    ));
    assert_eq!(out, vec!["012345"]);
}

#[test]
fn pic_s9_signed_positive() {
    let out = run_prints(&p(
        "01 N PIC S9(4) VALUE +123.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["+0123"]);
}

#[test]
fn pic_s9_signed_negative() {
    let out = run_prints(&p(
        "01 N PIC S9(4) VALUE -123.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["-0123"]);
}

#[test]
fn pic_9_fills_leading_zeros() {
    let out = run_prints(&p(
        "01 N PIC 9(6) VALUE 42.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["000042"]);
}

#[test]
fn pic_x_truncates_right_when_longer_source() {
    let out = run_prints(&p(
        "01 SRC PIC X(6) VALUE \"ABCDEF\".\n01 DST PIC X(3) VALUE \"   \".",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn pic_x_pads_right_when_shorter_source() {
    let out = run_prints(&p(
        "01 SRC PIC X(3) VALUE \"ABC\".\n01 DST PIC X(6) VALUE \"XXXXXX\".",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["ABC   "]);
}

#[test]
fn pic_9_truncates_left_high_digits() {
    let out = run_prints(&p(
        "01 SRC PIC 9(5) VALUE 12345.\n01 DST PIC 9(3) VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["345"]);
}

#[test]
fn pic_9_extends_left_with_zeros() {
    let out = run_prints(&p(
        "01 SRC PIC 9(3) VALUE 42.\n01 DST PIC 9(6) VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["000042"]);
}

#[test]
fn pic_9v99_to_9_truncates_decimal() {
    let out = run_prints(&p(
        "01 SRC PIC 9(3)V99 VALUE 123.45.\n01 DST PIC 9(3) VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["123"]);
}

#[test]
fn pic_9_to_9v99_extends_decimal() {
    let out = run_prints(&p(
        "01 SRC PIC 9(3) VALUE 123.\n01 DST PIC 9(3)V99 VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["12300"]);
}

#[test]
fn pic_s9_positive_becomes_positive_when_moved() {
    let out = run_prints(&p(
        "01 SRC PIC S9(3) VALUE -99.\n01 DST PIC 9(3) VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["099"]);
}

#[test]
fn pic_9_comp3_value_roundtrip() {
    let out = run_prints(&p(
        "01 N PIC 9(7) COMP-3 VALUE 1234567.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["1234567"]);
}

#[test]
fn pic_x_longer_string_truncated() {
    let out = run_prints(&p(
        "01 DST PIC X(4) VALUE \"XXXX\".",
        "    MOVE \"ABCDEFGH\" TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["ABCD"]);
}

#[test]
fn pic_9_literal_all_nines() {
    let out = run_prints(&p(
        "01 N PIC 9(5) VALUE 99999.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["99999"]);
}

#[test]
fn pic_x_value_spaces() {
    let out = run_prints(&p(
        "01 S PIC X(5) VALUE SPACES.",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["     "]);
}

#[test]
fn pic_9_value_zeros() {
    let out = run_prints(&p(
        "01 N PIC 9(5) VALUE ZEROS.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["00000"]);
}

#[test]
fn pic_s9v99_negative_with_decimal() {
    let out = run_prints(&p(
        "01 N PIC S9(3)V99 VALUE -100.50.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["-10050"]);
}

#[test]
fn pic_x_overlong_pic_truncation_at_definition_size() {
    let out = run_prints(&p(
        "01 S PIC X(3) VALUE \"AB\".",
        "    DISPLAY S.",
    ));
    assert_eq!(out, vec!["AB "]);
}

#[test]
fn pic_9_level77() {
    let out = run_prints(&p(
        "77 N PIC 9(4) VALUE 2025.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["2025"]);
}

#[test]
fn pic_9_inside_group() {
    let out = run_prints(&p(
        "01 GRP.\n   05 PART1 PIC 9(3) VALUE 100.\n   05 PART2 PIC 9(3) VALUE 200.",
        "    DISPLAY PART1.\n    DISPLAY PART2.",
    ));
    assert_eq!(out, vec!["100", "200"]);
}

#[test]
fn pic_x_inside_group() {
    let out = run_prints(&p(
        "01 GRP.\n   05 FNAME PIC X(5) VALUE \"ALICE\".\n   05 LNAME PIC X(5) VALUE \"JONES\".",
        "    DISPLAY FNAME.\n    DISPLAY LNAME.",
    ));
    assert_eq!(out, vec!["ALICE", "JONES"]);
}

#[test]
fn pic_9_arithmetic_then_display() {
    let out = run_prints(&p(
        "01 A PIC 9(3) VALUE 111.\n01 B PIC 9(3) VALUE 222.\n01 R PIC 9(4) VALUE 0.",
        "    ADD A B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["0333"]);
}

#[test]
fn pic_99_value_zero_padded() {
    let out = run_prints(&p(
        "01 N PIC 99 VALUE 7.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["07"]);
}
