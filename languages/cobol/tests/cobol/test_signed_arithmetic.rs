use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn signed_add_positive_result() {
    let out = run_prints(&p(
        "01 A PIC S9(4) VALUE +50.\n01 B PIC S9(4) VALUE +30.\n01 R PIC S9(5) VALUE 0.",
        "    ADD A B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["+0080"]);
}

#[test]
fn signed_subtract_negative_result() {
    let out = run_prints(&p(
        "01 A PIC S9(4) VALUE +10.\n01 B PIC S9(4) VALUE +30.\n01 R PIC S9(5) VALUE 0.",
        "    SUBTRACT B FROM A GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["-0020"]);
}

#[test]
fn signed_multiply_negative_by_positive() {
    let out = run_prints(&p(
        "01 A PIC S9(3) VALUE -4.\n01 B PIC S9(3) VALUE +7.\n01 R PIC S9(5) VALUE 0.",
        "    MULTIPLY A BY B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["-0028"]);
}

#[test]
fn signed_multiply_negative_by_negative() {
    let out = run_prints(&p(
        "01 A PIC S9(3) VALUE -5.\n01 B PIC S9(3) VALUE -6.\n01 R PIC S9(5) VALUE 0.",
        "    MULTIPLY A BY B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["+0030"]);
}

#[test]
fn signed_compute_chain_with_negatives() {
    let out = run_prints(&p(
        "01 R PIC S9(5) VALUE 0.",
        "    COMPUTE R = -3 * -4 + -2.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["+0010"]);
}

#[test]
fn signed_value_negative_literal() {
    let out = run_prints(&p(
        "01 N PIC S9(3) VALUE -99.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["-099"]);
}

#[test]
fn signed_move_negative_to_signed_field() {
    let out = run_prints(&p(
        "01 SRC PIC S9(3) VALUE -42.\n01 DST PIC S9(5) VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["-00042"]);
}

#[test]
fn signed_add_negative_literal() {
    let out = run_prints(&p(
        "01 A PIC S9(4) VALUE +100.",
        "    ADD -35 TO A.\n    DISPLAY A.",
    ));
    assert_eq!(out, vec!["+0065"]);
}

#[test]
fn signed_subtract_positive_gives_positive() {
    let out = run_prints(&p(
        "01 A PIC S9(4) VALUE -10.\n01 B PIC S9(4) VALUE -30.",
        "    SUBTRACT B FROM A.\n    DISPLAY A.",
    ));
    assert_eq!(out, vec!["+0020"]);
}

#[test]
fn signed_compute_absolute_value_simulation() {
    let out = run_prints(&p(
        "01 X PIC S9(4) VALUE -25.\n01 ABS-X PIC 9(4) VALUE 0.",
        "    IF X < 0\n        COMPUTE ABS-X = -X\n    ELSE\n        MOVE X TO ABS-X\n    END-IF.\n    DISPLAY ABS-X.",
    ));
    assert_eq!(out, vec!["0025"]);
}

#[test]
fn signed_divide_positive_by_negative() {
    let out = run_prints(&p(
        "01 R PIC S9(4) VALUE 0.",
        "    COMPUTE R = 20 / -4.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["-0005"]);
}

#[test]
fn signed_s9_pic_zero_shows_plus() {
    let out = run_prints(&p(
        "01 N PIC S9(3) VALUE 0.",
        "    DISPLAY N.",
    ));
    assert_eq!(out, vec!["+000"]);
}

#[test]
fn signed_condition_less_than_zero() {
    let out = run_prints(&p(
        "01 N PIC S9(3) VALUE -1.",
        "    IF N < 0\n        DISPLAY \"NEG\"\n    ELSE\n        DISPLAY \"POS\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["NEG"]);
}

#[test]
fn signed_condition_greater_than_zero() {
    let out = run_prints(&p(
        "01 N PIC S9(3) VALUE +5.",
        "    IF N > 0\n        DISPLAY \"POS\"\n    ELSE\n        DISPLAY \"NEG\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["POS"]);
}

#[test]
fn signed_arithmetic_overflow_to_negative() {
    let out = run_prints(&p(
        "01 A PIC S9(3) VALUE +50.\n01 B PIC S9(3) VALUE +80.\n01 R PIC S9(4) VALUE 0.",
        "    ADD A B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["+0130"]);
}

#[test]
fn signed_s9v99_decimal_computation() {
    let out = run_prints(&p(
        "01 A PIC S9(3)V99 VALUE -12.50.\n01 B PIC S9(3)V99 VALUE +7.25.\n01 R PIC S9(4)V99 VALUE 0.",
        "    ADD A B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["-0005.25"]);
}

#[test]
fn signed_compare_negative_to_zero() {
    let out = run_prints(&p(
        "01 N PIC S9(3) VALUE -1.",
        "    IF N = 0\n        DISPLAY \"ZERO\"\n    ELSE\n        DISPLAY \"NOT ZERO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["NOT ZERO"]);
}

#[test]
fn signed_move_unsigned_to_signed_field() {
    let out = run_prints(&p(
        "01 SRC PIC 9(3) VALUE 123.\n01 DST PIC S9(4) VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["+0123"]);
}

#[test]
fn signed_sum_positive_and_negative_zero_result() {
    let out = run_prints(&p(
        "01 A PIC S9(3) VALUE +50.\n01 B PIC S9(3) VALUE -50.\n01 R PIC S9(4) VALUE 0.",
        "    ADD A B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["+0000"]);
}

#[test]
fn signed_compute_with_variable_factor() {
    let out = run_prints(&p(
        "01 FACTOR PIC S9(2) VALUE -3.\n01 R PIC S9(5) VALUE 0.",
        "    COMPUTE R = 10 * FACTOR.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["-00030"]);
}

#[test]
fn signed_field_used_as_loop_counter() {
    let out = run_prints(&p(
        "01 I PIC S9(3) VALUE -2.\n01 S PIC S9(5) VALUE 0.",
        "    PERFORM UNTIL I > 2\n        ADD I TO S\n        ADD 1 TO I\n    END-PERFORM.\n    DISPLAY S.",
    ));
    // -2 + -1 + 0 + 1 + 2 = 0
    assert_eq!(out, vec!["+0000"]);
}

#[test]
fn signed_compare_two_negatives() {
    let out = run_prints(&p(
        "01 A PIC S9(3) VALUE -10.\n01 B PIC S9(3) VALUE -5.",
        "    IF A < B\n        DISPLAY \"A LESS\"\n    ELSE\n        DISPLAY \"B LESS OR EQUAL\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["A LESS"]);
}

#[test]
fn signed_negate_via_compute() {
    let out = run_prints(&p(
        "01 N PIC S9(4) VALUE +42.\n01 R PIC S9(4) VALUE 0.",
        "    COMPUTE R = 0 - N.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["-0042"]);
}

#[test]
fn signed_multiply_zero_result() {
    let out = run_prints(&p(
        "01 A PIC S9(3) VALUE -7.\n01 R PIC S9(5) VALUE 0.",
        "    MULTIPLY A BY 0 GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["+0000"]);
}

#[test]
fn signed_field_in_if_after_loop() {
    let out = run_prints(&p(
        "01 N PIC S9(4) VALUE 100.\n01 I PIC 9(2) VALUE 0.",
        "    PERFORM UNTIL I >= 20\n        SUBTRACT 10 FROM N\n        ADD 1 TO I\n    END-PERFORM.\n    IF N < 0\n        DISPLAY \"NEG\"\n    ELSE\n        DISPLAY \"POS\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["NEG"]);
}

#[test]
fn signed_s9_comp_compiles() {
    compile_ok(&p(
        "01 N PIC S9(8) COMP VALUE 0.",
        "    ADD 1 TO N.",
    ));
}

#[test]
fn signed_v99_positive_and_negative_display() {
    let out = run_prints(&p(
        "01 A PIC S9V99 VALUE +3.50.\n01 B PIC S9V99 VALUE -1.25.\n01 R PIC S9(2)V99 VALUE 0.",
        "    ADD A B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["+02.25"]);
}

#[test]
fn signed_absolute_value_with_function() {
    compile_ok(&p(
        "01 N PIC S9(5) VALUE -1234.\n01 R PIC 9(5) VALUE 0.",
        "    COMPUTE R = FUNCTION ABS(N).",
    ));
}

#[test]
fn signed_add_three_mixed_sign_operands() {
    let out = run_prints(&p(
        "01 A PIC S9(3) VALUE +10.\n01 B PIC S9(3) VALUE -5.\n01 C PIC S9(3) VALUE -3.\n01 R PIC S9(4) VALUE 0.",
        "    ADD A B C GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["+0002"]);
}

#[test]
fn signed_subtract_negative_from_negative() {
    let out = run_prints(&p(
        "01 A PIC S9(4) VALUE -20.\n01 B PIC S9(4) VALUE -8.\n01 R PIC S9(5) VALUE 0.",
        "    SUBTRACT B FROM A GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["-0012"]);
}
