use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn binary_comp_basic_add() {
    let out = run_prints(&p(
        "01 N PIC 9(8) COMP VALUE 0.",
        "    ADD 1 TO N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["1"]);
}

#[test]
fn binary_comp_large_value() {
    let out = run_prints(&p(
        "01 N PIC 9(9) COMP VALUE 0.",
        "    MOVE 1000000 TO N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["1000000"]);
}

#[test]
fn binary_comp3_packed_decimal_add() {
    let out = run_prints(&p(
        "01 N PIC 9(7) COMP-3 VALUE 0.",
        "    ADD 12345 TO N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn binary_comp3_multiply() {
    let out = run_prints(&p(
        "01 A PIC 9(5) COMP-3 VALUE 100.\n01 R PIC 9(7) COMP-3 VALUE 0.",
        "    MULTIPLY A BY 7 GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["700"]);
}

#[test]
fn binary_comp_subtract() {
    let out = run_prints(&p(
        "01 N PIC 9(8) COMP VALUE 500.",
        "    SUBTRACT 250 FROM N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["250"]);
}

#[test]
fn binary_comp_divide() {
    let out = run_prints(&p(
        "01 N PIC 9(8) COMP VALUE 100.\n01 R PIC 9(4) COMP VALUE 0.",
        "    DIVIDE 4 INTO N GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["25"]);
}

#[test]
fn binary_comp_compare_to_literal() {
    let out = run_prints(&p(
        "01 N PIC 9(4) COMP VALUE 42.",
        "    IF N = 42\n        DISPLAY \"YES\"\n    ELSE\n        DISPLAY \"NO\"\n    END-IF.",
    ));
    assert_eq!(out, vec!["YES"]);
}

#[test]
fn binary_comp_in_loop() {
    let out = run_prints(&p(
        "01 I PIC 9(4) COMP VALUE 0.\n01 S PIC 9(6) COMP VALUE 0.",
        "    PERFORM UNTIL I >= 100\n        ADD 1 TO I\n        ADD I TO S\n    END-PERFORM.\n    DISPLAY S.",
    ));
    assert_eq!(out, vec!["5050"]);
}

#[test]
fn binary_comp3_with_decimal() {
    let out = run_prints(&p(
        "01 A PIC 9(5)V99 COMP-3 VALUE 123.45.\n01 B PIC 9(5)V99 COMP-3 VALUE 67.89.\n01 R PIC 9(6)V99 COMP-3 VALUE 0.",
        "    ADD A B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["191.34"]);
}

#[test]
fn binary_comp5_basic_value() {
    compile_ok(&p("01 N PIC 9(9) COMP-5 VALUE 0.", "    ADD 42 TO N."));
}

#[test]
fn binary_comp1_float_compiles() {
    compile_ok(&p("01 F COMP-1 VALUE 3.14.", "    COMPUTE F = F * 2."));
}

#[test]
fn binary_comp2_double_compiles() {
    compile_ok(&p("01 D COMP-2 VALUE 3.14159.", "    COMPUTE D = D + 1."));
}

#[test]
fn binary_packed_decimal_synonymous_with_comp3() {
    compile_ok(&p(
        "01 N PIC 9(7) PACKED-DECIMAL VALUE 0.",
        "    ADD 100 TO N.",
    ));
}

#[test]
fn binary_synonymous_with_comp() {
    compile_ok(&p("01 N PIC 9(8) BINARY VALUE 0.", "    ADD 1 TO N."));
}

#[test]
fn binary_comp_zero_init() {
    let out = run_prints(&p("01 N PIC 9(5) COMP VALUE 0.", "    DISPLAY N."));
    assert_eq!(out, vec!["0"]);
}

#[test]
fn binary_comp3_zero_result_after_subtract() {
    let out = run_prints(&p(
        "01 N PIC 9(5) COMP-3 VALUE 100.",
        "    SUBTRACT 100 FROM N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["0"]);
}

#[test]
fn binary_comp_signed_negative() {
    let out = run_prints(&p("01 N PIC S9(5) COMP VALUE -500.", "    DISPLAY N."));
    assert_eq!(out, vec!["-500"]);
}

#[test]
fn binary_comp_add_to_signed() {
    let out = run_prints(&p(
        "01 N PIC S9(5) COMP VALUE -100.",
        "    ADD 200 TO N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["+100"]);
}

#[test]
fn binary_comp_used_as_subscript() {
    let out = run_prints(&p(
        "01 IDX PIC 9(4) COMP VALUE 2.\n01 T.\n   05 E PIC X OCCURS 5 TIMES.",
        "    MOVE \"B\" TO E(IDX).\n    DISPLAY E(IDX).",
    ));
    assert_eq!(out, vec!["B"]);
}

#[test]
fn binary_comp3_two_digits_precision() {
    let out = run_prints(&p(
        "01 PRICE PIC 9(5)V99 COMP-3 VALUE 19.99.\n01 QTY PIC 9(3) COMP VALUE 3.\n01 TOTAL PIC 9(7)V99 COMP-3 VALUE 0.",
        "    MULTIPLY PRICE BY QTY GIVING TOTAL.\n    DISPLAY TOTAL.",
    ));
    assert_eq!(out, vec!["59.97"]);
}

#[test]
fn binary_comp_accumulate_large() {
    let out = run_prints(&p(
        "01 N PIC 9(9) COMP VALUE 0.\n01 I PIC 9(4) COMP VALUE 0.",
        "    PERFORM UNTIL I >= 1000\n        ADD 1 TO I\n        ADD I TO N\n    END-PERFORM.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["500500"]);
}

#[test]
fn binary_comp3_remainder() {
    let out = run_prints(&p(
        "01 D PIC 9(5) COMP-3 VALUE 17.\n01 Q PIC 9(4) COMP-3 VALUE 0.\n01 R PIC 9(4) COMP-3 VALUE 0.",
        "    DIVIDE 5 INTO D GIVING Q REMAINDER R.\n    DISPLAY Q.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn binary_comp_move_from_standard() {
    let out = run_prints(&p(
        "01 SRC PIC 9(5) VALUE 12345.\n01 DST PIC 9(5) COMP VALUE 0.",
        "    MOVE SRC TO DST.\n    DISPLAY DST.",
    ));
    assert_eq!(out, vec!["12345"]);
}

#[test]
fn binary_comp_in_evaluate() {
    let out = run_prints(&p(
        "01 CODE PIC 9(2) COMP VALUE 2.",
        "    EVALUATE CODE\n        WHEN 1 DISPLAY \"ONE\"\n        WHEN 2 DISPLAY \"TWO\"\n        WHEN OTHER DISPLAY \"OTHER\"\n    END-EVALUATE.",
    ));
    assert_eq!(out, vec!["TWO"]);
}

#[test]
fn binary_comp3_compute_power() {
    compile_ok(&p(
        "01 BASE PIC 9(3) COMP-3 VALUE 4.\n01 R PIC 9(5) COMP-3 VALUE 0.",
        "    COMPUTE R = BASE ** 3.",
    ));
}

#[test]
fn binary_half_word_comp_boundary() {
    let out = run_prints(&p(
        "01 N PIC 9(4) COMP VALUE 9999.",
        "    ADD 1 TO N.\n    DISPLAY N.",
    ));
    assert_eq!(out, vec!["10000"]);
}

#[test]
fn binary_comp3_high_precision_decimal() {
    let out = run_prints(&p(
        "01 A PIC 9(7)V9(4) COMP-3 VALUE 1234567.1234.\n01 B PIC 9(7)V9(4) COMP-3 VALUE 0001.0001.\n01 R PIC 9(9)V9(4) COMP-3 VALUE 0.",
        "    ADD A B GIVING R.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["1234568.1235"]);
}
