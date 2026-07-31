use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn add_rounded_mode_nearest_even() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 10.5.\n01 B PIC 9(3)V9 VALUE 3.5.\n01 R PIC 9(4) VALUE 0.",
        "    ADD A B GIVING R ROUNDED MODE NEAREST-EVEN.",
    ));
}

#[test]
fn add_rounded_default() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 10.5.\n01 R PIC 9(3) VALUE 0.",
        "    ADD 1.5 TO A ROUNDED.",
    ));
}

#[test]
fn subtract_rounded() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 10.5.\n01 R PIC 9(3) VALUE 0.",
        "    SUBTRACT 0.5 FROM A ROUNDED.",
    ));
}

#[test]
fn multiply_rounded() {
    compile_ok(&p(
        "01 A PIC 9(3)V99 VALUE 3.33.\n01 R PIC 9(4)V9 VALUE 0.",
        "    MULTIPLY A BY 3 GIVING R ROUNDED.",
    ));
}

#[test]
fn divide_rounded() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 10.\n01 R PIC 9(3)V9 VALUE 0.",
        "    DIVIDE 3 INTO A GIVING R ROUNDED.",
    ));
}

#[test]
fn compute_rounded() {
    compile_ok(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R ROUNDED = 7 / 3.",
    ));
}

#[test]
fn compute_rounded_truncates_at_pic_scale() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R ROUNDED = 9.6.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["10"]);
}

#[test]
fn compute_rounded_down_when_lt_half() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R ROUNDED = 9.4.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["9"]);
}

#[test]
fn add_rounded_mode_truncation() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 5.6.\n01 R PIC 9(3) VALUE 0.",
        "    ADD A TO R ROUNDED MODE TRUNCATION.",
    ));
}

#[test]
fn add_rounded_mode_toward_greater() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 5.3.\n01 R PIC 9(3) VALUE 0.",
        "    ADD A TO R ROUNDED MODE TOWARD-GREATER.",
    ));
}

#[test]
fn add_rounded_mode_toward_lesser() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 5.7.\n01 R PIC 9(3) VALUE 0.",
        "    ADD A TO R ROUNDED MODE TOWARD-LESSER.",
    ));
}

#[test]
fn add_rounded_mode_away_from_zero() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 5.5.\n01 R PIC 9(3) VALUE 0.",
        "    ADD A TO R ROUNDED MODE AWAY-FROM-ZERO.",
    ));
}

#[test]
fn add_rounded_mode_nearest_toward_zero() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 5.5.\n01 R PIC 9(3) VALUE 0.",
        "    ADD A TO R ROUNDED MODE NEAREST-TOWARD-ZERO.",
    ));
}

#[test]
fn compute_rounded_integer_input_unchanged() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R ROUNDED = 42.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["42"]);
}

#[test]
fn subtract_giving_rounded() {
    compile_ok(&p(
        "01 A PIC 9(4)V9 VALUE 100.5.\n01 B PIC 9(4)V9 VALUE 33.3.\n01 R PIC 9(4) VALUE 0.",
        "    SUBTRACT B FROM A GIVING R ROUNDED.",
    ));
}

#[test]
fn multiply_giving_rounded() {
    compile_ok(&p(
        "01 A PIC 9(3)V99 VALUE 1.25.\n01 B PIC 9(2) VALUE 3.\n01 R PIC 9(3)V9 VALUE 0.",
        "    MULTIPLY A BY B GIVING R ROUNDED.",
    ));
}

#[test]
fn divide_by_giving_rounded() {
    let out = run_prints(&p(
        "01 R PIC 9(3)V9 VALUE 0.",
        "    DIVIDE 10 BY 3 GIVING R ROUNDED.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["003.3"]);
}

#[test]
fn add_rounded_gives_correct_penny_rounding() {
    let out = run_prints(&p(
        "01 TAX PIC 9(5)V99 VALUE 0.\n01 RATE PIC 9V9(4) VALUE 0.0875.",
        "    COMPUTE TAX ROUNDED = 100 * RATE.\n    DISPLAY TAX.",
    ));
    assert_eq!(out, vec!["008.75"]);
}

#[test]
fn compute_rounded_to_whole_number_five_rounds_up() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R ROUNDED = 4.5.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["5"]);
}

#[test]
fn compute_rounded_two_decimal_places() {
    let out = run_prints(&p(
        "01 R PIC 9(3)V99 VALUE 0.",
        "    COMPUTE R ROUNDED = 10 / 3.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["003.33"]);
}

#[test]
fn compute_rounded_with_large_divisor() {
    let out = run_prints(&p(
        "01 R PIC 9(4)V99 VALUE 0.",
        "    COMPUTE R ROUNDED = 1000 / 7.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["0142.86"]);
}

#[test]
fn add_rounded_with_on_size_error() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 999.9.",
        "    ADD 0.5 TO A ROUNDED\n    ON SIZE ERROR\n        DISPLAY \"OVERFLOW\"\n    END-ADD.",
    ));
}

#[test]
fn subtract_rounded_with_size_error() {
    compile_ok(&p(
        "01 A PIC 9(3) VALUE 0.",
        "    SUBTRACT 500 FROM A ROUNDED\n    ON SIZE ERROR\n        DISPLAY \"UNDERFLOW\"\n    END-SUBTRACT.",
    ));
}

#[test]
fn compute_rounded_product_of_decimals() {
    let out = run_prints(&p(
        "01 R PIC 9(4) VALUE 0.",
        "    COMPUTE R ROUNDED = 3.7 * 2.4.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["9"]);
}

#[test]
fn multiply_rounded_three_place_decimal() {
    compile_ok(&p(
        "01 A PIC 9(3)V999 VALUE 1.333.\n01 R PIC 9(4)V99 VALUE 0.",
        "    MULTIPLY 3 BY A GIVING R ROUNDED.",
    ));
}

#[test]
fn divide_remainder_rounded() {
    compile_ok(&p(
        "01 Q PIC 9(4) VALUE 0.\n01 REM PIC 9(4)V9 VALUE 0.",
        "    DIVIDE 7 INTO 22 GIVING Q REMAINDER REM.",
    ));
}

#[test]
fn compute_rounded_exact_half_point() {
    let out = run_prints(&p(
        "01 R PIC 9(3) VALUE 0.",
        "    COMPUTE R ROUNDED = 2.5.\n    DISPLAY R.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn add_two_rounded_targets() {
    compile_ok(&p(
        "01 A PIC 9(3)V9 VALUE 1.5.\n01 B PIC 9(3)V9 VALUE 2.5.\n01 R1 PIC 9(3) VALUE 0.\n01 R2 PIC 9(3) VALUE 0.",
        "    ADD A B GIVING R1 ROUNDED R2 ROUNDED.",
    ));
}

#[test]
fn compute_rounded_negative_result() {
    compile_ok(&p(
        "01 R PIC S9(4) VALUE 0.",
        "    COMPUTE R ROUNDED = -3.7.",
    ));
}
