use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn compute_addition_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 3.\n01 WS-B PIC 9(3) VALUE 4.\n01 WS-C PIC 9(3).",
        "    COMPUTE WS-C = WS-A + WS-B.",
    ));
}

#[test]
fn compute_subtraction_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 9.\n01 WS-B PIC 9(3) VALUE 4.\n01 WS-C PIC 9(3).",
        "    COMPUTE WS-C = WS-A - WS-B.",
    ));
}

#[test]
fn compute_multiplication_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 6.\n01 WS-B PIC 9(3) VALUE 7.\n01 WS-C PIC 9(3).",
        "    COMPUTE WS-C = WS-A * WS-B.",
    ));
}

#[test]
fn compute_division_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 8.\n01 WS-B PIC 9(3) VALUE 2.\n01 WS-C PIC 9(3).",
        "    COMPUTE WS-C = WS-A / WS-B.",
    ));
}

#[test]
fn add_subtract_and_move_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 10.\n01 WS-B PIC 9(3) VALUE 5.\n01 WS-C PIC 9(3).",
        "    ADD WS-A TO WS-B.\n    SUBTRACT WS-B FROM WS-A.\n    MOVE WS-A TO WS-C.",
    ));
}

#[test]
fn compute_addition_runtime_displays_expected_sum() {
    let output = run_prints(&p(
        "01 WS-A PIC 9(2) VALUE 12.\n01 WS-B PIC 9(2) VALUE 7.\n01 WS-C PIC 9(3) VALUE 0.",
        "    COMPUTE WS-C = WS-A + WS-B.\n    DISPLAY WS-C.",
    ));
    assert_eq!(output, vec!["19"]);
}

#[test]
fn compute_precedence_runtime_matches_parentheses() {
    let output = run_prints(&p(
        "01 A PIC 9 VALUE 2.\n01 B PIC 9 VALUE 3.\n01 C PIC 9 VALUE 4.\n01 R PIC 9(2) VALUE 0.",
        "    COMPUTE R = (A + B) * C.\n    DISPLAY R.",
    ));
    assert_eq!(output, vec!["20"]);
}

#[test]
fn divide_giving_remainder_runtime_reports_expected_values() {
    let output = run_prints(&p(
        "01 A PIC 99 VALUE 7.\n01 B PIC 99 VALUE 3.\n01 Q PIC 99 VALUE 0.\n01 R PIC 99 VALUE 0.",
        "    DIVIDE A BY B GIVING Q REMAINDER R.\n    DISPLAY Q.\n    DISPLAY R.",
    ));
    assert_eq!(output, vec!["2", "1"]);
}

#[test]
fn compute_mixed_add_sub_runtime_expected_result() {
    let output = run_prints(&p(
        "01 A PIC 99 VALUE 20.\n01 B PIC 99 VALUE 5.\n01 C PIC 99 VALUE 3.\n01 R PIC 99 VALUE 0.",
        "    COMPUTE R = A - B + C.\n    DISPLAY R.",
    ));
    assert_eq!(output, vec!["18"]);
}
