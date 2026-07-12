use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn move_numeric_to_alpha_runtime() {
    let out = run_prints(&p(
        "01 WS-SRC PIC 9(3) VALUE 42.\n01 WS-DST PIC X(5) VALUE SPACES.",
        "    MOVE WS-SRC TO WS-DST.\n    DISPLAY WS-DST.",
    ));
    assert_eq!(out, vec!["42"]);
}

#[test]
fn move_alpha_to_alpha_padding_runtime() {
    let out = run_prints(&p(
        "01 WS-SRC PIC X(3) VALUE \"ABC\".\n01 WS-DST PIC X(6).",
        "    MOVE WS-SRC TO WS-DST.\n    DISPLAY WS-DST.",
    ));
    assert_eq!(out, vec!["ABC"]);
}

#[test]
fn move_alpha_to_alpha_truncate_runtime() {
    let out = run_prints(&p(
        "01 WS-SRC PIC X(6) VALUE \"ABCDEF\".\n01 WS-DST PIC X(3).",
        "    MOVE WS-SRC TO WS-DST.\n    DISPLAY WS-DST.",
    ));
    assert_eq!(out, vec!["ABCDEF"]);
}

#[test]
fn move_decimal_to_integer_runtime() {
    let out = run_prints(&p(
        "01 WS-SRC PIC 9(3)V99 VALUE 123.45.\n01 WS-DST PIC 9(3).",
        "    MOVE WS-SRC TO WS-DST.\n    DISPLAY WS-DST.",
    ));
    assert_eq!(out, vec!["123"]);
}

#[test]
fn compute_then_move_conversion_compiles() {
    compile_ok(&p(
        "01 WS-A PIC 9(3) VALUE 3.\n01 WS-B PIC 9(3) VALUE 4.\n01 WS-C PIC 9(5).",
        "    COMPUTE WS-C = WS-A * WS-B.",
    ));
}

#[test]
fn string_numeric_conversion_via_move_compiles() {
    compile_ok(&p(
        "01 WS-STR PIC X(3) VALUE \"123\".\n01 WS-NUM PIC 9(3).",
        "    MOVE WS-STR TO WS-NUM.",
    ));
}

#[test]
fn move_signed_numeric_to_display_runtime() {
    let out = run_prints(&p(
        "01 WS-S PIC S9(3) VALUE -12.\n01 WS-D PIC X(6).",
        "    MOVE WS-S TO WS-D.\n    DISPLAY WS-D.",
    ));
    assert_eq!(out, vec!["-12"]);
}

#[test]
fn compute_numeric_then_move_to_alphanumeric_runtime() {
    let out = run_prints(&p(
        "01 A PIC 9(2) VALUE 8.\n01 B PIC 9(2) VALUE 5.\n01 N PIC 9(3) VALUE 0.\n01 T PIC X(5).",
        "    COMPUTE N = A + B.\n    MOVE N TO T.\n    DISPLAY T.",
    ));
    assert_eq!(out, vec!["13"]);
}
