use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn alphanumeric_type_compiles() {
    compile_ok(&p(
        "01 WS-TXT PIC X(10) VALUE \"HELLO\".",
        "    DISPLAY WS-TXT.",
    ));
}

#[test]
fn numeric_display_type_compiles() {
    compile_ok(&p("01 WS-NUM PIC 9(5) VALUE 12345.", "    DISPLAY WS-NUM."));
}

#[test]
fn signed_numeric_type_compiles() {
    compile_ok(&p("01 WS-NUM PIC S9(4) VALUE -25.", "    DISPLAY WS-NUM."));
}

#[test]
fn decimal_implied_type_compiles() {
    compile_ok(&p(
        "01 WS-AMT PIC 9(3)V99 VALUE 123.45.",
        "    DISPLAY WS-AMT.",
    ));
}

#[test]
fn binary_usage_type_compiles() {
    compile_ok(&p(
        "01 WS-BIN PIC 9(4) USAGE IS BINARY VALUE 7.",
        "    ADD 1 TO WS-BIN.",
    ));
}

#[test]
fn comp3_usage_type_compiles() {
    compile_ok(&p(
        "01 WS-PACK PIC 9(5) USAGE IS COMP-3 VALUE 20.",
        "    ADD 5 TO WS-PACK.",
    ));
}

#[test]
fn pointer_usage_type_compiles() {
    compile_ok(&p("01 WS-PTR USAGE IS POINTER.", "    SET WS-PTR TO NULL."));
}

#[test]
fn function_pointer_usage_type_compiles() {
    compile_ok(&p(
        "01 WS-FPTR USAGE IS FUNCTION-POINTER.",
        "    DISPLAY \"FPTR\".",
    ));
}

#[test]
fn procedure_pointer_usage_type_compiles() {
    compile_ok(&p(
        "01 WS-PPTR USAGE IS PROCEDURE-POINTER.",
        "    DISPLAY \"PPTR\".",
    ));
}

#[test]
fn type_display_runtime_check() {
    let out = run_prints(&p(
        "01 WS-A PIC X(3) VALUE \"ABC\".\n01 WS-B PIC 9(3) VALUE 123.",
        "    DISPLAY WS-A.\n    DISPLAY WS-B.",
    ));
    assert_eq!(out, vec!["ABC", "123"]);
}
