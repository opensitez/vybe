use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_intrinsics_char_ord_conversions() {
    let output = run_prints(&p(
        r#"
01 WS-CHAR PIC X VALUE SPACES.
01 WS-ORD PIC 9(3) VALUE 0.
"#,
        r#"
    MOVE FUNCTION CHAR(65) TO WS-CHAR.
    DISPLAY WS-CHAR.
    COMPUTE WS-ORD = FUNCTION ORD("A").
    DISPLAY WS-ORD.
"#,
    ));
    assert_eq!(output, vec!["A", "065"]);
}

#[test]
fn test_intrinsics_ord_min_max() {
    compile_ok(&p(
        "01 WS-ORD PIC 9(5).",
        r#"
    COMPUTE WS-ORD = FUNCTION ORD-MAX.
    COMPUTE WS-ORD = FUNCTION ORD-MIN.
"#,
    ));
}

#[test]
fn test_intrinsics_test_numval() {
    compile_ok(&p(
        "01 WS-RES PIC 9(9).",
        r#"
    COMPUTE WS-RES = FUNCTION TEST-NUMVAL("123.45").
    COMPUTE WS-RES = FUNCTION TEST-NUMVAL("ABC").
"#,
    ));
}

#[test]
fn test_intrinsics_case_conversion() {
    let output = run_prints(&p(
        r#"
01 WS-TEXT PIC X(5) VALUE "Hello".
01 WS-UP PIC X(5).
01 WS-LOW PIC X(5).
"#,
        r#"
    MOVE FUNCTION UPPER-CASE(WS-TEXT) TO WS-UP.
    MOVE FUNCTION LOWER-CASE(WS-TEXT) TO WS-LOW.
    DISPLAY WS-UP.
    DISPLAY WS-LOW.
"#,
    ));
    assert_eq!(output, vec!["HELLO", "hello"]);
}

#[test]
fn test_intrinsics_ord_max_min_runtime() {
    let output = run_prints(&p(
        "01 WS-ORD PIC 9(5) VALUE 0.",
        r#"
    COMPUTE WS-ORD = FUNCTION ORD-MAX.
    DISPLAY WS-ORD.
    COMPUTE WS-ORD = FUNCTION ORD-MIN.
    DISPLAY WS-ORD.
"#,
    ));
    assert_eq!(output.len(), 2);
}
