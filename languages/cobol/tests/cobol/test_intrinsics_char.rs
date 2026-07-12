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
