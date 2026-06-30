use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_intrinsics_trim_modes() {
    let output = run_prints(&p(
        "01 WS-STR PIC X(10) VALUE \"  hello   \".",
        r#"
    DISPLAY FUNCTION TRIM(WS-STR LEADING).
    DISPLAY FUNCTION TRIM(WS-STR TRAILING).
"#,
    ));
    assert_eq!(output, vec!["hello   ", "  hello"]);
}

#[test]
fn test_intrinsics_substitute_case() {
    compile_ok(&p(
        r#"
01 WS-STR PIC X(10) VALUE "Hello World".
01 WS-DST PIC X(15).
"#,
        r#"
    MOVE FUNCTION SUBSTITUTE-CASE(WS-STR "world" "COBOL") TO WS-DST.
"#,
    ));
}

#[test]
fn test_intrinsics_numval_cf() {
    compile_ok(&p(
        "01 WS-NUM PIC 9(5)V99.",
        r#"
    COMPUTE WS-NUM = FUNCTION NUMVAL-C("$12,345.67").
    COMPUTE WS-NUM = FUNCTION NUMVAL-F("123.45").
"#,
    ));
}

#[test]
fn test_intrinsics_string_lengths() {
    compile_ok(&p(
        r#"
01 WS-STR PIC X(10) VALUE "hello".
01 WS-LEN PIC 9(5).
"#,
        r#"
    COMPUTE WS-LEN = FUNCTION BYTE-LENGTH(WS-STR).
    COMPUTE WS-LEN = FUNCTION STORED-CHAR-LENGTH(WS-STR).
"#,
    ));
}
