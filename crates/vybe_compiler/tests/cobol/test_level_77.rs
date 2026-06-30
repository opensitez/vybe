use super::helpers::{compile_ok, run_prints};

#[test]
fn test_level_77_elementary_numeric() {
    let output = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
77 WS-NUM PIC 9(3) VALUE 10.
PROCEDURE DIVISION.
    ADD 5 TO WS-NUM.
    DISPLAY WS-NUM.
    STOP RUN.
"#,
    );
    assert_eq!(output, vec!["015"]);
}

#[test]
fn test_level_77_elementary_alpha() {
    let output = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
77 WS-STR PIC X(5) VALUE "HELLO".
PROCEDURE DIVISION.
    DISPLAY WS-STR.
    STOP RUN.
"#,
    );
    assert_eq!(output, vec!["HELLO"]);
}

#[test]
fn test_level_77_usages() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
77 WS-BIN PIC 9(4) USAGE IS BINARY VALUE 123.
77 WS-COMP3 PIC 9(4) USAGE IS PACKED-DECIMAL VALUE 456.
PROCEDURE DIVISION.
    ADD 1 TO WS-BIN WS-COMP3.
    STOP RUN.
"#,
    );
}
