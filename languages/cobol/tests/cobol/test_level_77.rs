use super::helpers::run_prints;

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
    let output = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
77 WS-BIN PIC 9(4) USAGE IS BINARY VALUE 123.
77 WS-COMP3 PIC 9(4) USAGE IS PACKED-DECIMAL VALUE 456.
77 WS-OK PIC X VALUE 'N'.
PROCEDURE DIVISION.
    ADD 1 TO WS-BIN WS-COMP3.
    IF WS-BIN = 124
        IF WS-COMP3 = 457
            MOVE 'Y' TO WS-OK
        END-IF
    END-IF
    DISPLAY WS-OK.
    STOP RUN.
"#,
    );
    assert_eq!(output, vec!["Y"]);
}

#[test]
fn test_level_77_with_conditional_name() {
    let output = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
77 WS-STATE PIC X VALUE "Y".
88 STATE-OK VALUE "Y".
77 WS-COUNT PIC 9(2) VALUE 0.
PROCEDURE DIVISION.
    IF STATE-OK
        ADD 1 TO WS-COUNT
    END-IF
    DISPLAY WS-COUNT.
    STOP RUN.
"#,
    );
    assert_eq!(output, vec!["01"]);
}

#[test]
fn test_level_77_display_usage_and_redefines() {
    let output = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
77 WS-NUM PIC 9(5) USAGE COMP.
77 WS-ALPHA PIC X(3) VALUE "ABC".
PROCEDURE DIVISION.
    MOVE 12 TO WS-NUM.
    DISPLAY WS-ALPHA.
    DISPLAY WS-NUM.
    STOP RUN.
"#,
    );
    assert_eq!(output, vec!["ABC", "00012"]);
}
