use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_usage_clauses() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(4) USAGE IS DISPLAY.
01 WS-B PIC 9(4) USAGE IS BINARY.
01 WS-C PIC 9(4) USAGE IS PACKED-DECIMAL.
01 WS-D PIC 9(4) USAGE IS COMPUTATIONAL.
01 WS-E PIC 9(4) USAGE IS COMP.
01 WS-F PIC 9(4) USAGE IS COMP-3.
"#,
        r#"
    ADD 5 TO WS-A.
    ADD 5 TO WS-B.
    ADD 5 TO WS-C.
    ADD 5 TO WS-D.
    ADD 5 TO WS-E.
    ADD 5 TO WS-F.
    DISPLAY WS-A.
    DISPLAY WS-B.
    DISPLAY WS-C.
    DISPLAY WS-D.
    DISPLAY WS-E.
    DISPLAY WS-F.
"#,
    ));
    assert_eq!(output, vec!["5", "5", "5", "5", "5", "5"]);
}

#[test]
fn test_usage_floats() {
    let output = run_prints(&p(
        r#"
01 WS-FLOAT-S USAGE IS COMP-1.
01 WS-FLOAT-D USAGE IS COMP-2.
"#,
        r#"
    MOVE 1 TO WS-FLOAT-S.
    MOVE 2 TO WS-FLOAT-D.
    DISPLAY WS-FLOAT-S.
    DISPLAY WS-FLOAT-D.
"#,
    ));
    assert_eq!(output.len(), 2);
}

#[test]
fn test_usage_native_binary() {
    let output = run_prints(&p(
        r#"
01 WS-BIN PIC 9(4) USAGE IS COMP-5.
"#,
        r#"
    MOVE 0 TO WS-BIN.
    ADD 1 TO WS-BIN.
    DISPLAY WS-BIN.
"#,
    ));
    assert_eq!(output, vec!["1"]);
}

#[test]
fn test_usage_pointers() {
    let output = run_prints(&p(
        r#"
01 WS-PTR USAGE IS POINTER.
01 WS-PPTR USAGE IS PROCEDURE-POINTER.
01 WS-FLAG PIC X VALUE 'N'.
"#,
        r#"
    SET WS-PTR TO NULL.
    IF WS-PTR = NULL
        MOVE 'Y' TO WS-FLAG
    END-IF.
    DISPLAY WS-FLAG.
"#,
    ));
    assert_eq!(output, vec!["Y"]);
}

#[test]
fn test_usage_group_inheritance() {
    let output = run_prints(&p(
        r#"
01 WS-GROUP USAGE IS BINARY.
   05 WS-A PIC 9(4).
   05 WS-B PIC 9(4).
"#,
        r#"
    MOVE 10 TO WS-A.
    MOVE 20 TO WS-B.
    ADD 10 TO WS-A.
    ADD 5 TO WS-B.
    DISPLAY WS-A.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["20", "25"]);
}
