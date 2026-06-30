use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_qualification_basic() {
    compile_ok(&p(
        r#"
01 WS-GROUP-A.
   05 WS-NAME PIC X(5) VALUE "ALICE".
01 WS-GROUP-B.
   05 WS-NAME PIC X(5) VALUE "BOB  ".
"#,
        r#"
    DISPLAY WS-NAME IN WS-GROUP-A.
    DISPLAY WS-NAME OF WS-GROUP-B.
"#,
    ));
}

#[test]
fn test_qualification_nested() {
    compile_ok(&p(
        r#"
01 WS-TOP.
   05 WS-SUB.
      10 WS-FIELD PIC X(3) VALUE "XYZ".
"#,
        r#"
    DISPLAY WS-FIELD IN WS-SUB IN WS-TOP.
"#,
    ));
}

#[test]
fn test_qualification_statement_usage() {
    compile_ok(&p(
        r#"
01 WS-GROUP-A.
   05 WS-VAL PIC 9(3) VALUE 10.
01 WS-GROUP-B.
   05 WS-VAL PIC 9(3) VALUE 20.
"#,
        r#"
    ADD WS-VAL IN WS-GROUP-A TO WS-VAL IN WS-GROUP-B.
"#,
    ));
}
