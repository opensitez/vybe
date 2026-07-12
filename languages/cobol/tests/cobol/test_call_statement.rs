use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_call_no_params() {
    compile_ok(&p(
        "",
        r#"
    CALL "SUBPROG".
"#,
    ));
}

#[test]
fn test_call_using_parameters() {
    compile_ok(&p(
        r#"
01 WS-A PIC 9(3) VALUE 100.
01 WS-B PIC X(5) VALUE "HELLO".
"#,
        r#"
    CALL "SUBPROG" USING BY REFERENCE WS-A
                         BY CONTENT WS-B.
    CALL "SUBPROG" USING BY VALUE WS-A.
"#,
    ));
}

#[test]
fn test_call_returning() {
    compile_ok(&p(
        r#"
01 WS-RET PIC 9(3) VALUE 0.
"#,
        r#"
    CALL "SUBPROG" RETURNING WS-RET.
"#,
    ));
}

#[test]
fn test_call_exception_handling() {
    compile_ok(&p(
        "",
        r#"
    CALL "NONEXIST"
        ON EXCEPTION
            DISPLAY "ERROR"
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL.
"#,
    ));
}
