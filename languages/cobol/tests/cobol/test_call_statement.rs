use super::helpers::{compile_ok, run_prints};

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

#[test]
fn test_call_exception_runtime() {
    let out = run_prints(
        &p(
            "",
            r#"
    CALL "NONEXIST"
        ON EXCEPTION
            DISPLAY "ERR"
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL.
"#,
        ),
    );
    assert_eq!(out, vec!["ERR"]);
}

#[test]
fn test_call_with_reference_semantics_runtime() {
    let out = run_prints(
        &p(
            "01 WS-A PIC 9(3) VALUE 1.",
            r#"
    CALL "SUBPROG" USING BY VALUE WS-A
        ON EXCEPTION
            DISPLAY WS-A
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL.
"#,
        ),
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_call_with_mixed_modes_and_on_exception() {
    let out = run_prints(
        &p(
            "01 WS-ARG PIC X(5) VALUE \"HELLO\".\n01 WS-NUM PIC 9(3) VALUE 123.\n01 WS-PROG PIC X(20) VALUE \"SUBPROG\".",
            r#"
    CALL WS-PROG
        USING BY REFERENCE WS-ARG
        ON EXCEPTION
            DISPLAY "ERR"
        NOT ON EXCEPTION
            DISPLAY "OK"
        END-CALL.
"#,
        ),
    );
    assert_eq!(out, vec!["ERR"]);
}

#[test]
fn test_call_uses_dynamic_name() {
    compile_ok(&p(
        "01 WS-PROG PIC X(20) VALUE \"SUBPROG\".",
        r#"
    CALL WS-PROG
        ON EXCEPTION DISPLAY "MISS"
        NOT ON EXCEPTION DISPLAY "HIT"
    END-CALL.
"#,
    ));
}

#[test]
fn test_call_with_mixed_passing_modes() {
    let out = run_prints(
        &p(
            "01 WS-ARG PIC 9(3) VALUE 10.\n01 WS-TEXT PIC X(4) VALUE \"ABCD\".",
            r#"
    CALL "SUBPROG"
        USING BY VALUE WS-ARG
        BY REFERENCE WS-TEXT
        ON EXCEPTION
            DISPLAY "ERR"
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL.
"#,
        ),
    );
    assert_eq!(out, vec!["ERR"]);
}

#[test]
fn test_call_with_reference_only_on_exception() {
    let out = run_prints(
        &p(
            "01 WS-TEXT PIC X(5) VALUE \"HELLO\".",
            r#"
    CALL "MISSING-PROG" USING BY REFERENCE WS-TEXT
        ON EXCEPTION
            DISPLAY "EX"
    END-CALL.
"#,
        ),
    );
    assert_eq!(out, vec!["EX"]);
}

#[test]
fn test_call_with_explicit_returning_and_exception() {
    let out = run_prints(
        &p(
            "01 WS-RET PIC 9(3) VALUE 0.\n01 WS-ARG PIC 9(3) VALUE 55.",
            r#"
    CALL "SUBPROG" RETURNING WS-RET
        USING WS-ARG
        ON EXCEPTION
            DISPLAY WS-ARG
            DISPLAY WS-RET
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL.
"#,
        ),
    );
    assert_eq!(out, vec!["55", "0"]);
}

#[test]
fn test_call_by_name_without_end_call_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    CALL \"SUBPROG\".\n    STOP RUN.",
    );
}
