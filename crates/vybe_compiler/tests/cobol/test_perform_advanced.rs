use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_perform_para_times() {
    let output = run_prints(&p(
        "",
        r#"
    PERFORM TEST-PARA 3 TIMES.
    STOP RUN.
TEST-PARA.
    DISPLAY "HELLO".
"#,
    ));
    assert_eq!(output, vec!["HELLO", "HELLO", "HELLO"]);
}

#[test]
fn test_perform_para_until() {
    let output = run_prints(&p(
        "01 WS-COUNT PIC 9 VALUE 0.",
        r#"
    PERFORM TEST-PARA UNTIL WS-COUNT >= 3.
    STOP RUN.
TEST-PARA.
    ADD 1 TO WS-COUNT.
    DISPLAY WS-COUNT.
"#,
    ));
    assert_eq!(output, vec!["1", "2", "3"]);
}

#[test]
fn test_perform_para_varying() {
    let output = run_prints(&p(
        "01 WS-I PIC 9 VALUE 0.",
        r#"
    PERFORM TEST-PARA VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3.
    STOP RUN.
TEST-PARA.
    DISPLAY WS-I.
"#,
    ));
    assert_eq!(output, vec!["1", "2", "3"]);
}

#[test]
fn test_perform_thru_span() {
    let output = run_prints(&p(
        "",
        r#"
    PERFORM PARA1 THRU PARA3.
    STOP RUN.
PARA1.
    DISPLAY "P1".
PARA2.
    DISPLAY "P2".
PARA3.
    DISPLAY "P3".
"#,
    ));
    assert_eq!(output, vec!["P1", "P2", "P3"]);
}

#[test]
fn test_perform_test_after() {
    let output = run_prints(&p(
        "01 WS-I PIC 9 VALUE 5.",
        r#"
    PERFORM WITH TEST AFTER UNTIL WS-I >= 5
        DISPLAY "ONCE"
        ADD 1 TO WS-I
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["ONCE"]);
}

#[test]
fn test_perform_varying_down() {
    let output = run_prints(&p(
        "01 WS-I PIC S9(3).",
        r#"
    PERFORM VARYING WS-I FROM 5 BY -2 UNTIL WS-I < 1
        DISPLAY WS-I
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["005", "003", "001"]);
}

#[test]
fn test_perform_exit_perform() {
    let output = run_prints(&p(
        "01 WS-I PIC 9.",
        r#"
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 5
        IF WS-I = 3
            EXIT PERFORM
        END-IF
        DISPLAY WS-I
    END-PERFORM.
"#,
    ));
    assert_eq!(output, vec!["1", "2"]);
}
