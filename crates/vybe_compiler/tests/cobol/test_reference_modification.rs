use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_refmod_single_char() {
    let output = run_prints(&p(
        "01 WS-TXT PIC X(5) VALUE \"ABCDE\".",
        r#"
    DISPLAY WS-TXT(1:1).
    DISPLAY WS-TXT(5:1).
"#,
    ));
    assert_eq!(output, vec!["A", "E"]);
}

#[test]
fn test_refmod_omitted_length() {
    compile_ok(&p(
        "01 WS-TXT PIC X(5) VALUE \"ABCDE\".",
        r#"
    DISPLAY WS-TXT(3:).
"#,
    ));
}

#[test]
fn test_refmod_target_write() {
    let output = run_prints(&p(
        "01 WS-TXT PIC X(6) VALUE \"AABBCC\".",
        r#"
    MOVE "XX" TO WS-TXT(3:2).
    DISPLAY WS-TXT.
"#,
    ));
    assert_eq!(output, vec!["AAXXCC"]);
}

#[test]
fn test_refmod_dynamic_variables() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(10) VALUE "ABCDEFGHIJ".
01 WS-START PIC 99 VALUE 3.
01 WS-LEN PIC 99 VALUE 4.
01 WS-SUB PIC X(4) VALUE SPACES.
"#,
        r#"
    MOVE WS-TXT(WS-START:WS-LEN) TO WS-SUB.
    DISPLAY WS-SUB.
"#,
    ));
    assert_eq!(output, vec!["CDEF"]);
}

#[test]
fn test_refmod_in_condition() {
    let output = run_prints(&p(
        "01 WS-TXT PIC X(5) VALUE \"ABCDE\".",
        r#"
    IF WS-TXT(1:3) = "ABC"
        DISPLAY "MATCH"
    ELSE
        DISPLAY "NO-MATCH"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["MATCH"]);
}
