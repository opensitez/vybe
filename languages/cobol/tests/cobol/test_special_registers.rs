use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_return_code() {
    compile_ok(&p(
        "",
        r#"
    MOVE 8 TO RETURN-CODE.
    IF RETURN-CODE NOT = 0
        DISPLAY "ERROR"
    END-IF.
"#,
    ));
}

#[test]
fn test_length_of() {
    let output = run_prints(&p(
        r#"
01 WS-STR PIC X(10) VALUE "HELLO".
01 WS-GROUP.
   05 WS-A PIC X(3).
   05 WS-B PIC X(7).
01 WS-LEN PIC 9(3).
"#,
        r#"
    COMPUTE WS-LEN = FUNCTION LENGTH(WS-STR).
    DISPLAY WS-LEN.
    COMPUTE WS-LEN = FUNCTION LENGTH(WS-GROUP).
    DISPLAY WS-LEN.
"#,
    ));
    assert_eq!(output, vec!["010", "010"]);
}

#[test]
fn test_tally() {
    compile_ok(&p(
        r#"
01 WS-STR PIC X(10) VALUE "A B C D".
"#,
        r#"
    INSPECT WS-STR TALLYING TALLY FOR ALL " ".
"#,
    ));
}

#[test]
fn test_xml_json_code() {
    compile_ok(&p(
        "",
        r#"
    DISPLAY XML-CODE.
    DISPLAY JSON-CODE.
"#,
    ));
}
