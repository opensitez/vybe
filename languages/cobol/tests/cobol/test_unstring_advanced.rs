use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_unstring_delimited_by_comma() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(10) VALUE "A,BB,CCC".
01 WS-F1 PIC X(3) VALUE SPACES.
01 WS-F2 PIC X(3) VALUE SPACES.
01 WS-F3 PIC X(3) VALUE SPACES.
"#,
        r#"
    UNSTRING WS-SRC DELIMITED BY "," INTO WS-F1 WS-F2 WS-F3.
    DISPLAY WS-F1.
    DISPLAY WS-F2.
    DISPLAY WS-F3.
"#,
    ));
    assert_eq!(output, vec!["A  ", "BB ", "CCC"]);
}

#[test]
fn test_unstring_tallying_pointer() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(10) VALUE "A,BB,CCC".
01 WS-F1 PIC X(3) VALUE SPACES.
01 WS-F2 PIC X(3) VALUE SPACES.
01 WS-F3 PIC X(3) VALUE SPACES.
01 WS-TALLY PIC 9(2) VALUE 0.
01 WS-PTR PIC 9(2) VALUE 1.
"#,
        r#"
    UNSTRING WS-SRC DELIMITED BY "," INTO WS-F1 WS-F2 WS-F3
        WITH POINTER WS-PTR
        TALLYING IN WS-TALLY.
    DISPLAY WS-TALLY.
"#,
    ));
    assert_eq!(output, vec!["03"]);
}

#[test]
fn test_unstring_counts_delimiters() {
    compile_ok(&p(
        r#"
01 WS-SRC PIC X(10) VALUE "A,BB;CCC".
01 WS-F1 PIC X(3).
01 WS-F2 PIC X(3).
01 WS-DEL1 PIC X.
01 WS-CNT1 PIC 99.
"#,
        r#"
    UNSTRING WS-SRC DELIMITED BY "," OR ";"
        INTO WS-F1 DELIMITER IN WS-DEL1 COUNT IN WS-CNT1
             WS-F2.
"#,
    ));
}

#[test]
fn test_unstring_all_delimiters() {
    compile_ok(&p(
        r#"
01 WS-SRC PIC X(10) VALUE "A,,BB,,CCC".
01 WS-F1 PIC X(3).
01 WS-F2 PIC X(3).
"#,
        r#"
    UNSTRING WS-SRC DELIMITED BY ALL "," INTO WS-F1 WS-F2.
"#,
    ));
}
