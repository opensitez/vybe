use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_string_delimited_by_size() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC X(5) VALUE "HELLO".
01 WS-B PIC X(5) VALUE "WORLD".
01 WS-DST PIC X(11) VALUE SPACES.
"#,
        r#"
    STRING WS-A DELIMITED BY SIZE
           " " DELIMITED BY SIZE
           WS-B DELIMITED BY SIZE
           INTO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["HELLO WORLD"]);
}

#[test]
fn test_string_with_pointer() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(5) VALUE "WORLD".
01 WS-DST PIC X(11) VALUE "HELLO      ".
01 WS-PTR PIC 9(2) VALUE 7.
"#,
        r#"
    STRING WS-SRC DELIMITED BY SIZE INTO WS-DST WITH POINTER WS-PTR.
    DISPLAY WS-DST.
    DISPLAY WS-PTR.
"#,
    ));
    assert_eq!(output, vec!["HELLO WORLD", "12"]);
}

#[test]
fn test_string_overflow() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC X(10) VALUE "ABCDEFGHIJ".
01 WS-B PIC X(10) VALUE "KLQRSTUVWX".
01 WS-DST PIC X(15).
"#,
        r#"
    STRING WS-A WS-B DELIMITED BY SIZE INTO WS-DST
        ON OVERFLOW
            DISPLAY "OVERFLOW"
        NOT ON OVERFLOW
            DISPLAY "OK"
    END-STRING.
"#,
    ));
    assert_eq!(output, vec!["OVERFLOW"]);
}

#[test]
fn test_string_delimited_by_space() {
    let output = run_prints(&p(
        r#"
01 WS-SRC PIC X(10) VALUE "HELLO     ".
01 WS-DST PIC X(5) VALUE SPACES.
"#,
        r#"
    STRING WS-SRC DELIMITED BY SPACE INTO WS-DST.
    DISPLAY WS-DST.
"#,
    ));
    assert_eq!(output, vec!["HELLO"]);
}
