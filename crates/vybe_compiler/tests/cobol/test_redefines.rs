use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_redefines_basic() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(4) VALUE "AAAA".
01 WS-ALIAS REDEFINES WS-BASE PIC X(4).
"#,
        r#"
    MOVE "BBBB" TO WS-ALIAS.
    DISPLAY WS-BASE.
"#,
    ));
    assert_eq!(output, vec!["BBBB"]);
}

#[test]
fn test_redefines_types() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(4) VALUE "1234".
01 WS-ALIAS REDEFINES WS-BASE PIC 9(4).
"#,
        r#"
    DISPLAY WS-ALIAS.
"#,
    ));
    assert_eq!(output, vec!["1234"]);
}

#[test]
fn test_redefines_partial() {
    let output = run_prints(&p(
        r#"
01 WS-BASE PIC X(4) VALUE "ABCD".
01 WS-ALIAS REDEFINES WS-BASE PIC X(2).
"#,
        r#"
    DISPLAY WS-ALIAS.
"#,
    ));
    assert_eq!(output, vec!["AB"]);
}

#[test]
fn test_redefines_group() {
    compile_ok(&p(
        r#"
01 WS-BASE PIC X(6) VALUE "123456".
01 WS-ALIAS REDEFINES WS-BASE.
   05 WS-A PIC X(3).
   05 WS-B PIC X(3).
"#,
        r#"
    DISPLAY WS-A.
"#,
    ));
}

#[test]
fn test_redefines_multiple() {
    compile_ok(&p(
        r#"
01 WS-BASE PIC X(4) VALUE "AAAA".
01 WS-ALIAS1 REDEFINES WS-BASE PIC 9(4).
01 WS-ALIAS2 REDEFINES WS-BASE PIC X(4).
"#,
        r#"
    DISPLAY WS-ALIAS1.
    DISPLAY WS-ALIAS2.
"#,
    ));
}
