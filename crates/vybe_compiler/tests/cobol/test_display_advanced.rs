use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_display_concatenated() {
    let output = run_prints(&p(
        r#"
01 WS-NAME PIC X(5) VALUE "ALICE".
01 WS-AGE PIC 9(3) VALUE 30.
"#,
        r#"
    DISPLAY "NAME: " WS-NAME " AGE: " WS-AGE.
"#,
    ));
    assert_eq!(output, vec!["NAME: ALICE AGE: 030"]);
}

#[test]
fn test_display_no_advancing() {
    compile_ok(&p(
        "",
        r#"
    DISPLAY "HELLO " WITH NO ADVANCING.
    DISPLAY "WORLD".
"#,
    ));
}

#[test]
fn test_display_special_destinations() {
    compile_ok(&p(
        "",
        r#"
    DISPLAY "HELLO" UPON CONSOLE.
    DISPLAY "ERROR" UPON SYSERR.
"#,
    ));
}

#[test]
fn test_display_group() {
    let output = run_prints(&p(
        r#"
01 WS-GROUP.
   05 WS-A PIC X(3) VALUE "123".
   05 WS-B PIC X(3) VALUE "456".
"#,
        r#"
    DISPLAY WS-GROUP.
"#,
    ));
    assert_eq!(output, vec!["123456"]);
}

#[test]
fn test_display_inline_function() {
    let output = run_prints(&p(
        "01 WS-STR PIC X(5) VALUE \"hello\".",
        r#"
    DISPLAY FUNCTION UPPER-CASE(WS-STR).
"#,
    ));
    assert_eq!(output, vec!["HELLO"]);
}

#[test]
fn test_display_signed_negative() {
    let output = run_prints(&p(
        "01 WS-NUM PIC S9(3) VALUE -42.",
        r#"
    DISPLAY WS-NUM.
"#,
    ));
    assert!(output.len() >= 1);
}
