use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_inspect_tallying_spaces() {
    let output = run_prints(&p(
        r#"
01 WS-STR PIC X(10) VALUE "A B C D  ".
01 WS-CNT PIC 9(3) VALUE 0.
"#,
        r#"
    INSPECT WS-STR TALLYING WS-CNT FOR ALL " ".
    DISPLAY WS-CNT.
"#,
    ));
    assert_eq!(output, vec!["005"]);
}

#[test]
fn test_inspect_tallying_leading() {
    let output = run_prints(&p(
        r#"
01 WS-STR PIC X(10) VALUE "0000123450".
01 WS-CNT PIC 9(3) VALUE 0.
"#,
        r#"
    INSPECT WS-STR TALLYING WS-CNT FOR LEADING "0".
    DISPLAY WS-CNT.
"#,
    ));
    assert_eq!(output, vec!["004"]);
}

#[test]
fn test_inspect_replacing_leading() {
    let output = run_prints(&p(
        "01 WS-STR PIC X(6) VALUE \"004200\".",
        r#"
    INSPECT WS-STR REPLACING LEADING "0" BY " ".
    DISPLAY WS-STR.
"#,
    ));
    assert_eq!(output, vec!["  4200"]);
}

#[test]
fn test_inspect_replacing_characters() {
    let output = run_prints(&p(
        "01 WS-STR PIC X(6) VALUE \"SECRET\".",
        r#"
    INSPECT WS-STR REPLACING CHARACTERS BY "*".
    DISPLAY WS-STR.
"#,
    ));
    assert_eq!(output, vec!["******"]);
}

#[test]
fn test_inspect_converting() {
    let output = run_prints(&p(
        "01 WS-STR PIC X(5) VALUE \"hello\".",
        r#"
    INSPECT WS-STR CONVERTING "aeiou" TO "AEIOU".
    DISPLAY WS-STR.
"#,
    ));
    assert_eq!(output, vec!["hEllO"]);
}

#[test]
fn test_inspect_before_initial() {
    compile_ok(&p(
        r#"
01 WS-STR PIC X(10) VALUE "ABC,DEF,GHI".
01 WS-CNT PIC 9(3) VALUE 0.
"#,
        r#"
    INSPECT WS-STR TALLYING WS-CNT FOR ALL "D" BEFORE INITIAL ",".
"#,
    ));
}
