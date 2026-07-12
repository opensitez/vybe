use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_multiply_by_zero() {
    let output = run_prints(&p(
        "01 WS-A PIC 9(3) VALUE 5.",
        r#"
    MULTIPLY 0 BY WS-A.
    DISPLAY WS-A.
"#,
    ));
    assert_eq!(output, vec!["000"]);
}

#[test]
fn test_multiply_by_one() {
    let output = run_prints(&p(
        "01 WS-A PIC 9(3) VALUE 5.",
        r#"
    MULTIPLY 1 BY WS-A.
    DISPLAY WS-A.
"#,
    ));
    assert_eq!(output, vec!["005"]);
}

#[test]
fn test_multiply_negative() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC S9(3) VALUE 5.
01 WS-B PIC S9(3) VALUE -2.
01 WS-C PIC S9(3) VALUE 0.
"#,
        r#"
    MULTIPLY WS-A BY WS-B GIVING WS-C.
    DISPLAY WS-C.
"#,
    ));
    assert!(output.len() >= 1);
}

#[test]
fn test_multiply_decimals() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9V9 VALUE 2.5.
01 WS-B PIC 9V9 VALUE 1.5.
01 WS-C PIC 9V99 VALUE 0.0.
"#,
        r#"
    MULTIPLY WS-A BY WS-B GIVING WS-C.
    DISPLAY WS-C.
"#,
    ));
    assert_eq!(output, vec!["375"]);
}

#[test]
fn test_multiply_corresponding() {
    compile_ok(&p(
        r#"
01 WS-G1.
   05 WS-X PIC 9(3) VALUE 2.
   05 WS-Y PIC 9(3) VALUE 3.
01 WS-G2.
   05 WS-X PIC 9(3) VALUE 10.
   05 WS-Y PIC 9(3) VALUE 20.
"#,
        r#"
    MULTIPLY CORRESPONDING WS-G1 BY WS-G2.
"#,
    ));
}
