use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_subtract_multiple() {
    let output = run_prints(&p(
        "01 WS-X PIC 9(3) VALUE 20.",
        r#"
    SUBTRACT 3 4 FROM WS-X.
    DISPLAY WS-X.
"#,
    ));
    assert_eq!(output, vec!["013"]);
}

#[test]
fn test_subtract_giving_zero() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 9(3) VALUE 0.
"#,
        r#"
    SUBTRACT WS-A FROM WS-A GIVING WS-B.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["000"]);
}

#[test]
fn test_subtract_decimals() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9V99 VALUE 5.50.
01 WS-B PIC 9V99 VALUE 1.25.
01 WS-C PIC 9V99 VALUE 0.00.
"#,
        r#"
    SUBTRACT WS-B FROM WS-A GIVING WS-C.
    DISPLAY WS-C.
"#,
    ));
    assert_eq!(output, vec!["425"]);
}

#[test]
fn test_subtract_corresponding() {
    let output = run_prints(&p(
        r#"
01 WS-G1.
   05 WS-X PIC 9(3) VALUE 5.
   05 WS-Y PIC 9(3) VALUE 10.
01 WS-G2.
   05 WS-X PIC 9(3) VALUE 20.
   05 WS-Y PIC 9(3) VALUE 30.
"#,
        r#"
    SUBTRACT CORRESPONDING WS-G1 FROM WS-G2.
    DISPLAY WS-X.
    DISPLAY WS-Y.
"#,
    ));
    assert_eq!(output, vec!["015", "020"]);
}
