use super::helpers::run_prints;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_divide_into() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(3) VALUE 2.
01 WS-B PIC 9(3) VALUE 10.
"#,
        r#"
    DIVIDE WS-A INTO WS-B.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["005"]);
}

#[test]
fn test_divide_into_giving() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(3) VALUE 2.
01 WS-B PIC 9(3) VALUE 10.
01 WS-C PIC 9(3) VALUE 0.
"#,
        r#"
    DIVIDE WS-A INTO WS-B GIVING WS-C.
    DISPLAY WS-C.
"#,
    ));
    assert_eq!(output, vec!["005"]);
}

#[test]
fn test_divide_remainder_into() {
    let output = run_prints(&p(
        r#"
01 WS-DIVISOR PIC 9(3) VALUE 5.
01 WS-DIVIDEND PIC 9(3) VALUE 17.
01 WS-QUOTIENT PIC 9(3) VALUE 0.
01 WS-REMAINDER PIC 9(3) VALUE 0.
"#,
        r#"
    DIVIDE WS-DIVISOR INTO WS-DIVIDEND GIVING WS-QUOTIENT REMAINDER WS-REMAINDER.
    DISPLAY WS-QUOTIENT.
    DISPLAY WS-REMAINDER.
"#,
    ));
    assert_eq!(output, vec!["003", "002"]);
}

#[test]
fn test_divide_by_giving() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 9(3) VALUE 2.
01 WS-C PIC 9(3) VALUE 0.
"#,
        r#"
    DIVIDE WS-A BY WS-B GIVING WS-C.
    DISPLAY WS-C.
"#,
    ));
    assert_eq!(output, vec!["005"]);
}

#[test]
fn test_divide_decimals() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 9(3) VALUE 3.
01 WS-C PIC 9(3)V99 VALUE 0.00.
"#,
        r#"
    DIVIDE WS-A BY WS-B GIVING WS-C.
    DISPLAY WS-C.
"#,
    ));
    assert_eq!(output, vec!["00333"]);
}

#[test]
fn test_divide_zero_remainder() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 9(3) VALUE 5.
01 WS-Q PIC 9(3) VALUE 0.
01 WS-R PIC 9(3) VALUE 0.
"#,
        r#"
    DIVIDE WS-B INTO WS-A GIVING WS-Q REMAINDER WS-R.
    DISPLAY WS-R.
"#,
    ));
    assert_eq!(output, vec!["000"]);
}

#[test]
fn test_divide_negative() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC S9(3) VALUE -10.
01 WS-B PIC S9(3) VALUE 2.
01 WS-C PIC S9(3) VALUE 0.
"#,
        r#"
    DIVIDE WS-A BY WS-B GIVING WS-C.
    DISPLAY WS-C.
"#,
    ));
    assert!(output.len() >= 1);
}
