use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_add_multiple_literals() {
    let output = run_prints(&p(
        "01 WS-A PIC 9(3) VALUE 0.",
        r#"
    ADD 1 2 3 TO WS-A.
    DISPLAY WS-A.
"#,
    ));
    assert_eq!(output, vec!["006"]);
}

#[test]
fn test_add_giving_multiple() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(3) VALUE 5.
01 WS-B PIC 9(3) VALUE 0.
01 WS-C PIC 9(3) VALUE 0.
"#,
        r#"
    ADD 10 TO WS-A GIVING WS-B WS-C.
    DISPLAY WS-B.
    DISPLAY WS-C.
"#,
    ));
    assert_eq!(output, vec!["015", "015"]);
}

#[test]
fn test_add_decimals() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(3)V99 VALUE 1.25.
01 WS-B PIC 9(3)V99 VALUE 2.50.
01 WS-C PIC 9(3)V99 VALUE 0.0.
"#,
        r#"
    ADD WS-A TO WS-B GIVING WS-C.
    DISPLAY WS-C.
"#,
    ));
    assert_eq!(output, vec!["00375"]);
}

#[test]
fn test_add_negative_literal() {
    let output = run_prints(&p(
        "01 WS-A PIC S9(3) VALUE 10.",
        r#"
    ADD -5 TO WS-A.
    DISPLAY WS-A.
"#,
    ));
    assert_eq!(output, vec!["005"]);
}

#[test]
fn test_add_rounded() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9V99 VALUE 1.25.
01 WS-B PIC 9V9 VALUE 0.0.
"#,
        r#"
    ADD WS-A TO WS-B ROUNDED.
    DISPLAY WS-B.
"#,
    ));
    assert_eq!(output, vec!["013"]);
}

#[test]
fn test_add_corresponding() {
    let output = run_prints(&p(
        r#"
01 WS-G1.
   05 WS-X PIC 9(3) VALUE 10.
   05 WS-Y PIC 9(3) VALUE 20.
01 WS-G2.
   05 WS-X PIC 9(3) VALUE 5.
   05 WS-Y PIC 9(3) VALUE 15.
"#,
        r#"
    ADD CORRESPONDING WS-G1 TO WS-G2.
    DISPLAY WS-G2.
"#,
    ));
    assert_eq!(output, vec!["1535"]);
}

#[test]
fn test_add_size_error_false_path() {
    let output = run_prints(&p(
        "01 WS-A PIC 9(3) VALUE 10.\n01 WS-B PIC 99 VALUE 0.",
        r#"
    ADD WS-A TO WS-B ON SIZE ERROR
        DISPLAY "ERR"
    NOT ON SIZE ERROR
        DISPLAY WS-B
    END-ADD.
"#,
    ));
    assert_eq!(output, vec!["10"]);
}

#[test]
fn test_add_giving_and_size_error() {
    compile_ok(&p(
        "01 WS-A PIC 9(5) VALUE 99999.\n01 WS-B PIC 9(5) VALUE 0.",
        r#"
    ADD WS-A TO WS-A GIVING WS-B ON SIZE ERROR
        DISPLAY WS-B
    END-ADD.
"#,
    ));
}

#[test]
fn test_add_multiple_receiving_fields_runtime() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9(2) VALUE 10.
01 WS-B PIC 9(2) VALUE 20.
01 WS-C PIC 9(2) VALUE 0.
01 WS-D PIC 9(2) VALUE 0.
"#,
        r#"
    ADD WS-A WS-B TO WS-C GIVING WS-D.
    DISPLAY WS-D.
"#,
    ));
    assert_eq!(output, vec!["030"]);
}
