use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_condition_compound_and() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9 VALUE 5.
01 WS-B PIC 9 VALUE 10.
01 WS-C PIC 9 VALUE 15.
"#,
        r#"
    IF WS-A > 0 AND WS-B > 5 AND WS-C > 10
        DISPLAY "ALL-TRUE"
    ELSE
        DISPLAY "SOME-FALSE"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["ALL-TRUE"]);
}

#[test]
fn test_condition_compound_or() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9 VALUE 0.
01 WS-B PIC 9 VALUE 10.
"#,
        r#"
    IF WS-A > 5 OR WS-B > 5
        DISPLAY "AT-LEAST-ONE"
    ELSE
        DISPLAY "NONE"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["AT-LEAST-ONE"]);
}

#[test]
fn test_condition_negated_parens() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9 VALUE 5.
01 WS-B PIC 9 VALUE 0.
"#,
        r#"
    IF NOT (WS-A > 0 AND WS-B > 0)
        DISPLAY "NEGATED"
    ELSE
        DISPLAY "NORMAL"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["NEGATED"]);
}

#[test]
fn test_condition_class_numeric() {
    let output = run_prints(&p(
        r#"
01 WS-TXT1 PIC X(3) VALUE "123".
01 WS-TXT2 PIC X(3) VALUE "ABC".
"#,
        r#"
    IF WS-TXT1 IS NUMERIC
        DISPLAY "TXT1-NUM"
    END-IF.
    IF WS-TXT2 IS NOT NUMERIC
        DISPLAY "TXT2-NOT-NUM"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["TXT1-NUM", "TXT2-NOT-NUM"]);
}

#[test]
fn test_condition_class_alphabetic() {
    let output = run_prints(&p(
        r#"
01 WS-TXT PIC X(3) VALUE "ABC".
"#,
        r#"
    IF WS-TXT IS ALPHABETIC
        DISPLAY "ALPHA"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["ALPHA"]);
}

#[test]
fn test_condition_sign() {
    let output = run_prints(&p(
        r#"
01 WS-POS PIC S9(3) VALUE 10.
01 WS-NEG PIC S9(3) VALUE -10.
01 WS-ZER PIC S9(3) VALUE 0.
"#,
        r#"
    IF WS-POS IS POSITIVE
        DISPLAY "POS"
    END-IF.
    IF WS-NEG IS NEGATIVE
        DISPLAY "NEG"
    END-IF.
    IF WS-ZER IS ZERO
        DISPLAY "ZER"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["POS", "NEG", "ZER"]);
}

#[test]
fn test_condition_abbreviated() {
    let output = run_prints(&p(
        "01 WS-A PIC 9 VALUE 5.",
        r#"
    IF WS-A > 0 AND < 10
        DISPLAY "BETWEEN"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["BETWEEN"]);
}

#[test]
fn test_condition_figuratives() {
    let output = run_prints(&p(
        r#"
01 WS-NAME PIC X(5) VALUE SPACES.
01 WS-VAL PIC 9(3) VALUE ZEROS.
"#,
        r#"
    IF WS-NAME = SPACES
        DISPLAY "NAME-SPACES"
    END-IF.
    IF WS-VAL = ZEROS
        DISPLAY "VAL-ZEROS"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["NAME-SPACES", "VAL-ZEROS"]);
}

#[test]
fn test_condition_with_nested_parentheses_precedence() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9 VALUE 1.
01 WS-B PIC 9 VALUE 8.
01 WS-C PIC 9 VALUE 9.
"#,
        r#"
    IF (WS-A > 0 AND WS-B > 5) OR WS-C > 10
        DISPLAY "COND"
    ELSE
        DISPLAY "NONE"
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["COND"]);
}

#[test]
fn test_condition_nested_if_else_chain() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9 VALUE 2.
"#,
        r#"
    IF WS-A > 5
        DISPLAY "BIG"
    ELSE
        IF WS-A = 2
            DISPLAY "TWO"
        ELSE
            DISPLAY "OTHER"
        END-IF
    END-IF.
"#,
    ));
    assert_eq!(output, vec!["TWO"]);
}

#[test]
fn condition_with_truthy_falsy_literals_compile() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 FLAG PIC X VALUE "Y".
PROCEDURE DIVISION.
    IF FLAG = "Y"
        DISPLAY "YES"
    ELSE
        DISPLAY "NO"
    END-IF.
    STOP RUN.
"#,
    );
}
