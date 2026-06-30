use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_value_clause_zero() {
    let output = run_prints(&p(
        r#"
01 WS-NUM1 PIC 9(3) VALUE ZERO.
01 WS-NUM2 PIC 9(3) VALUE ZEROS.
01 WS-NUM3 PIC 9(3) VALUE ZEROES.
"#,
        r#"
    DISPLAY WS-NUM1.
    DISPLAY WS-NUM2.
    DISPLAY WS-NUM3.
"#,
    ));
    assert_eq!(output, vec!["000", "000", "000"]);
}

#[test]
fn test_value_clause_spaces() {
    let output = run_prints(&p(
        r#"
01 WS-STR1 PIC X(5) VALUE SPACE.
01 WS-STR2 PIC X(5) VALUE SPACES.
"#,
        r#"
    DISPLAY WS-STR1.
    DISPLAY WS-STR2.
"#,
    ));
    assert_eq!(output, vec!["     ", "     "]);
}

#[test]
fn test_value_clause_all_literal() {
    let output = run_prints(&p(
        r#"
01 WS-STR1 PIC X(5) VALUE ALL "*".
01 WS-STR2 PIC X(6) VALUE ALL "AB".
"#,
        r#"
    DISPLAY WS-STR1.
    DISPLAY WS-STR2.
"#,
    ));
    assert_eq!(output, vec!["*****", "ABABAB"]);
}

#[test]
fn test_value_clause_decimals_negatives() {
    let output = run_prints(&p(
        r#"
01 WS-DEC PIC 9V99 VALUE 3.14.
01 WS-NEG PIC S9(3) VALUE -100.
"#,
        r#"
    DISPLAY WS-DEC.
    DISPLAY WS-NEG.
"#,
    ));
    assert!(output.len() >= 2);
}

#[test]
fn test_value_clause_group_propagation() {
    compile_ok(&p(
        r#"
01 WS-GROUP VALUE "ABCDEF".
   05 WS-A PIC X(3).
   05 WS-B PIC X(3).
"#,
        r#"
    DISPLAY WS-GROUP.
"#,
    ));
}

#[test]
fn test_value_clause_no_value_implicit() {
    let output = run_prints(&p(
        r#"
01 WS-NUM PIC 9(3).
01 WS-STR PIC X(3).
"#,
        r#"
    DISPLAY WS-NUM.
    DISPLAY WS-STR.
"#,
    ));
    // Implicit values are typically 0/spaces in many COBOL implementations, check compilation and displaying
    assert_eq!(output.len(), 2);
}

#[test]
fn test_value_clause_justified_right() {
    let output = run_prints(&p(
        r#"
01 WS-STR PIC X(10) VALUE "HELLO" JUSTIFIED RIGHT.
"#,
        r#"
    DISPLAY WS-STR.
"#,
    ));
    assert_eq!(output, vec!["     HELLO"]);
}

#[test]
fn test_value_clause_occurs() {
    compile_ok(&p(
        r#"
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 5 TIMES VALUE 100.
"#,
        r#"
    DISPLAY WS-ITEM(1).
"#,
    ));
}
