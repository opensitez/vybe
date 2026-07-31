use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_evaluate_range() {
    let output = run_prints(&p(
        "01 WS-VAL PIC 9 VALUE 3.",
        r#"
    EVALUATE WS-VAL
        WHEN 1 THRU 5
            DISPLAY "LOW"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["LOW"]);
}

#[test]
fn test_evaluate_not_value() {
    let output = run_prints(&p(
        "01 WS-VAL PIC 9 VALUE 4.",
        r#"
    EVALUATE WS-VAL
        WHEN NOT 3
            DISPLAY "NOT-THREE"
        WHEN OTHER
            DISPLAY "THREE"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["NOT-THREE"]);
}

#[test]
fn test_evaluate_multiple_subjects() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9 VALUE 1.
01 WS-B PIC 9 VALUE 2.
"#,
        r#"
    EVALUATE WS-A ALSO WS-B
        WHEN 1 ALSO 2
            DISPLAY "MATCH"
        WHEN 1 ALSO ANY
            DISPLAY "PARTIAL"
        WHEN OTHER
            DISPLAY "NONE"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["MATCH"]);
}

#[test]
fn test_evaluate_false() {
    let output = run_prints(&p(
        "01 WS-A PIC 9 VALUE 5.",
        r#"
    EVALUATE FALSE
        WHEN WS-A = 5
            DISPLAY "NOT-FIVE"
        WHEN OTHER
            DISPLAY "FIVE"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["FIVE"]);
}

#[test]
fn test_evaluate_expr_subject() {
    let output = run_prints(&p(
        r#"
01 WS-A PIC 9 VALUE 1.
01 WS-B PIC 9 VALUE 2.
"#,
        r#"
    EVALUATE WS-A + WS-B
        WHEN 3
            DISPLAY "THREE"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["THREE"]);
}

#[test]
fn test_evaluate_function_subject() {
    let output = run_prints(&p(
        "01 WS-X PIC 9 VALUE 5.",
        r#"
    EVALUATE FUNCTION MOD(WS-X 3)
        WHEN 0
            DISPLAY "DIV3"
        WHEN 2
            DISPLAY "REM2"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["REM2"]);
}

#[test]
fn test_evaluate_88_subject() {
    let output = run_prints(&p(
        r#"
01 WS-FLAG PIC 9 VALUE 1.
   88 IS-ACTIVE VALUE 1.
"#,
        r#"
    EVALUATE IS-ACTIVE
        WHEN TRUE
            DISPLAY "ACTIVE"
        WHEN FALSE
            DISPLAY "INACTIVE"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["ACTIVE"]);
}

#[test]
fn test_evaluate_grading_ranges() {
    let output = run_prints(&p(
        "01 WS-SCORE PIC 9(3) VALUE 85.",
        r#"
    EVALUATE TRUE
        WHEN WS-SCORE >= 90
            DISPLAY "A"
        WHEN WS-SCORE >= 80
            DISPLAY "B"
        WHEN WS-SCORE >= 70
            DISPLAY "C"
        WHEN OTHER
            DISPLAY "F"
    END-EVALUATE.
"#,
    ));
    assert_eq!(output, vec!["B"]);
}
