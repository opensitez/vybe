use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_sections_fallthrough() {
    let output = run_prints(&p(
        "",
        r#"
    PERFORM SEC1.
    STOP RUN.
SEC1 SECTION.
PARA1.
    DISPLAY "SEC1-P1".
PARA2.
    DISPLAY "SEC1-P2".
SEC2 SECTION.
PARA3.
    DISPLAY "SEC2-P3".
"#,
    ));
    assert_eq!(output, vec!["SEC1-P1", "SEC1-P2"]);
}

#[test]
fn test_sections_perform_thru() {
    let output = run_prints(&p(
        "",
        r#"
    PERFORM SEC1 THRU SEC2.
    STOP RUN.
SEC1 SECTION.
    DISPLAY "S1".
SEC2 SECTION.
    DISPLAY "S2".
"#,
    ));
    assert_eq!(output, vec!["S1", "S2"]);
}

#[test]
fn test_declaratives_block() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. DECLPROG.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT FILE1 ASSIGN TO "dummy".
DATA DIVISION.
FILE SECTION.
FD FILE1.
01 REC PIC X(10).
PROCEDURE DIVISION.
DECLARATIVES.
ERR-HANDLER SECTION.
    USE AFTER STANDARD ERROR PROCEDURE ON FILE1.
ERR-PARA.
    DISPLAY "ERROR OCCURRED".
END DECLARATIVES.
MAIN SECTION.
    OPEN INPUT FILE1.
    STOP RUN.
"#,
    );
}
