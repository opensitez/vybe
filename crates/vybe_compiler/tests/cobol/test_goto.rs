use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_goto_basic() {
    let output = run_prints(&p(
        "",
        r#"
    GO TO PARA-TARGET.
    DISPLAY "SKIPPED".
PARA-TARGET.
    DISPLAY "TARGET".
"#,
    ));
    assert_eq!(output, vec!["TARGET"]);
}

#[test]
fn test_goto_depending_on() {
    compile_ok(&p(
        "01 WS-IDX PIC 9 VALUE 2.",
        r#"
    GO TO PARA1 PARA2 PARA3 DEPENDING ON WS-IDX.
    DISPLAY "OTHER".
    STOP RUN.
PARA1.
    DISPLAY "P1".
PARA2.
    DISPLAY "P2".
PARA3.
    DISPLAY "P3".
"#,
    ));
}

#[test]
fn test_goto_depending_on_out_of_bounds() {
    let output = run_prints(&p(
        "01 WS-IDX PIC 9 VALUE 0.",
        r#"
    GO TO PARA1 DEPENDING ON WS-IDX.
    DISPLAY "FALLTHROUGH".
    STOP RUN.
PARA1.
    DISPLAY "P1".
"#,
    ));
    assert_eq!(output, vec!["FALLTHROUGH"]);
}

#[test]
fn test_goto_exit_paragraph() {
    compile_ok(&p(
        "",
        r#"
    PERFORM MY-PARA.
    STOP RUN.
MY-PARA.
    DISPLAY "START".
    GO TO MY-PARA-EXIT.
    DISPLAY "SKIPPED".
MY-PARA-EXIT.
    EXIT.
"#,
    ));
}
