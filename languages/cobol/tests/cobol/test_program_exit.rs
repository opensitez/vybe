use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn test_exit_stop_run() {
    let output = run_prints(&p(
        "",
        r#"
    DISPLAY "START".
    STOP RUN.
    DISPLAY "END".
"#,
    ));
    assert_eq!(output, vec!["START"]);
}

#[test]
fn test_exit_goback() {
    let output = run_prints(&p(
        "",
        r#"
    DISPLAY "START".
    GOBACK.
    DISPLAY "END".
"#,
    ));
    assert_eq!(output, vec!["START"]);
}

#[test]
fn test_exit_program_verb() {
    compile_ok(&p(
        "",
        r#"
    EXIT PROGRAM.
"#,
    ));
}

#[test]
fn test_exit_perform_cycle() {
    compile_ok(&p(
        "01 WS-I PIC 9.",
        r#"
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 5
        IF WS-I = 3
            EXIT PERFORM CYCLE
        END-IF
        DISPLAY WS-I
    END-PERFORM.
"#,
    ));
}

#[test]
fn test_exit_section_paragraph() {
    compile_ok(&p(
        "",
        r#"
    PERFORM MY-SEC.
    STOP RUN.
MY-SEC SECTION.
MY-PARA.
    DISPLAY "PARA".
    EXIT PARAGRAPH.
    DISPLAY "NOT-SHOWN".
MY-EXIT.
    EXIT SECTION.
"#,
    ));
}
