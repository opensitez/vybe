use super::helpers::{compile_ok, run_prints};

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn accept_from_date_compiles() {
    compile_ok(&p(
        "01 TODAY PIC 9(6).",
        "    ACCEPT TODAY FROM DATE.",
    ));
}

#[test]
fn accept_from_date_yyyymmdd_compiles() {
    compile_ok(&p(
        "01 TODAY PIC 9(8).",
        "    ACCEPT TODAY FROM DATE YYYYMMDD.",
    ));
}

#[test]
fn accept_from_time_compiles() {
    compile_ok(&p(
        "01 NOW PIC 9(8).",
        "    ACCEPT NOW FROM TIME.",
    ));
}

#[test]
fn accept_from_day_compiles() {
    compile_ok(&p(
        "01 DAY-OF-YEAR PIC 9(5).",
        "    ACCEPT DAY-OF-YEAR FROM DAY.",
    ));
}

#[test]
fn accept_from_day_of_week_compiles() {
    compile_ok(&p(
        "01 DOW PIC 9.",
        "    ACCEPT DOW FROM DAY-OF-WEEK.",
    ));
}

#[test]
fn accept_from_day_yyyyddd_compiles() {
    compile_ok(&p(
        "01 D PIC 9(7).",
        "    ACCEPT D FROM DAY YYYYDDD.",
    ));
}

#[test]
fn accept_from_console_compiles() {
    compile_ok(&p(
        "01 S PIC X(20).",
        "    ACCEPT S FROM CONSOLE.",
    ));
}

#[test]
fn accept_from_command_line_compiles() {
    compile_ok(&p(
        "01 ARG PIC X(80).",
        "    ACCEPT ARG FROM COMMAND-LINE.",
    ));
}

#[test]
fn accept_multiple_fields_from_date() {
    compile_ok(&p(
        "01 D1 PIC 9(6).\n01 D2 PIC 9(8).",
        "    ACCEPT D1 FROM DATE.\n    ACCEPT D2 FROM DATE YYYYMMDD.",
    ));
}

#[test]
fn accept_default_from_stdin_compiles() {
    compile_ok(&p(
        "01 INPUT-LINE PIC X(80).",
        "    ACCEPT INPUT-LINE.",
    ));
}

#[test]
fn stop_run_terminates_program() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    STOP RUN.",
    );
}

#[test]
fn stop_run_after_display() {
    let out = run_prints(&p(
        "",
        "    DISPLAY \"BEFORE STOP\".\n    STOP RUN.\n    DISPLAY \"UNREACHABLE\".",
    ));
    assert_eq!(out, vec!["BEFORE STOP"]);
}

#[test]
fn stop_literal_compiles() {
    compile_ok(&p(
        "",
        "    STOP \"PAUSE MESSAGE\".",
    ));
}

#[test]
fn exit_program_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    EXIT PROGRAM.",
    );
}

#[test]
fn goback_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nPROCEDURE DIVISION.\n    GOBACK.",
    );
}

#[test]
fn exit_paragraph_compiles() {
    compile_ok(&p(
        "",
        r#"    PERFORM MY-PARA.
    STOP RUN.
MY-PARA.
    DISPLAY "PARA".
    EXIT."#,
    ));
}

#[test]
fn stop_run_in_conditional() {
    let out = run_prints(&p(
        "01 FLAG PIC 9 VALUE 1.",
        "    IF FLAG = 0\n        STOP RUN\n    END-IF.\n    DISPLAY \"CONTINUED\".",
    ));
    assert_eq!(out, vec!["CONTINUED"]);
}

#[test]
fn exit_in_last_paragraph_compiles() {
    compile_ok(&p(
        "",
        r#"    PERFORM LAST-PARA.
    STOP RUN.
LAST-PARA.
    DISPLAY "LAST".
    EXIT."#,
    ));
}

#[test]
fn goback_in_subprogram_context_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. SUB.\nPROCEDURE DIVISION.\n    DISPLAY \"SUB\".\n    GOBACK.",
    );
}

#[test]
fn stop_run_after_loop() {
    let out = run_prints(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM UNTIL I >= 3\n        ADD 1 TO I\n    END-PERFORM.\n    DISPLAY I.\n    STOP RUN.",
    ));
    assert_eq!(out, vec!["3"]);
}

#[test]
fn accept_from_date_used_in_display() {
    compile_ok(&p(
        "01 TODAY PIC 9(6).",
        "    ACCEPT TODAY FROM DATE.\n    DISPLAY TODAY.",
    ));
}

#[test]
fn stop_run_inside_perform_paragraph() {
    compile_ok(&p(
        "",
        r#"    PERFORM DO-WORK.
    STOP RUN.
DO-WORK.
    DISPLAY "WORKING".
    STOP RUN."#,
    ));
}

#[test]
fn exit_section_compiles() {
    compile_ok(&p(
        "",
        r#"    PERFORM MY-SECTION.
    STOP RUN.
MY-SECTION SECTION.
    DISPLAY "IN SECTION".
    EXIT SECTION."#,
    ));
}

#[test]
fn accept_from_environment_variable() {
    compile_ok(&p(
        "01 ENV-VAL PIC X(80).",
        "    ACCEPT ENV-VAL FROM ENVIRONMENT \"PATH\".",
    ));
}

#[test]
fn stop_run_with_return_code() {
    compile_ok(&p(
        "01 RETURN-CODE PIC 9(4).",
        "    MOVE 0 TO RETURN-CODE.\n    STOP RUN.",
    ));
}

#[test]
fn accept_numeric_from_date() {
    compile_ok(&p(
        "01 YEAR PIC 9(4).\n01 FULL-DATE PIC 9(8).",
        "    ACCEPT FULL-DATE FROM DATE YYYYMMDD.\n    MOVE FULL-DATE(1:4) TO YEAR.",
    ));
}

#[test]
fn exit_in_inline_loop_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3\n        IF I = 2\n            CONTINUE\n        END-IF\n    END-PERFORM.",
    ));
}

#[test]
fn goback_sequence_in_main_program() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. MAIN.\nPROCEDURE DIVISION.\n    DISPLAY \"MAIN\".\n    GOBACK.",
    );
}

#[test]
fn accept_group_item_from_console() {
    compile_ok(&p(
        "01 RESPONSE.\n   05 CODE PIC X.\n   05 DETAIL PIC X(10).",
        "    ACCEPT RESPONSE FROM CONSOLE.",
    ));
}
