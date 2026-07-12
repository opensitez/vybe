use super::helpers::compile_ok;

fn p(data: &str, body: &str) -> String {
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.",
        data, body
    )
}

#[test]
fn declaratives_use_after_error_compiles() {
    compile_ok(&p(
        "",
        "    DECLARATIVES.\n    D-SEC SECTION.\n        USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.\n    END DECLARATIVES.\n    DISPLAY \"RUN\".",
    ));
}
#[test]
fn cics_delay_statement_compiles() {
    compile_ok(&p("", "    EXEC CICS DELAY SECONDS(1) END-EXEC."));
}
#[test]
fn cics_start_statement_compiles() {
    compile_ok(&p("", "    EXEC CICS START TRANSID(NXTT) END-EXEC."));
}
#[test]
fn xml_parse_processing_procedure_compiles() {
    compile_ok(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n01 X PIC X(200) VALUE \"<a>1</a>\".\nPROCEDURE DIVISION.\n    XML PARSE X PROCESSING PROCEDURE P-H.\n    STOP RUN.\nP-H SECTION.\n    DISPLAY \"H\".",
    );
}
#[test]
fn perform_loop_with_evaluate_compiles() {
    compile_ok(&p(
        "01 N PIC 9 VALUE 0.",
        "    PERFORM UNTIL N >= 2\n        ADD 1 TO N\n        EVALUATE N\n            WHEN 1 DISPLAY \"A\"\n            WHEN 2 DISPLAY \"B\"\n        END-EVALUATE\n    END-PERFORM.",
    ));
}
#[test]
fn if_else_branching_compiles() {
    compile_ok(&p(
        "01 F PIC 9 VALUE 1.",
        "    IF F = 1 DISPLAY \"Y\" ELSE DISPLAY \"N\" END-IF.",
    ));
}
#[test]
fn call_on_exception_branch_compiles() {
    compile_ok(&p(
        "",
        "    CALL \"SUBX\"\n        ON EXCEPTION DISPLAY \"E\"\n        NOT ON EXCEPTION DISPLAY \"O\"\n    END-CALL.",
    ));
}
#[test]
fn nested_perform_blocks_compiles() {
    compile_ok(&p(
        "01 I PIC 9 VALUE 0.\n01 J PIC 9 VALUE 0.",
        "    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 2\n        PERFORM VARYING J FROM 1 BY 1 UNTIL J > 2\n            DISPLAY I\n        END-PERFORM\n    END-PERFORM.",
    ));
}
