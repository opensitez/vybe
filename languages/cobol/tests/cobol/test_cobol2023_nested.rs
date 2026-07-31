use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// COBOL 2023: Nested programs, COPY, and modular features
// ═══════════════════════════════════════════════════════════

#[test]
fn nested_program_basic() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-PROG.
PROCEDURE DIVISION.
    DISPLAY "Main program".
    STOP RUN.
IDENTIFICATION DIVISION.
PROGRAM-ID. HELPER.
PROCEDURE DIVISION.
    DISPLAY "Helper".
    STOP RUN.
END PROGRAM HELPER.
END PROGRAM MAIN-PROG.
"#,
    );
}

#[test]
fn nested_program_with_data() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. OUTER.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SHARED PIC X(20) VALUE "Shared".
PROCEDURE DIVISION.
    DISPLAY WS-SHARED.
    STOP RUN.
IDENTIFICATION DIVISION.
PROGRAM-ID. INNER.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-LOCAL PIC X(20) VALUE "Local".
PROCEDURE DIVISION.
    DISPLAY WS-LOCAL.
    STOP RUN.
END PROGRAM INNER.
END PROGRAM OUTER.
"#,
    );
}

#[test]
fn copy_basic() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    COPY STANDARD-HEADER.
    DISPLAY "After copy".
    STOP RUN.
"#,
    );
}

#[test]
fn copy_replacing() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    COPY CUSTOMER-REC REPLACING ==OLD-NAME== BY ==CUST-NAME==.
    DISPLAY "After copy replacing".
    STOP RUN.
"#,
    );
}

#[test]
fn copy_of_library() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    COPY DATE-UTILS OF COMMON-LIB.
    DISPLAY "Copy from library".
    STOP RUN.
"#,
    );
}

#[test]
fn paragraphs_as_modules() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TOTAL PIC 9(10) VALUE 0.
01 WS-TAX PIC 9(5) VALUE 500.
01 WS-SUBTOTAL PIC 9(7) VALUE 10000.
PROCEDURE DIVISION.
    PERFORM CALCULATE-TOTAL.
    PERFORM DISPLAY-RESULT.
    STOP RUN.
CALCULATE-TOTAL.
    ADD WS-SUBTOTAL TO WS-TOTAL.
    ADD WS-TAX TO WS-TOTAL.
DISPLAY-RESULT.
    DISPLAY "Total: " WS-TOTAL.
"#,
    );
}

#[test]
fn perform_thru_sections() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNTER PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM INIT-PARA THRU CLEANUP-PARA.
    DISPLAY WS-COUNTER.
    STOP RUN.
INIT-PARA.
    ADD 1 TO WS-COUNTER.
PROCESS-PARA.
    ADD 10 TO WS-COUNTER.
CLEANUP-PARA.
    ADD 100 TO WS-COUNTER.
"#,
    );
}

#[test]
fn exit_section() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM WORK-PARA.
    DISPLAY "Done".
    STOP RUN.
WORK-PARA.
    DISPLAY "Working".
    EXIT SECTION.
    DISPLAY "Never reached".
"#,
    );
}

#[test]
fn exit_method() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
CLASS-ID. MY-CLASS.
OBJECT.
METHOD-ID. PROCESS.
PROCEDURE DIVISION.
    DISPLAY "Start".
    EXIT METHOD.
    DISPLAY "Never reached".
END METHOD PROCESS.
END OBJECT.
END CLASS MY-CLASS.
"#,
    );
}

#[test]
fn global_data_item() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SHARED PIC X(20) GLOBAL.
01 WS-LOCAL PIC X(20).
PROCEDURE DIVISION.
    MOVE "Global value" TO WS-SHARED.
    DISPLAY WS-SHARED.
    STOP RUN.
"#,
    );
}

#[test]
fn external_data_item() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-EXT PIC X(20) EXTERNAL.
PROCEDURE DIVISION.
    DISPLAY WS-EXT.
    STOP RUN.
"#,
    );
}

#[test]
fn local_storage_section() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PERSIST PIC 9(5) VALUE 100.
LOCAL-STORAGE SECTION.
01 WS-LOCAL PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    ADD WS-PERSIST TO WS-LOCAL.
    DISPLAY WS-LOCAL.
    STOP RUN.
"#,
    );
}

#[test]
fn linkage_section() {
    compile_ok(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-MAIN PIC X(20) VALUE "Main".
LINKAGE SECTION.
01 LS-PARAM PIC X(20).
PROCEDURE DIVISION.
    DISPLAY WS-MAIN.
    STOP RUN.
"#,
    );
}

#[test]
fn nested_programs_runtime_display() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-PROG.
PROCEDURE DIVISION.
    DISPLAY "Main program".
    END PROGRAM MAIN-PROG.
"#,
    );
    assert_eq!(out, vec!["Main program"]);
}

#[test]
fn nested_with_inner_call_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. OUTER.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SHARED PIC X(20) VALUE "Shared".
PROCEDURE DIVISION.
    DISPLAY WS-SHARED.
    CALL "INNER"
    END-CALL
    STOP RUN.
IDENTIFICATION DIVISION.
PROGRAM-ID. INNER.
PROCEDURE DIVISION.
    DISPLAY "Inner".
    STOP RUN.
END PROGRAM INNER.
END PROGRAM OUTER.
"#,
    );
    assert_eq!(out, vec!["Shared", "Inner"]);
}

#[test]
fn perform_thru_sections_runtime_total() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNTER PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM INIT-PARA THRU CLEANUP-PARA.
    DISPLAY WS-COUNTER.
    STOP RUN.
INIT-PARA.
    ADD 1 TO WS-COUNTER.
PROCESS-PARA.
    ADD 10 TO WS-COUNTER.
CLEANUP-PARA.
    ADD 100 TO WS-COUNTER.
"#,
    );
    assert_eq!(out, vec!["111"]);
}

#[test]
fn global_data_item_runtime() {
    let out = run_prints(
        r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SHARED PIC X(20) GLOBAL VALUE "GLOBAL VALUE".
PROCEDURE DIVISION.
    DISPLAY WS-SHARED.
    STOP RUN.
"#,
    );
    assert_eq!(out, vec!["GLOBAL VALUE"]);
}
