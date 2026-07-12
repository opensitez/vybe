use super::helpers::compile_ok;

// ── DECLARATIVES basic structure ──────────────────────────────

#[test]
fn declaratives_section_empty() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-flag PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       error-section SECTION.
           USE AFTER STANDARD ERROR PROCEDURE.
           DISPLAY "error handler".
       END DECLARATIVES.
       main-logic SECTION.
           MOVE "Y" TO ws-flag
           DISPLAY ws-flag
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_use_after_error_on_file() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT my-file ASSIGN TO "test.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD my-file.
       01 my-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-error-msg PIC X(50) VALUE SPACES.
       PROCEDURE DIVISION.
       DECLARATIVES.
       file-error SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON my-file.
           MOVE "File error occurred" TO ws-error-msg
           DISPLAY ws-error-msg.
       END DECLARATIVES.
       main-logic SECTION.
           OPEN INPUT my-file
           CLOSE my-file
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_use_after_exception_input() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT in-file ASSIGN TO "in.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD in-file.
       01 in-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-err PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       in-err SECTION.
           USE AFTER STANDARD EXCEPTION PROCEDURE ON INPUT.
           MOVE "Y" TO ws-err.
       END DECLARATIVES.
       main-para SECTION.
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_use_after_exception_output() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT out-file ASSIGN TO "out.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD out-file.
       01 out-rec PIC X(80).
       WORKING-STORAGE SECTION.
       01 ws-err PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       out-err SECTION.
           USE AFTER STANDARD EXCEPTION PROCEDURE ON OUTPUT.
           MOVE "Y" TO ws-err.
       END DECLARATIVES.
       main-para SECTION.
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_use_after_all_files() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       all-file-errors SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           ADD 1 TO ws-count.
       END DECLARATIVES.
       main-section SECTION.
           DISPLAY ws-count
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_multiple_use_sections() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       ENVIRONMENT DIVISION.
       INPUT-OUTPUT SECTION.
       FILE-CONTROL.
           SELECT file-a ASSIGN TO "a.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
           SELECT file-b ASSIGN TO "b.dat"
               ORGANIZATION IS LINE SEQUENTIAL.
       DATA DIVISION.
       FILE SECTION.
       FD file-a.
       01 rec-a PIC X(80).
       FD file-b.
       01 rec-b PIC X(80).
       PROCEDURE DIVISION.
       DECLARATIVES.
       file-a-error SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON file-a.
           DISPLAY "file-a error".
       file-b-error SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON file-b.
           DISPLAY "file-b error".
       END DECLARATIVES.
       main-para SECTION.
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_with_working_storage_access() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-error-code  PIC 99  VALUE 0.
       01 ws-error-msg   PIC X(40) VALUE SPACES.
       01 ws-error-count PIC 999  VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       error-handler SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           ADD 1 TO ws-error-count
           MOVE ws-error-count TO ws-error-code
           STRING "Error #" DELIMITED SIZE
                  ws-error-code DELIMITED SIZE
                  INTO ws-error-msg.
       END DECLARATIVES.
       main-section SECTION.
           DISPLAY ws-error-count
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_use_for_debugging() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-counter PIC 999 VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       debug-section SECTION.
           USE FOR DEBUGGING ON ALL PROCEDURES.
           DISPLAY "Debugging: " DEBUG-NAME.
       END DECLARATIVES.
       main-section SECTION.
           ADD 1 TO ws-counter
           DISPLAY ws-counter
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_use_for_debugging_specific() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-x PIC 99 VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       debug-x SECTION.
           USE FOR DEBUGGING ON ws-x.
           DISPLAY "ws-x changed to: " DEBUG-CONTENTS.
       END DECLARATIVES.
       main-para SECTION.
           MOVE 42 TO ws-x
           DISPLAY ws-x
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_procedural_flow_unaffected() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9(5) VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       err-sec SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           DISPLAY "error".
       END DECLARATIVES.
       main-section SECTION.
           PERFORM VARYING ws-result FROM 1 BY 1
               UNTIL ws-result > 5
               CONTINUE
           END-PERFORM
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_goback_in_handler() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-handled PIC X VALUE "N".
       PROCEDURE DIVISION.
       DECLARATIVES.
       fatal-error SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           MOVE "Y" TO ws-handled
           DISPLAY "Fatal error - handled".
       END DECLARATIVES.
       main-section SECTION.
           DISPLAY ws-handled
           STOP RUN.
"#,
    );
}

#[test]
fn declaratives_with_perform_in_handler() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-err-count PIC 99 VALUE 0.
       PROCEDURE DIVISION.
       DECLARATIVES.
       err-handler SECTION.
           USE AFTER STANDARD ERROR PROCEDURE ON ALL.
           PERFORM log-error.
       END DECLARATIVES.
       main-section SECTION.
           DISPLAY ws-err-count
           STOP RUN.
       log-error.
           ADD 1 TO ws-err-count
           DISPLAY "Error logged: " ws-err-count.
"#,
    );
}
