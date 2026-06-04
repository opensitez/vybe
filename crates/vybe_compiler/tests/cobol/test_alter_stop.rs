use super::helpers::compile_ok;

// ── ALTER statement (legacy COBOL-74/85) ─────────────────────
// ALTER changes the target of a GO TO so it jumps to a
// different paragraph on next execution.

#[test]
fn alter_basic() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-path PIC X VALUE "A".
       PROCEDURE DIVISION.
           ALTER jump-point TO PROCEED TO path-b
           GO TO jump-point
           STOP RUN.
       jump-point.
           GO TO path-a.
       path-a.
           MOVE "A" TO ws-path
           DISPLAY ws-path
           STOP RUN.
       path-b.
           MOVE "B" TO ws-path
           DISPLAY ws-path
           STOP RUN.
"#,
    );
}

#[test]
fn alter_conditional() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-mode PIC X VALUE "N".
       01 ws-result PIC X(10) VALUE SPACES.
       PROCEDURE DIVISION.
           IF ws-mode = "Y"
               ALTER dispatch TO PROCEED TO fast-path
           ELSE
               ALTER dispatch TO PROCEED TO slow-path
           END-IF
           GO TO dispatch
           STOP RUN.
       dispatch.
           GO TO slow-path.
       fast-path.
           MOVE "FAST" TO ws-result
           DISPLAY ws-result
           STOP RUN.
       slow-path.
           MOVE "SLOW" TO ws-result
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn alter_multiple_targets() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-step PIC 9 VALUE 1.
       PROCEDURE DIVISION.
           ALTER step-1-exit TO PROCEED TO step-2
           ALTER step-2-exit TO PROCEED TO step-3
           GO TO step-1
           STOP RUN.
       step-1.
           DISPLAY "step 1"
           GO TO step-1-exit.
       step-1-exit.
           GO TO step-3.
       step-2.
           DISPLAY "step 2"
           GO TO step-2-exit.
       step-2-exit.
           GO TO step-3.
       step-3.
           DISPLAY "step 3"
           STOP RUN.
"#,
    );
}

#[test]
fn alter_then_reset() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-count PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           ALTER router TO PROCEED TO path-x
           GO TO router
           STOP RUN.
       router.
           GO TO path-y.
       path-x.
           MOVE 1 TO ws-count
           ALTER router TO PROCEED TO path-y
           DISPLAY ws-count
           STOP RUN.
       path-y.
           MOVE 2 TO ws-count
           DISPLAY ws-count
           STOP RUN.
"#,
    );
}

#[test]
fn alter_in_loop() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-iter PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           GO TO loop-body.
       loop-body.
           ADD 1 TO ws-iter
           IF ws-iter >= 3
               ALTER loop-exit TO PROCEED TO done
           END-IF
           IF ws-iter < 3
               GO TO loop-body
           ELSE
               GO TO loop-exit
           END-IF.
       loop-exit.
           GO TO loop-body.
       done.
           DISPLAY ws-iter
           STOP RUN.
"#,
    );
}

// ── STOP literal ─────────────────────────────────────────────
// STOP "message" pauses execution and displays the literal.
// Legacy feature; STOP RUN terminates the program.

#[test]
fn stop_literal_string() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-flag PIC X VALUE "N".
       PROCEDURE DIVISION.
           MOVE "Y" TO ws-flag
           DISPLAY ws-flag
           STOP "Press Enter to continue"
           STOP RUN.
"#,
    );
}

#[test]
fn stop_literal_numeric() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           DISPLAY "before pause"
           STOP 0
           DISPLAY "after pause"
           STOP RUN.
"#,
    );
}

#[test]
fn stop_literal_in_conditional() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-debug PIC X VALUE "Y".
       PROCEDURE DIVISION.
           IF ws-debug = "Y"
               STOP "Debug checkpoint reached"
           END-IF
           DISPLAY "continuing"
           STOP RUN.
"#,
    );
}

#[test]
fn stop_literal_after_compute() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9(5) VALUE 0.
       PROCEDURE DIVISION.
           COMPUTE ws-result = 12345 * 2
           STOP "Check result"
           DISPLAY ws-result
           STOP RUN.
"#,
    );
}

#[test]
fn stop_run_from_nested_perform() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-found PIC X VALUE "N".
       01 ws-i     PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING ws-i FROM 1 BY 1 UNTIL ws-i > 9
               IF ws-i = 5
                   MOVE "Y" TO ws-found
               END-IF
           END-PERFORM
           DISPLAY ws-found
           STOP RUN.
"#,
    );
}

#[test]
fn stop_literal_with_spaces() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           DISPLAY "starting"
           STOP "   "
           DISPLAY "done"
           STOP RUN.
"#,
    );
}

#[test]
fn stop_literal_long_message() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           STOP "This is a longer pause message for the operator"
           STOP RUN.
"#,
    );
}
