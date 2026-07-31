use super::helpers::{compile_ok, run_prints};

// ── ALTER statement (legacy COBOL-74/85) ─────────────────────
// ALTER changes the target of a GO TO so it jumps to a
// different paragraph on next execution.

#[test]
fn alter_basic() {
    let out = run_prints(
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
    assert_eq!(out, vec!["B"]);
}

#[test]
fn alter_conditional() {
    let out = run_prints(
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
    assert_eq!(out, vec!["SLOW"]);
}

#[test]
fn alter_multiple_targets() {
    let out = run_prints(
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
    assert_eq!(out, vec!["step 1", "step 3"]);
}

#[test]
fn alter_then_reset() {
    let out = run_prints(
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
    assert_eq!(out, vec!["2"]);
}

#[test]
fn alter_in_loop() {
    let out = run_prints(
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
    assert_eq!(out, vec!["3"]);
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
    let out = run_prints(
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
    assert_eq!(out, vec!["Y"]);
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

#[test]
fn alter_override_to_alt_path_runtime() {
    let out = run_prints(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-choice PIC X VALUE "Y".
       01 ws-result PIC X(10) VALUE SPACES.
       PROCEDURE DIVISION.
           IF ws-choice = "Y"
               ALTER entry-point TO PROCEED TO branch-a
           ELSE
               ALTER entry-point TO PROCEED TO branch-b
           END-IF
           GO TO entry-point
           STOP RUN.
       entry-point.
           GO TO branch-b.
       branch-a.
           MOVE "A" TO ws-result
           DISPLAY ws-result
           STOP RUN.
       branch-b.
           MOVE "B" TO ws-result
           DISPLAY ws-result
           STOP RUN.
"#,
    );
    assert_eq!(out, vec!["A"]);
}

#[test]
fn stop_run_runtime_after_display() {
    let out = run_prints(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           DISPLAY "start"
           STOP RUN.
           DISPLAY "after"
       "#,
    );
    assert_eq!(out, vec!["start"]);
}

#[test]
fn alter_sentence_target_is_honoured_compiles() {
    compile_ok(
        r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           ALTER entry-point TO PROCEED TO target-path
           GO TO entry-point
           STOP RUN.
       entry-point.
           GO TO default-path.
       default-path.
           DISPLAY "DEFAULT".
           STOP RUN.
       target-path.
           DISPLAY "TARGET".
           STOP RUN.
"#,
    );
}
