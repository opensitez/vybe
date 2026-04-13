use super::helpers::compile_ok;

// ═══════════════════════════════════════════════════════════
// COBOL 2023: Async, threading, and concurrency
// Tests for modern async patterns beyond test_async_threads.rs
// which covers CALL ASYNC, WAIT, RUN-UNIT, LOCK/UNLOCK,
// PERFORM ASYNC, YIELD, SUSPEND.
// ═══════════════════════════════════════════════════════════

#[test]
fn call_async_returning() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-HANDLE PIC X(20).
01 WS-RESULT PIC X(50).
PROCEDURE DIVISION.
    CALL "PROCESS-DATA" ASYNC RETURNING WS-HANDLE.
    WAIT FOR WS-HANDLE.
    DISPLAY "Done".
    STOP RUN.
"#);
}

#[test]
fn multiple_async_calls_sync() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-H1 PIC X(20).
01 WS-H2 PIC X(20).
01 WS-H3 PIC X(20).
PROCEDURE DIVISION.
    CALL "TASK-A" ASYNC RETURNING WS-H1.
    CALL "TASK-B" ASYNC RETURNING WS-H2.
    CALL "TASK-C" ASYNC RETURNING WS-H3.
    WAIT FOR WS-H1.
    WAIT FOR WS-H2.
    WAIT FOR WS-H3.
    DISPLAY "All complete".
    STOP RUN.
"#);
}

#[test]
fn lock_unlock_resource() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNTER PIC 9(10) VALUE 0.
01 WS-MUTEX PIC X(20).
PROCEDURE DIVISION.
    LOCK WS-MUTEX.
    ADD 1 TO WS-COUNTER.
    UNLOCK WS-MUTEX.
    DISPLAY WS-COUNTER.
    STOP RUN.
"#);
}

#[test]
fn perform_async_fiber() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM ASYNC WORKER-TASK.
    DISPLAY "Main continues".
    STOP RUN.
WORKER-TASK.
    DISPLAY "Worker running".
"#);
}

#[test]
fn yield_in_loop() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-I PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    PERFORM 5 TIMES
        ADD 1 TO WS-I
        DISPLAY WS-I
        YIELD
    END-PERFORM.
    STOP RUN.
"#);
}

#[test]
fn suspend_resume() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    DISPLAY "Before suspend".
    SUSPEND.
    DISPLAY "After resume".
    STOP RUN.
"#);
}

#[test]
fn call_by_reference() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARG PIC 9(5) VALUE 100.
PROCEDURE DIVISION.
    CALL "UPDATE-VALUE" USING BY REFERENCE WS-ARG.
    DISPLAY WS-ARG.
    STOP RUN.
"#);
}

#[test]
fn call_by_content() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARG PIC 9(5) VALUE 100.
PROCEDURE DIVISION.
    CALL "READ-VALUE" USING BY CONTENT WS-ARG.
    DISPLAY WS-ARG.
    STOP RUN.
"#);
}

#[test]
fn call_by_value() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARG PIC 9(5) VALUE 100.
PROCEDURE DIVISION.
    CALL "COPY-VALUE" USING BY VALUE WS-ARG.
    DISPLAY WS-ARG.
    STOP RUN.
"#);
}

#[test]
fn call_on_exception() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC X(20).
PROCEDURE DIVISION.
    CALL "MAYBE-FAIL" USING WS-RESULT
        ON EXCEPTION DISPLAY "Call failed"
        NOT ON EXCEPTION DISPLAY "Call succeeded"
    END-CALL.
    STOP RUN.
"#);
}
