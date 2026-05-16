use super::helpers::{compile_ok, parse_ok, compile_ok_check};



fn p(data: &str, body: &str) -> String {
    format!("IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n{}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.", data, body)
}

// ═══════════════════════════════════════════════════════════
// CALL ASYNC — spawn thread
// ═══════════════════════════════════════════════════════════
#[test]
fn call_async_basic() {
    compile_ok(&p(
        "01 WS-HANDLE PIC X(10).",
        "    CALL \"WORKER\" ASYNC RETURNING WS-HANDLE."
    ));
}

#[test]
fn call_async_with_args() {
    compile_ok(&p(
        "01 WS-HANDLE PIC X(10).\n01 WS-INPUT PIC X(20) VALUE \"Data\".",
        "    CALL \"PROCESSOR\" ASYNC USING WS-INPUT RETURNING WS-HANDLE."
    ));
}

#[test]
fn call_async_no_handle() {
    compile_ok(&p(
        "01 WS-DATA PIC X(20) VALUE \"Fire and forget\".",
        "    CALL \"BACKGROUND\" ASYNC USING WS-DATA."
    ));
}

// ═══════════════════════════════════════════════════════════
// WAIT — join thread
// ═══════════════════════════════════════════════════════════
#[test]
fn wait_for_handle() {
    compile_ok(&p(
        "01 WS-HANDLE PIC X(10).",
        "    WAIT FOR WS-HANDLE."
    ));
}

#[test]
fn wait_basic() {
    compile_ok(&p(
        "01 WS-H PIC X(10).",
        "    WAIT WS-H."
    ));
}

// ═══════════════════════════════════════════════════════════
// CALL ASYNC + WAIT pattern (spawn then join)
// ═══════════════════════════════════════════════════════════
#[test]
fn async_spawn_and_join() {
    compile_ok(&p(
        "01 WS-HANDLE PIC X(10).\n01 WS-RESULT PIC X(20).",
        "    CALL \"COMPUTE-TASK\" ASYNC RETURNING WS-HANDLE.\n    DISPLAY \"Working...\".\n    WAIT FOR WS-HANDLE.\n    DISPLAY \"Done\"."
    ));
}

#[test]
fn multiple_async_calls() {
    compile_ok(&p(
        "01 WS-H1 PIC X(10).\n01 WS-H2 PIC X(10).\n01 WS-H3 PIC X(10).",
        "    CALL \"TASK1\" ASYNC RETURNING WS-H1.\n    CALL \"TASK2\" ASYNC RETURNING WS-H2.\n    CALL \"TASK3\" ASYNC RETURNING WS-H3.\n    WAIT FOR WS-H1.\n    WAIT FOR WS-H2.\n    WAIT FOR WS-H3.\n    DISPLAY \"All done\"."
    ));
}

// ═══════════════════════════════════════════════════════════
// RUN UNIT — separate execution thread
// ═══════════════════════════════════════════════════════════
#[test]
fn run_unit_basic() {
    compile_ok(&p(
        "",
        "    RUN-UNIT \"BATCH-PROCESS\"."
    ));
}

#[test]
fn run_unit_with_args() {
    compile_ok(&p(
        "01 WS-FILE PIC X(20) VALUE \"input.dat\".",
        "    RUN-UNIT \"PROCESSOR\" USING WS-FILE."
    ));
}

// ═══════════════════════════════════════════════════════════
// LOCK / UNLOCK — mutex/monitor
// ═══════════════════════════════════════════════════════════
#[test]
fn lock_unlock_basic() {
    compile_ok(&p(
        "01 WS-MUTEX PIC X(10).",
        "    LOCK WS-MUTEX.\n    DISPLAY \"Critical section\".\n    UNLOCK WS-MUTEX."
    ));
}

#[test]
fn lock_with_data_access() {
    compile_ok(&p(
        "01 WS-MUTEX PIC X(10).\n01 WS-COUNTER PIC 9(10) VALUE 0.",
        "    LOCK WS-MUTEX.\n    ADD 1 TO WS-COUNTER.\n    UNLOCK WS-MUTEX."
    ));
}

// ═══════════════════════════════════════════════════════════
// PERFORM ASYNC — fiber/coroutine
// ═══════════════════════════════════════════════════════════
#[test]
fn perform_async_paragraph() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. FIBER.
PROCEDURE DIVISION.
    PERFORM ASYNC WORKER-PARA.
    DISPLAY "Main continues".
    STOP RUN.
WORKER-PARA.
    DISPLAY "Worker running".
"#);
}

// ═══════════════════════════════════════════════════════════
// YIELD / SUSPEND — fiber control
// ═══════════════════════════════════════════════════════════
#[test]
fn yield_in_paragraph() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. YIELDTEST.
PROCEDURE DIVISION.
    DISPLAY "Start".
    STOP RUN.
GENERATOR-PARA.
    DISPLAY "Step 1".
    YIELD.
    DISPLAY "Step 2".
    YIELD.
    DISPLAY "Step 3".
"#);
}

#[test]
fn suspend_stmt() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. SUSPTEST.
PROCEDURE DIVISION.
    DISPLAY "Before suspend".
    SUSPEND.
    DISPLAY "After suspend".
    STOP RUN.
"#);
}

// ═══════════════════════════════════════════════════════════
// COMPLEX ASYNC PROGRAMS
// ═══════════════════════════════════════════════════════════
#[test]
fn parallel_file_processing() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PARALLEL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-HANDLE-1 PIC X(10).
01 WS-HANDLE-2 PIC X(10).
01 WS-FILE-1   PIC X(30) VALUE "customers.dat".
01 WS-FILE-2   PIC X(30) VALUE "orders.dat".
PROCEDURE DIVISION.
    DISPLAY "Starting parallel processing".
    CALL "PROCESS-FILE" ASYNC USING WS-FILE-1
        RETURNING WS-HANDLE-1.
    CALL "PROCESS-FILE" ASYNC USING WS-FILE-2
        RETURNING WS-HANDLE-2.
    DISPLAY "Both tasks launched".
    WAIT FOR WS-HANDLE-1.
    DISPLAY "File 1 done".
    WAIT FOR WS-HANDLE-2.
    DISPLAY "File 2 done".
    DISPLAY "All processing complete".
    STOP RUN.
"#);
}

#[test]
fn producer_consumer() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. PRODCONS.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-MUTEX   PIC X(10).
01 WS-BUFFER  PIC X(100).
01 WS-COUNT   PIC 9(5) VALUE 0.
01 WS-I       PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 10
        LOCK WS-MUTEX
        ADD 1 TO WS-COUNT
        UNLOCK WS-MUTEX
    END-PERFORM.
    DISPLAY "Produced " WS-COUNT " items".
    STOP RUN.
"#);
}

#[test]
fn async_batch_with_monitor() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ASYNCBATCH.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-LOCK    PIC X(10).
01 WS-TOTAL   PIC 9(10) VALUE 0.
01 WS-HANDLE  PIC X(10).
01 WS-I       PIC 9(5) VALUE 0.
PROCEDURE DIVISION.
    DISPLAY "Starting async batch".
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 5
        CALL "BATCH-WORKER" ASYNC USING WS-I
            RETURNING WS-HANDLE
        DISPLAY "Launched worker " WS-I
    END-PERFORM.
    DISPLAY "All workers launched".
    LOCK WS-LOCK.
    DISPLAY "Total: " WS-TOTAL.
    UNLOCK WS-LOCK.
    STOP RUN.
"#);
}

#[test]
fn fiber_generator_pattern() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. GENERATOR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VALUE PIC 9(10) VALUE 0.
PROCEDURE DIVISION.
    DISPLAY "Main program".
    STOP RUN.
NUMBER-GENERATOR.
    MOVE 1 TO WS-VALUE.
    DISPLAY WS-VALUE.
    YIELD.
    MOVE 2 TO WS-VALUE.
    DISPLAY WS-VALUE.
    YIELD.
    MOVE 3 TO WS-VALUE.
    DISPLAY WS-VALUE.
"#);
}

#[test]
fn concurrent_counter() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. CONCOUNT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-LOCK    PIC X(10).
01 WS-COUNTER PIC 9(10) VALUE 0.
01 WS-H1      PIC X(10).
01 WS-H2      PIC X(10).
PROCEDURE DIVISION.
    CALL "INCREMENT" ASYNC RETURNING WS-H1.
    CALL "INCREMENT" ASYNC RETURNING WS-H2.
    WAIT FOR WS-H1.
    WAIT FOR WS-H2.
    DISPLAY "Final counter: " WS-COUNTER.
    STOP RUN.
"#);
}

#[test]
fn async_with_error_handling() {
    compile_ok(r#"
IDENTIFICATION DIVISION.
PROGRAM-ID. ASYNCERR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-HANDLE PIC X(10).
01 WS-STATUS PIC 9(1) VALUE 0.
PROCEDURE DIVISION.
    CALL "RISKY-TASK" ASYNC RETURNING WS-HANDLE.
    WAIT FOR WS-HANDLE.
    IF WS-STATUS = 0
        DISPLAY "Task succeeded"
    ELSE
        DISPLAY "Task failed"
    END-IF.
    STOP RUN.
"#);
}
