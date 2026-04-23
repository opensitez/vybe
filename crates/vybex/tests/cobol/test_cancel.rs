use super::helpers::compile_ok;

// ── CANCEL statement ──────────────────────────────────────────
// CANCEL releases the storage and initialization state of a
// called subprogram so that the next CALL re-initializes it.

#[test] fn cancel_basic() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           CALL "utility-sub"
           CANCEL "utility-sub"
           DISPLAY "cancelled"
           STOP RUN.
"#);
}

#[test] fn cancel_then_recall() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 9(5) VALUE 0.
       PROCEDURE DIVISION.
           CALL "counter-sub" USING ws-result
           DISPLAY ws-result
           CANCEL "counter-sub"
           MOVE 0 TO ws-result
           CALL "counter-sub" USING ws-result
           DISPLAY ws-result
           STOP RUN.
"#);
}

#[test] fn cancel_by_identifier() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-prog-name PIC X(20) VALUE "helper-module".
       PROCEDURE DIVISION.
           CALL ws-prog-name
           CANCEL ws-prog-name
           DISPLAY "done"
           STOP RUN.
"#);
}

#[test] fn cancel_multiple_programs() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           CALL "module-a"
           CALL "module-b"
           CALL "module-c"
           CANCEL "module-a"
           CANCEL "module-b"
           CANCEL "module-c"
           DISPLAY "all cancelled"
           STOP RUN.
"#);
}

#[test] fn cancel_after_exception() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-status PIC X VALUE "N".
       PROCEDURE DIVISION.
           CALL "risky-module"
               ON EXCEPTION
                   MOVE "E" TO ws-status
           END-CALL
           CANCEL "risky-module"
           DISPLAY ws-status
           STOP RUN.
"#);
}

#[test] fn cancel_uncalled_program() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           CANCEL "never-called-prog"
           DISPLAY "no error expected"
           STOP RUN.
"#);
}

#[test] fn cancel_in_conditional() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-flag PIC X VALUE "Y".
       PROCEDURE DIVISION.
           CALL "temp-module"
           IF ws-flag = "Y"
               CANCEL "temp-module"
           END-IF
           DISPLAY ws-flag
           STOP RUN.
"#);
}

#[test] fn cancel_in_loop() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-i PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           PERFORM VARYING ws-i FROM 1 BY 1 UNTIL ws-i > 3
               CALL "batch-proc" USING ws-i
               CANCEL "batch-proc"
           END-PERFORM
           DISPLAY "loop done"
           STOP RUN.
"#);
}

#[test] fn cancel_resets_initial_program() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-count PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           CALL "init-counter" USING ws-count
           DISPLAY ws-count
           CANCEL "init-counter"
           MOVE 0 TO ws-count
           CALL "init-counter" USING ws-count
           DISPLAY ws-count
           STOP RUN.
"#);
}

#[test] fn cancel_with_call_on_exception() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-loaded PIC X VALUE "N".
       PROCEDURE DIVISION.
           CALL "plugin-module"
               ON EXCEPTION
                   DISPLAY "load failed"
                   GO TO end-prog
               NOT ON EXCEPTION
                   MOVE "Y" TO ws-loaded
           END-CALL
           IF ws-loaded = "Y"
               CANCEL "plugin-module"
           END-IF
       end-prog.
           STOP RUN.
"#);
}

#[test] fn cancel_variable_list() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-mods.
           05 ws-mod-1 PIC X(20) VALUE "module-x".
           05 ws-mod-2 PIC X(20) VALUE "module-y".
       PROCEDURE DIVISION.
           CALL ws-mod-1
           CALL ws-mod-2
           CANCEL ws-mod-1
           CANCEL ws-mod-2
           DISPLAY "freed"
           STOP RUN.
"#);
}

#[test] fn cancel_nested_call_chain() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       PROCEDURE DIVISION.
           CALL "level-1"
           CANCEL "level-1"
           CALL "level-2"
           CANCEL "level-2"
           CALL "level-3"
           CANCEL "level-3"
           STOP RUN.
"#);
}

#[test] fn cancel_preserves_caller_state() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-local PIC 9(5) VALUE 42.
       PROCEDURE DIVISION.
           CALL "sub-prog"
           CANCEL "sub-prog"
           DISPLAY ws-local
           STOP RUN.
"#);
}

#[test] fn cancel_dynamic_dispatch() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-handler PIC X(30).
       01 ws-mode    PIC X(10) VALUE "fast".
       PROCEDURE DIVISION.
           IF ws-mode = "fast"
               MOVE "fast-handler" TO ws-handler
           ELSE
               MOVE "slow-handler" TO ws-handler
           END-IF
           CALL ws-handler
           CANCEL ws-handler
           DISPLAY "dispatched"
           STOP RUN.
"#);
}
