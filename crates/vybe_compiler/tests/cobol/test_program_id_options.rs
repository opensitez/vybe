use super::helpers::compile_ok;

// ── PROGRAM-ID INITIAL ────────────────────────────────────────
// INITIAL: each CALL reinitializes all working storage and
// internal files as if called for the first time.

#[test] fn program_id_initial_basic() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. test INITIAL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-counter PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           ADD 1 TO ws-counter
           DISPLAY ws-counter
           STOP RUN.
"#);
}

#[test] fn program_id_initial_with_call() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. main-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-result PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           CALL "fresh-sub" USING ws-result
           DISPLAY ws-result
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. fresh-sub INITIAL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-internal PIC 99 VALUE 0.
       LINKAGE SECTION.
       01 ls-result PIC 99.
       PROCEDURE DIVISION USING ls-result.
           ADD 1 TO ws-internal
           MOVE ws-internal TO ls-result
           GOBACK.
"#);
}

#[test] fn program_id_initial_reset_behavior() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. reset-test INITIAL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 99 VALUE 10.
       01 ws-txt PIC X(10) VALUE "original".
       PROCEDURE DIVISION.
           ADD 5 TO ws-val
           MOVE "changed" TO ws-txt
           DISPLAY ws-val
           DISPLAY ws-txt
           STOP RUN.
"#);
}

// ── PROGRAM-ID COMMON ─────────────────────────────────────────
// COMMON: nested program visible to ALL programs in the compile
// unit, not just the directly containing program.

#[test] fn program_id_common_nested() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. outer-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-shared PIC 99 VALUE 0.
       PROCEDURE DIVISION.
           CALL "common-util"
           DISPLAY ws-shared
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. common-util IS COMMON.
       DATA DIVISION.
       PROCEDURE DIVISION.
           DISPLAY "common utility called"
           GOBACK.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. inner-prog.
       PROCEDURE DIVISION.
           CALL "common-util"
           GOBACK.
       END PROGRAM inner-prog.

       END PROGRAM common-util.
       END PROGRAM outer-prog.
"#);
}

#[test] fn program_id_common_basic() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. host-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-count PIC 9 VALUE 0.
       PROCEDURE DIVISION.
           ADD 1 TO ws-count
           DISPLAY ws-count
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. shared-sub IS COMMON.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-msg PIC X(20) VALUE "shared utility".
       PROCEDURE DIVISION.
           DISPLAY ws-msg
           GOBACK.

       END PROGRAM shared-sub.
       END PROGRAM host-prog.
"#);
}

// ── PROGRAM-ID RECURSIVE ──────────────────────────────────────
// RECURSIVE: allows a program to call itself (directly or
// indirectly). Each activation has its own LOCAL-STORAGE.

#[test] fn program_id_recursive_basic() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. factorial IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-sub-result PIC 9(10) VALUE 0.
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC 9(10).
       PROCEDURE DIVISION USING lk-n lk-result.
           IF lk-n <= 1
               MOVE 1 TO lk-result
           ELSE
               SUBTRACT 1 FROM lk-n
               CALL "factorial" USING lk-n ls-sub-result
               ADD 1 TO lk-n
               MULTIPLY lk-n BY ls-sub-result GIVING lk-result
           END-IF
           GOBACK.
"#);
}

#[test] fn program_id_recursive_fibonacci() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. fib IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-a PIC 9(10) VALUE 0.
       01 ls-b PIC 9(10) VALUE 0.
       01 ls-n-minus-1 PIC 9(5) VALUE 0.
       01 ls-n-minus-2 PIC 9(5) VALUE 0.
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC 9(10).
       PROCEDURE DIVISION USING lk-n lk-result.
           EVALUATE TRUE
               WHEN lk-n = 0 MOVE 0 TO lk-result
               WHEN lk-n = 1 MOVE 1 TO lk-result
               WHEN OTHER
                   SUBTRACT 1 FROM lk-n GIVING ls-n-minus-1
                   SUBTRACT 2 FROM lk-n GIVING ls-n-minus-2
                   CALL "fib" USING ls-n-minus-1 ls-a
                   CALL "fib" USING ls-n-minus-2 ls-b
                   ADD ls-a ls-b GIVING lk-result
           END-EVALUATE
           GOBACK.
"#);
}

#[test] fn program_id_recursive_with_local_storage() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. rec-sum IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-partial PIC 9(8) VALUE 0.
       01 ls-n-dec   PIC 9(5) VALUE 0.
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC 9(8).
       PROCEDURE DIVISION USING lk-n lk-result.
           IF lk-n = 0
               MOVE 0 TO lk-result
           ELSE
               SUBTRACT 1 FROM lk-n GIVING ls-n-dec
               CALL "rec-sum" USING ls-n-dec ls-partial
               ADD lk-n ls-partial GIVING lk-result
           END-IF
           GOBACK.
"#);
}

#[test] fn program_id_recursive_countdown() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. countdown IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-next PIC 9(3) VALUE 0.
       LINKAGE SECTION.
       01 lk-n PIC 9(3).
       PROCEDURE DIVISION USING lk-n.
           DISPLAY lk-n
           IF lk-n > 0
               SUBTRACT 1 FROM lk-n GIVING ls-next
               CALL "countdown" USING ls-next
           END-IF
           GOBACK.
"#);
}

// ── INITIAL + RECURSIVE combination ──────────────────────────

#[test] fn program_id_initial_and_recursive() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. init-rec IS INITIAL RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-depth PIC 9 VALUE 0.
       LINKAGE SECTION.
       01 lk-max PIC 9.
       PROCEDURE DIVISION USING lk-max.
           IF ls-depth < lk-max
               ADD 1 TO ls-depth
               CALL "init-rec" USING lk-max
           ELSE
               DISPLAY ls-depth
           END-IF
           GOBACK.
"#);
}

// ── PROGRAM-ID with PROGRAM keyword explicit ──────────────────

#[test] fn program_id_with_end_program() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. bounded-prog.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-val PIC 99 VALUE 42.
       PROCEDURE DIVISION.
           DISPLAY ws-val
           STOP RUN.
       END PROGRAM bounded-prog.
"#);
}

#[test] fn program_id_nested_end_program() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. outer.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-x PIC 9 VALUE 1.
       PROCEDURE DIVISION.
           DISPLAY ws-x
           STOP RUN.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. inner.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 ws-y PIC 9 VALUE 2.
       PROCEDURE DIVISION.
           DISPLAY ws-y
           GOBACK.
       END PROGRAM inner.

       END PROGRAM outer.
"#);
}

// ── RECURSIVE with mutual recursion ──────────────────────────

#[test] fn mutual_recursion_even_odd() {
    compile_ok(r#"
       IDENTIFICATION DIVISION.
       PROGRAM-ID. is-even IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-n-minus-1 PIC 9(5) VALUE 0.
       01 ls-sub-result PIC X VALUE "?".
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC X.
       PROCEDURE DIVISION USING lk-n lk-result.
           IF lk-n = 0
               MOVE "Y" TO lk-result
           ELSE
               SUBTRACT 1 FROM lk-n GIVING ls-n-minus-1
               CALL "is-odd" USING ls-n-minus-1 lk-result
           END-IF
           GOBACK.

       IDENTIFICATION DIVISION.
       PROGRAM-ID. is-odd IS RECURSIVE.
       DATA DIVISION.
       LOCAL-STORAGE SECTION.
       01 ls-n-minus-1 PIC 9(5) VALUE 0.
       LINKAGE SECTION.
       01 lk-n      PIC 9(5).
       01 lk-result PIC X.
       PROCEDURE DIVISION USING lk-n lk-result.
           IF lk-n = 0
               MOVE "N" TO lk-result
           ELSE
               SUBTRACT 1 FROM lk-n GIVING ls-n-minus-1
               CALL "is-even" USING ls-n-minus-1 lk-result
           END-IF
           GOBACK.
"#);
}
