*> vybe-test: cobol/modules_multifile_calls/call_external_program_no_args_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_modules_multifile_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-A.
PROCEDURE DIVISION.
    CALL "SUB-A".
    STOP RUN.

