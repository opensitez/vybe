*> vybe-test: cobol/modules_multifile_calls/call_external_program_by_reference_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_modules_multifile_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-C.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5) VALUE 10.
PROCEDURE DIVISION.
    CALL "SUB-C" USING BY REFERENCE WS-A.
    STOP RUN.

