*> vybe-test: cobol/delegate_pointer_binding/callback_style_call_with_pointer_args_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_delegate_pointer_binding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CALLBACK USAGE IS PROCEDURE-POINTER.
01 WS-ARG PIC X(10) VALUE "PAYLOAD".
PROCEDURE DIVISION.
    CALL "INVOKE-CALLBACK" USING WS-CALLBACK WS-ARG.
    STOP RUN.

