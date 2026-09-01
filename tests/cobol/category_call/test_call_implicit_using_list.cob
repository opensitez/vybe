*> vybe-test: cobol/category_call/test_call_implicit_using_list
*> vybe-test-mode: compile
*> `CALL "…"` names a program that exists nowhere in this run unit, and the
*> source carries no ON EXCEPTION phrase to catch it. cobc compiles this and
*> then aborts — `libcob: error: module not found` — so "runs and exits 0" is
*> not a property it has under any COBOL, and no compiler change can give it
*> one. Asserting that it COMPILES is the strongest true claim available.
*> origin: languages/cobol/tests/cobol/test_category_call.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. CALL-LIST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 A PIC 9 VALUE 1.
       01 B PIC 9 VALUE 2.
       PROCEDURE DIVISION.
           CALL "SUB-LIST" USING A B.
           DISPLAY "OK".
    MOVE SPACES TO WS-VYBE-L
    STRING "OK" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

