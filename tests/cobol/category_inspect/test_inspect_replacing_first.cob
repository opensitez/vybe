*> vybe-test: cobol/category_inspect/test_inspect_replacing_first
*> origin: languages/cobol/tests/cobol/test_category_inspect.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-REPL-FIRST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(10) VALUE "100-200-30".
       PROCEDURE DIVISION.
           INSPECT STR REPLACING FIRST "-" BY "X".
           DISPLAY STR.
    MOVE SPACES TO WS-VYBE-L
    STRING STR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "100X200-30"
        DISPLAY "FAIL: want [100X200-30] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

