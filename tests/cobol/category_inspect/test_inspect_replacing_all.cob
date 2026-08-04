*> vybe-test: cobol/category_inspect/test_inspect_replacing_all
*> origin: languages/cobol/tests/cobol/test_category_inspect.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-REPL-ALL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(10) VALUE "A-B-C-D-".
       PROCEDURE DIVISION.
           INSPECT STR REPLACING ALL "-" BY " ".
           DISPLAY STR.
    MOVE SPACES TO WS-VYBE-L
    STRING STR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A B C D   "
        DISPLAY "FAIL: want [A B C D   ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

