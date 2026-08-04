*> vybe-test: cobol/category_inspect/test_inspect_converting
*> origin: languages/cobol/tests/cobol/test_category_inspect.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. INSPECT-CONV.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 STR PIC X(5) VALUE "HELLO".
       PROCEDURE DIVISION.
           INSPECT STR CONVERTING "EOL" TO "e01".
           DISPLAY STR.
    MOVE SPACES TO WS-VYBE-L
    STRING STR DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "He110"
        DISPLAY "FAIL: want [He110] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

