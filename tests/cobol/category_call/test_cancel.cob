*> vybe-test: cobol/category_call/test_cancel
*> origin: languages/cobol/tests/cobol/test_category_call.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. CANCEL-TEST.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       PROCEDURE DIVISION.
           CANCEL "EXT-PROG".
           DISPLAY "CANCEL PARSED".
    MOVE SPACES TO WS-VYBE-L
    STRING "CANCEL PARSED" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "CANCEL PARSED"
        DISPLAY "FAIL: want [CANCEL PARSED] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

