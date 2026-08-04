*> vybe-test: cobol/category_copybook_advanced/test_replace_statement
*> origin: languages/cobol/tests/cobol/test_category_copybook_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). REPLACE ==A== BY ==B==. 01 A PIC X VALUE '1'. PROCEDURE DIVISION. DISPLAY B.
    MOVE SPACES TO WS-VYBE-L
    STRING B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

