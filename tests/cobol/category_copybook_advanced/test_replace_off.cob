*> vybe-test: cobol/category_copybook_advanced/test_replace_off
*> origin: languages/cobol/tests/cobol/test_category_copybook_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). REPLACE ==A== BY ==B==. REPLACE OFF. 01 A PIC X VALUE '1'. PROCEDURE DIVISION. DISPLAY A.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

