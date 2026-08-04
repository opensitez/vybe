*> vybe-test: cobol/category_data_division_occurs/test_occurs_depending_max
*> origin: languages/cobol/tests/cobol/test_category_data_division_occurs.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 N PIC 9 VALUE 3. 01 TBL. 05 EL OCCURS 1 TO 3 TIMES DEPENDING ON N PIC 9 VALUE 1. PROCEDURE DIVISION. DISPLAY EL(3).
    MOVE SPACES TO WS-VYBE-L
    STRING EL(3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

