*> vybe-test: cobol/category_table_handling_advanced/test_tbl_occurs_depending_on
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 N PIC 9 VALUE 2. 01 TBL. 05 EL OCCURS 1 TO 3 TIMES DEPENDING ON N PIC 9 VALUE 4. PROCEDURE DIVISION. DISPLAY EL(N).
    MOVE SPACES TO WS-VYBE-L
    STRING EL(N) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "4"
        DISPLAY "FAIL: want [4] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

