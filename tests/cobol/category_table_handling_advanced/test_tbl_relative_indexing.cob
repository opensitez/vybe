*> vybe-test: cobol/category_table_handling_advanced/test_tbl_relative_indexing
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 N PIC 9 VALUE 2. 01 TBL. 05 EL OCCURS 5 TIMES PIC 9 VALUE 1. PROCEDURE DIVISION. DISPLAY EL(N + 1).
    MOVE SPACES TO WS-VYBE-L
    STRING EL(N DELIMITED SIZE + DELIMITED SIZE 1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

