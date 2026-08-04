*> vybe-test: cobol/category_table_handling_advanced/test_tbl_index_set
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 EL OCCURS 3 TIMES INDEXED BY I PIC 9 VALUE 5. PROCEDURE DIVISION. SET I TO 3. DISPLAY EL(I).
    MOVE SPACES TO WS-VYBE-L
    STRING EL(I) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

