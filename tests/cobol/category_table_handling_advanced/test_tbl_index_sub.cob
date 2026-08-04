*> vybe-test: cobol/category_table_handling_advanced/test_tbl_index_sub
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 EL OCCURS 5 TIMES INDEXED BY I PIC 9 VALUE 9. PROCEDURE DIVISION. SET I TO 4. DISPLAY EL(I - 2).
    MOVE SPACES TO WS-VYBE-L
    STRING EL(I DELIMITED SIZE - DELIMITED SIZE 2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "9"
        DISPLAY "FAIL: want [9] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

