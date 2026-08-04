*> vybe-test: cobol/category_table_handling_advanced/test_tbl_index_compare
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 EL OCCURS 3 TIMES INDEXED BY I J PIC 9. PROCEDURE DIVISION. SET I TO 2. SET J TO 2. IF I = J DISPLAY 'Y' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

