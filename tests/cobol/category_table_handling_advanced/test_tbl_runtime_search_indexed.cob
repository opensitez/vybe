*> vybe-test: cobol/category_table_handling_advanced/test_tbl_runtime_search_indexed_match
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 EL OCCURS 4 TIMES INDEXED BY I. 10 K PIC 9. PROCEDURE DIVISION. MOVE 1 TO K(1) MOVE 3 TO K(2) MOVE 5 TO K(3) MOVE 7 TO K(4) SEARCH EL WHEN K(I) = 5 DISPLAY 'FOUND' END-SEARCH STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING 'FOUND' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FOUND"
        DISPLAY "FAIL: want [FOUND] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

