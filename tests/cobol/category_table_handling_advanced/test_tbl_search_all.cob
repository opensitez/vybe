*> vybe-test: cobol/category_table_handling_advanced/test_tbl_search_all
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 EL OCCURS 3 TIMES ASCENDING KEY K INDEXED BY I. 10 K PIC 9. PROCEDURE DIVISION. MOVE 1 TO K(1). MOVE 2 TO K(2). MOVE 3 TO K(3). SEARCH ALL EL WHEN K(I) = 2 DISPLAY 'Y' END-SEARCH.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

