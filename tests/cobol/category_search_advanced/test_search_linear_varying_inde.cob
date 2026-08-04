*> vybe-test: cobol/category_search_advanced/test_search_linear_varying_index
*> origin: languages/cobol/tests/cobol/test_category_search_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 EL OCCURS 3 TIMES INDEXED BY I J PIC X. PROCEDURE DIVISION. MOVE 'A' TO EL(1). MOVE 'B' TO EL(2). MOVE 'C' TO EL(3). SET I TO 1. SET J TO 1. SEARCH EL VARYING J WHEN EL(I) = 'B' DISPLAY J END-SEARCH.
    MOVE SPACES TO WS-VYBE-L
    STRING J DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

