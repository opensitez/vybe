*> vybe-test: cobol/category_data_division_occurs/test_occurs_indexed_search_lookup
*> origin: languages/cobol/tests/cobol/test_category_data_division_occurs.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 E OCCURS 4 TIMES INDEXED BY I PIC 9(2) VALUE 0. PROCEDURE DIVISION. MOVE 10 TO E(1) MOVE 20 TO E(2) MOVE 30 TO E(3) SET I TO 2 DISPLAY E(I) STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING E(I) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "20"
        DISPLAY "FAIL: want [20] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

