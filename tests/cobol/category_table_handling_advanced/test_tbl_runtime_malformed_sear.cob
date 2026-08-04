*> vybe-test: cobol/category_table_handling_advanced/test_tbl_runtime_malformed_search_guard
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 TBL. 05 EL OCCURS 2 TIMES INDEXED BY I PIC 9 VALUE 1. PROCEDURE DIVISION. MOVE 1 TO EL(1) MOVE 9 TO EL(2) IF EL(1) > EL(2) DISPLAY 'NEG' ELSE DISPLAY 'OK' END-IF STOP RUN.
    MOVE SPACES TO WS-VYBE-L
    STRING 'NEG' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK"
        DISPLAY "FAIL: want [OK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

