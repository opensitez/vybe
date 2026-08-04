*> vybe-test: cobol/category_search_advanced/test_search_linear_no_match
*> origin: languages/cobol/tests/cobol/test_category_search_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 TBL. 05 EL OCCURS 3 TIMES INDEXED BY I PIC X. PROCEDURE DIVISION. MOVE 'A' TO EL(1). MOVE 'B' TO EL(2). MOVE 'C' TO EL(3). SET I TO 1. SEARCH EL AT END DISPLAY 'N' WHEN EL(I) = 'Z' DISPLAY 'Y' END-SEARCH. STOP RUN.

