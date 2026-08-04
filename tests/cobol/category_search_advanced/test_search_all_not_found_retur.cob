*> vybe-test: cobol/category_search_advanced/test_search_all_not_found_returns_end
*> origin: languages/cobol/tests/cobol/test_category_search_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 TBL. 05 EL OCCURS 2 TIMES ASCENDING K INDEXED BY I. 10 K PIC 9. PROCEDURE DIVISION. MOVE 1 TO K(1). MOVE 2 TO K(2). SEARCH ALL EL AT END DISPLAY 'NOT' WHEN K(I) = 9 DISPLAY 'YES' END-SEARCH DISPLAY 'DONE'. STOP RUN.

