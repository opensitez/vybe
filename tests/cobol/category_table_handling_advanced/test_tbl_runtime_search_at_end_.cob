*> vybe-test: cobol/category_table_handling_advanced/test_tbl_runtime_search_at_end_with_set
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION. 01 TBL. 05 EL OCCURS 3 TIMES ASCENDING KEY K INDEXED BY I. 10 K PIC 9. PROCEDURE DIVISION. MOVE 1 TO K(1) MOVE 2 TO K(2) MOVE 3 TO K(3) SEARCH ALL EL AT END DISPLAY 'NOT'. WHEN K(I) = 99 DISPLAY 'YES' END-SEARCH STOP RUN.

