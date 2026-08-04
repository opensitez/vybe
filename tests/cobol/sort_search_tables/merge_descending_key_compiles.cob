*> vybe-test: cobol/sort_search_tables/merge_descending_key_compiles
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC X(10).
01 K PIC 9(5).
PROCEDURE DIVISION.
    MERGE F ON DESCENDING KEY K.
    STOP RUN.

