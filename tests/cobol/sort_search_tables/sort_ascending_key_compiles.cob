*> vybe-test: cobol/sort_search_tables/sort_ascending_key_compiles
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC X(10).
01 K PIC 9(5).
PROCEDURE DIVISION.
    SORT F ON ASCENDING KEY K.
    STOP RUN.

