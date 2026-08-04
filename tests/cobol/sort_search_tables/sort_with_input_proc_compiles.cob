*> vybe-test: cobol/sort_search_tables/sort_with_input_proc_compiles
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC X(10).
01 K PIC 9(5).
PROCEDURE DIVISION.
    SORT F ON ASCENDING KEY K INPUT PROCEDURE IS P1.
    STOP RUN.

