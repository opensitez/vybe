*> vybe-test: cobol/date_time_expanded/date_compare_branch_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(8).
PROCEDURE DIVISION.
    ACCEPT D FROM DATE.
    IF D > "20250101" DISPLAY "NEW" ELSE DISPLAY "OLD" END-IF.
    STOP RUN.

