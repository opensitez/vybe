*> vybe-test: cobol/date_time_expanded/display_current_date_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    DISPLAY CURRENT-DATE.
    STOP RUN.

