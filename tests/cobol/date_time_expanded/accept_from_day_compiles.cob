*> vybe-test: cobol/date_time_expanded/accept_from_day_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(5).
PROCEDURE DIVISION.
    ACCEPT D FROM DAY.
    STOP RUN.

