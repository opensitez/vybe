*> vybe-test: cobol/date_time_expanded/date_to_yyddd_style_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(5).
PROCEDURE DIVISION.
    ACCEPT D FROM DAY YYYYDDD.
    STOP RUN.

