*> vybe-test: cobol/date_time_expanded/date_to_yyyymmdd_style_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(8).
PROCEDURE DIVISION.
    ACCEPT D FROM DATE YYYYMMDD.
    STOP RUN.

