*> vybe-test: cobol/date_time_expanded/accept_from_day_of_week_compiles
*> origin: languages/cobol/tests/cobol/test_date_time_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 W PIC X(1).
PROCEDURE DIVISION.
    ACCEPT W FROM DAY-OF-WEEK.
    STOP RUN.

