*> vybe-test: cobol/accept_forms/accept_from_day_of_week_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DOW PIC 9.
PROCEDURE DIVISION.
    ACCEPT DOW FROM DAY-OF-WEEK.
    STOP RUN.

