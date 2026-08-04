*> vybe-test: cobol/accept_forms/accept_from_day_yyyyddd_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC 9(7).
PROCEDURE DIVISION.
    ACCEPT D FROM DAY YYYYDDD.
    STOP RUN.

