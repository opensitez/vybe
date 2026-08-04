*> vybe-test: cobol/accept_forms/accept_from_day_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DAY-OF-YEAR PIC 9(5).
PROCEDURE DIVISION.
    ACCEPT DAY-OF-YEAR FROM DAY.
    STOP RUN.

