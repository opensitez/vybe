*> vybe-test: cobol/accept_forms/accept_from_date_yyyymmdd_compiles
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TODAY PIC 9(8).
PROCEDURE DIVISION.
    ACCEPT TODAY FROM DATE YYYYMMDD.
    STOP RUN.

