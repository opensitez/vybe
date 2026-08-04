*> vybe-test: cobol/accept_forms/accept_multiple_fields_from_date
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D1 PIC 9(6).
01 D2 PIC 9(8).
PROCEDURE DIVISION.
    ACCEPT D1 FROM DATE.
    ACCEPT D2 FROM DATE YYYYMMDD.
    STOP RUN.

