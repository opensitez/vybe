*> vybe-test: cobol/accept_forms/accept_numeric_from_date
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 YEAR PIC 9(4).
01 FULL-DATE PIC 9(8).
PROCEDURE DIVISION.
    ACCEPT FULL-DATE FROM DATE YYYYMMDD.
    MOVE FULL-DATE(1:4) TO YEAR.
    STOP RUN.

