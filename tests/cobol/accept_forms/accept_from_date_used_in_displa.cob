*> vybe-test: cobol/accept_forms/accept_from_date_used_in_display
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 TODAY PIC 9(6).
PROCEDURE DIVISION.
    ACCEPT TODAY FROM DATE.
    DISPLAY TODAY.
    STOP RUN.

