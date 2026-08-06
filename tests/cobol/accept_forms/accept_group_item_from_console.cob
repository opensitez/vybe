*> vybe-test: cobol/accept_forms/accept_group_item_from_console
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RESPONSE.
   05 CODE PIC X.
   05 DETAIL PIC X(10).
PROCEDURE DIVISION.
    ACCEPT RESPONSE FROM CONSOLE.
    STOP RUN.

