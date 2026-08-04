*> vybe-test: cobol/accept_forms/stop_run_with_return_code
*> origin: languages/cobol/tests/cobol/test_accept_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RETURN-CODE PIC 9(4).
PROCEDURE DIVISION.
    MOVE 0 TO RETURN-CODE.
    STOP RUN.

