*> vybe-test: cobol/call_statement/test_call_no_params
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

    CALL "SUBPROG".
    STOP RUN.

